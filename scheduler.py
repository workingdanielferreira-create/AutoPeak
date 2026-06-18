"""
scheduler.py
============
Pure scheduling engine. No Outlook / no GUI dependencies, so it can be
unit-tested anywhere. Given a config and a list of existing busy intervals,
it returns a list of proposed appointments.

Schedule types
--------------
  1 = on/before the 10th of the month
  2 = between the 11th and the 20th
  3 = between the 21st and the last day of the month
  4 = once per "bucket", spread across up to 4 weeks (weeks start Sunday)
  5 = anytime in the month (flexible)  <-- used for "Coach SKEP Audits"
"""

from __future__ import annotations
from dataclasses import dataclass, field
from datetime import datetime, date, time, timedelta
import calendar


# --------------------------------------------------------------------------- #
# Data structures
# --------------------------------------------------------------------------- #
@dataclass
class Meeting:
    name: str
    category: str            # "FS" or "DS"
    schedule_type: int       # 1..5
    duration_minutes: int
    count: int               # final, resolved number of instances
    invite_all_staff: bool = False   # FS meetings can invite the whole team


@dataclass
class Proposal:
    name: str
    start: datetime
    end: datetime
    category: str
    schedule_type: int
    attendees: list[str] = field(default_factory=list)

    def as_row(self):
        if not self.attendees:
            who = ""
        elif len(self.attendees) == 1:
            who = self.attendees[0]
        else:
            who = f"{len(self.attendees)} invitees"
        return (
            self.start.strftime("%a %d %b"),
            self.start.strftime("%H:%M"),
            self.end.strftime("%H:%M"),
            f"{int((self.end - self.start).total_seconds() // 60)}m",
            self.name,
            self.category,
            self.schedule_type,
            who,
        )


@dataclass
class Interval:
    start: datetime
    end: datetime


# --------------------------------------------------------------------------- #
# Config resolution
# --------------------------------------------------------------------------- #
def staff_emails(cfg: dict) -> list[str]:
    """The list of direct-report email addresses. Accepts either a list of
    plain email strings or a list of {"email": ...} objects."""
    raw = cfg.get("staff", [])
    out = []
    for item in raw:
        email = item.get("email") if isinstance(item, dict) else item
        email = (email or "").strip()
        if email:
            out.append(email)
    return out


def resolve_meetings(cfg: dict) -> list[Meeting]:
    """Turn the config's meeting list into concrete Meeting objects,
    resolving dynamic (DS) counts from the number of staff."""
    staff = staff_emails(cfg)
    out = []
    for m in cfg["meetings"]:
        if m["category"] == "DS":
            count = int(m.get("per_staff", 1)) * len(staff)
        else:
            count = int(m.get("count", 1))
        if count <= 0:
            continue
        out.append(
            Meeting(
                name=m["name"],
                category=m["category"],
                schedule_type=int(m["schedule_type"]),
                duration_minutes=int(m["duration_minutes"]),
                count=count,
                invite_all_staff=bool(m.get("invite_all_staff", False)),
            )
        )
    return out


# --------------------------------------------------------------------------- #
# Calendar helpers
# --------------------------------------------------------------------------- #
def month_days(year: int, month: int) -> list[date]:
    last = calendar.monthrange(year, month)[1]
    return [date(year, month, d) for d in range(1, last + 1)]


def week_start_sunday(d: date) -> date:
    """Return the Sunday that begins the week containing d."""
    # Mon=0 .. Sun=6  ->  days since the most recent Sunday
    return d - timedelta(days=(d.weekday() + 1) % 7)


def candidate_days(meeting: Meeting, year: int, month: int,
                   work_days: set[int]) -> list[date]:
    """Workdays that are valid for this meeting's schedule type."""
    days = [d for d in month_days(year, month) if d.weekday() in work_days]
    t = meeting.schedule_type
    if t == 1:
        return [d for d in days if d.day <= 10]
    if t == 2:
        return [d for d in days if 11 <= d.day <= 20]
    if t == 3:
        return [d for d in days if d.day >= 21]
    # types 4 and 5 use the whole month
    return days


def distribute(total: int, buckets: int) -> list[int]:
    """Split `total` into `buckets` as evenly as possible."""
    if buckets <= 0:
        return []
    base, extra = divmod(total, buckets)
    return [base + (1 if i < extra else 0) for i in range(buckets)]


# --------------------------------------------------------------------------- #
# Slot finding
# --------------------------------------------------------------------------- #
def _busy_for_day(d: date, busy: list[Interval]) -> list[Interval]:
    return sorted(
        [b for b in busy if b.start.date() == d or b.end.date() == d],
        key=lambda b: b.start,
    )


def _free_slots(d: date, duration: int, win_start: time, win_end: time,
                gran: int, gap: int, busy: list[Interval]) -> list[datetime]:
    """All slot start-times on day d where a meeting of `duration` minutes
    fits inside the window without colliding with busy intervals (+gap)."""
    day_busy = _busy_for_day(d, busy)
    slots = []
    cur = datetime.combine(d, win_start)
    end_limit = datetime.combine(d, win_end)
    step = timedelta(minutes=gran)
    dur = timedelta(minutes=duration)
    gapd = timedelta(minutes=gap)
    while cur + dur <= end_limit:
        s, e = cur, cur + dur
        clash = any(
            (s - gapd) < b.end and (e + gapd) > b.start for b in day_busy
        )
        if not clash:
            slots.append(cur)
        cur += step
    return slots


def _spread_pick(free: list[datetime], placed_today: list[datetime]) -> datetime | None:
    """Pick the free slot that is farthest from anything already placed that
    day, so meetings spread out instead of clumping at the top of the window."""
    if not free:
        return None
    if not placed_today:
        return free[len(free) // 2]  # start mid-window
    best, best_dist = None, -1.0
    for s in free:
        dist = min(abs((s - p).total_seconds()) for p in placed_today)
        if dist > best_dist:
            best, best_dist = s, dist
    return best


# --------------------------------------------------------------------------- #
# Main planner
# --------------------------------------------------------------------------- #
def build_plan(cfg: dict, existing_busy: list[Interval]) -> list[Proposal]:
    year, month = (int(x) for x in cfg["month"].split("-"))
    win_start = datetime.strptime(cfg["time_window"]["start"], "%H:%M").time()
    win_end = datetime.strptime(cfg["time_window"]["end"], "%H:%M").time()
    gran = int(cfg.get("slot_granularity_minutes", 30))
    gap = int(cfg.get("min_gap_minutes", 0))
    work_days = set(cfg.get("work_days", [0, 1, 2, 3, 4]))  # Mon..Fri

    meetings = resolve_meetings(cfg)
    staff = staff_emails(cfg)

    # Mutable busy list grows as we place new meetings (prevents self-clash).
    busy = list(existing_busy)
    # Track minutes booked per day for load balancing across all meeting types.
    day_load: dict[date, int] = {}
    placed_per_day: dict[date, list[datetime]] = {}
    proposals: list[Proposal] = []

    def place_one(meeting: Meeting, day: date, attendees: list[str]) -> bool:
        free = _free_slots(day, meeting.duration_minutes, win_start, win_end,
                           gran, gap, busy)
        slot = _spread_pick(free, placed_per_day.get(day, []))
        if slot is None:
            return False
        end = slot + timedelta(minutes=meeting.duration_minutes)
        proposals.append(Proposal(meeting.name, slot, end, meeting.category,
                                   meeting.schedule_type, list(attendees)))
        busy.append(Interval(slot, end))
        day_load[day] = day_load.get(day, 0) + meeting.duration_minutes
        placed_per_day.setdefault(day, []).append(slot)
        return True

    def least_loaded(days: list[date]) -> list[date]:
        return sorted(days, key=lambda d: (day_load.get(d, 0), d))

    # Larger meetings first so the heavy items get the easy slots.
    meetings.sort(key=lambda m: m.duration_minutes, reverse=True)

    unplaced: list[tuple[str, int]] = []

    for meeting in meetings:
        cands = candidate_days(meeting, year, month, work_days)
        if not cands:
            unplaced.append((meeting.name, meeting.count))
            continue

        # Who gets invited to each of this meeting's instances.
        if meeting.category == "DS" and staff:
            att_queue = [[staff[i % len(staff)]] for i in range(meeting.count)]
        elif meeting.invite_all_staff and staff:
            att_queue = [list(staff) for _ in range(meeting.count)]
        else:
            att_queue = [[] for _ in range(meeting.count)]
        ai = 0  # index into att_queue

        if meeting.schedule_type == 4:
            # Group candidate days into Sunday-start weeks, keep first 4.
            weeks: dict[date, list[date]] = {}
            for d in cands:
                weeks.setdefault(week_start_sunday(d), []).append(d)
            week_keys = sorted(weeks)[:4]
            per_week = distribute(meeting.count, len(week_keys))
            for wk, n in zip(week_keys, per_week):
                for _ in range(n):
                    attendees = att_queue[ai] if ai < len(att_queue) else []
                    for day in least_loaded(weeks[wk]):
                        if place_one(meeting, day, attendees):
                            break
                    else:
                        unplaced.append((meeting.name, 1))
                    ai += 1
        else:
            for _ in range(meeting.count):
                attendees = att_queue[ai] if ai < len(att_queue) else []
                for day in least_loaded(cands):
                    if place_one(meeting, day, attendees):
                        break
                else:
                    unplaced.append((meeting.name, 1))
                ai += 1

    proposals.sort(key=lambda p: p.start)
    return proposals, unplaced


# --------------------------------------------------------------------------- #
# Conflict check (used by tests and the GUI when the user edits times)
# --------------------------------------------------------------------------- #
def has_conflicts(proposals: list[Proposal],
                  existing_busy: list[Interval]) -> list[str]:
    issues = []
    allivals = [Interval(p.start, p.end) for p in proposals]
    names = [p.name for p in proposals]
    # vs each other
    for i in range(len(allivals)):
        for j in range(i + 1, len(allivals)):
            a, b = allivals[i], allivals[j]
            if a.start < b.end and b.start < a.end:
                issues.append(f"'{names[i]}' overlaps '{names[j]}'")
    # vs existing calendar
    for p, iv in zip(proposals, allivals):
        for b in existing_busy:
            if iv.start < b.end and b.start < iv.end:
                issues.append(f"'{p.name}' overlaps an existing calendar item")
                break
    return issues
