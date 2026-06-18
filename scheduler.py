"""
scheduler.py
============
Pure scheduling engine. No Outlook / no GUI / no cloud dependencies, so it can
be unit-tested anywhere. Given a config and existing busy intervals, it returns
proposed appointments.

What changed in the agents/coaches version
-------------------------------------------
Meetings now belong to one of three *populations*:

  * "manager"  - the manager's own meeting (e.g. Focus KPI). Hosted on the
                 manager's calendar. No invitee unless invite_all_staff.
  * "coaches"  - one instance per coach, invites that coach. Hosted on the
                 MANAGER's calendar (these are the manager's 1:1s with coaches).
  * "agents"   - one instance per agent, invites that agent. Hosted on that
                 agent's COACH's calendar, with the coach as organiser. This is
                 what keeps agent meetings OFF the manager's calendar.

Because each calendar owner has their own existing commitments, scheduling now
runs once per owner (manager + each coach) against that owner's busy list, via
`build_all_plans`. The legacy single-calendar `build_plan` is kept intact for
the existing tests and for anyone still on the flat `staff` model.

Schedule types (unchanged)
--------------------------
  1 = on/before the 10th        2 = 11th-20th         3 = 21st-end
  4 = once per week (up to 4 Sunday-start weeks)       5 = anytime
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
    category: str                 # "FS" or "DS"
    schedule_type: int            # 1..5
    duration_minutes: int
    count: int                    # final, resolved number of instances
    invite_all_staff: bool = False
    population: str = "agents"    # "manager" | "coaches" | "agents"


@dataclass
class Proposal:
    name: str
    start: datetime
    end: datetime
    category: str
    schedule_type: int
    attendees: list[str] = field(default_factory=list)
    owner: str = ""               # email of the calendar owner / organiser

    def as_row(self):
        if not self.attendees:
            who = ""
        elif len(self.attendees) == 1:
            who = self.attendees[0]
        else:
            who = f"{len(self.attendees)} invitees"
        return (
            self.owner,                                  # host calendar
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
def _emails(raw) -> list[str]:
    """Normalise a list that may contain plain strings or {"email": ...}."""
    out = []
    for item in raw or []:
        email = item.get("email") if isinstance(item, dict) else item
        email = (email or "").strip()
        if email:
            out.append(email)
    return out


def staff_emails(cfg: dict) -> list[str]:
    """Legacy flat staff list (kept for backward compatibility / old tests)."""
    return _emails(cfg.get("staff", []))


def organizer(cfg: dict) -> str:
    return (cfg.get("organizer") or "").strip()


def coaches(cfg: dict) -> list[str]:
    return _emails(cfg.get("coaches", []))


def agents_by_coach(cfg: dict) -> dict[str, list[str]]:
    """Map of coach-email -> list of that coach's agent emails.

    The mapping itself usually arrives from the shared CoachAgentMapping list
    (written to mapping.json by a flow) and is merged into cfg before planning;
    this just normalises whatever is present."""
    raw = cfg.get("agents_by_coach", {}) or {}
    out: dict[str, list[str]] = {}
    for coach, agents in raw.items():
        c = (coach or "").strip()
        if c:
            out[c] = _emails(agents)
    return out


def agents_of(cfg: dict, coach: str) -> list[str]:
    return agents_by_coach(cfg).get(coach, [])


def all_owners(cfg: dict) -> list[str]:
    """Every calendar that will be written to: the manager plus each coach."""
    owners = [organizer(cfg)] if organizer(cfg) else []
    owners += [c for c in coaches(cfg) if c]
    return owners


def meeting_templates(cfg: dict) -> list[dict]:
    """Raw meeting dicts with a guaranteed `population` key."""
    out = []
    for m in cfg["meetings"]:
        m = dict(m)
        m.setdefault("population", "agents" if m.get("category") == "DS" else "manager")
        out.append(m)
    return out


def resolve_meetings(cfg: dict) -> list[Meeting]:
    """LEGACY: flat-staff resolution used by the old single-calendar build_plan.
    DS count = per_staff * len(staff). Untouched so existing tests still pass."""
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
                name=m["name"], category=m["category"],
                schedule_type=int(m["schedule_type"]),
                duration_minutes=int(m["duration_minutes"]),
                count=count,
                invite_all_staff=bool(m.get("invite_all_staff", False)),
                population=m.get("population", "agents" if m["category"] == "DS" else "manager"),
            )
        )
    return out


# --------------------------------------------------------------------------- #
# Calendar helpers  (unchanged)
# --------------------------------------------------------------------------- #
def month_days(year: int, month: int) -> list[date]:
    last = calendar.monthrange(year, month)[1]
    return [date(year, month, d) for d in range(1, last + 1)]


def week_start_sunday(d: date) -> date:
    return d - timedelta(days=(d.weekday() + 1) % 7)


def candidate_days(meeting: Meeting, year: int, month: int,
                   work_days: set[int]) -> list[date]:
    days = [d for d in month_days(year, month) if d.weekday() in work_days]
    t = meeting.schedule_type
    if t == 1:
        return [d for d in days if d.day <= 10]
    if t == 2:
        return [d for d in days if 11 <= d.day <= 20]
    if t == 3:
        return [d for d in days if d.day >= 21]
    return days


def distribute(total: int, buckets: int) -> list[int]:
    if buckets <= 0:
        return []
    base, extra = divmod(total, buckets)
    return [base + (1 if i < extra else 0) for i in range(buckets)]


# --------------------------------------------------------------------------- #
# Slot finding  (unchanged)
# --------------------------------------------------------------------------- #
def _busy_for_day(d: date, busy: list[Interval]) -> list[Interval]:
    return sorted(
        [b for b in busy if b.start.date() == d or b.end.date() == d],
        key=lambda b: b.start,
    )


def _free_slots(d: date, duration: int, win_start: time, win_end: time,
                gran: int, gap: int, busy: list[Interval]) -> list[datetime]:
    day_busy = _busy_for_day(d, busy)
    slots = []
    cur = datetime.combine(d, win_start)
    end_limit = datetime.combine(d, win_end)
    step = timedelta(minutes=gran)
    dur = timedelta(minutes=duration)
    gapd = timedelta(minutes=gap)
    while cur + dur <= end_limit:
        s, e = cur, cur + dur
        clash = any((s - gapd) < b.end and (e + gapd) > b.start for b in day_busy)
        if not clash:
            slots.append(cur)
        cur += step
    return slots


def _spread_pick(free: list[datetime], placed_today: list[datetime]) -> datetime | None:
    if not free:
        return None
    if not placed_today:
        return free[len(free) // 2]
    best, best_dist = None, -1.0
    for s in free:
        dist = min(abs((s - p).total_seconds()) for p in placed_today)
        if dist > best_dist:
            best, best_dist = s, dist
    return best


# --------------------------------------------------------------------------- #
# Core placement  (extracted so both build_plan and build_all_plans reuse it)
# --------------------------------------------------------------------------- #
def _place_meetings(cfg: dict, plan_items: list[tuple[Meeting, list[list[str]]]],
                    existing_busy: list[Interval], owner: str = ""):
    """Place a set of meetings for a single calendar owner.

    plan_items: list of (Meeting, attendee_queue) where attendee_queue is a list
    of attendee-lists, one per instance (len == meeting.count).
    Returns (proposals, unplaced)."""
    year, month = (int(x) for x in cfg["month"].split("-"))
    win_start = datetime.strptime(cfg["time_window"]["start"], "%H:%M").time()
    win_end = datetime.strptime(cfg["time_window"]["end"], "%H:%M").time()
    gran = int(cfg.get("slot_granularity_minutes", 30))
    gap = int(cfg.get("min_gap_minutes", 0))
    work_days = set(cfg.get("work_days", [0, 1, 2, 3, 4]))

    busy = list(existing_busy)
    day_load: dict[date, int] = {}
    placed_per_day: dict[date, list[datetime]] = {}
    proposals: list[Proposal] = []
    unplaced: list[tuple[str, int]] = []

    def place_one(meeting: Meeting, day: date, attendees: list[str]) -> bool:
        free = _free_slots(day, meeting.duration_minutes, win_start, win_end,
                           gran, gap, busy)
        slot = _spread_pick(free, placed_per_day.get(day, []))
        if slot is None:
            return False
        end = slot + timedelta(minutes=meeting.duration_minutes)
        proposals.append(Proposal(meeting.name, slot, end, meeting.category,
                                  meeting.schedule_type, list(attendees), owner))
        busy.append(Interval(slot, end))
        day_load[day] = day_load.get(day, 0) + meeting.duration_minutes
        placed_per_day.setdefault(day, []).append(slot)
        return True

    def least_loaded(days: list[date]) -> list[date]:
        return sorted(days, key=lambda d: (day_load.get(d, 0), d))

    # Larger meetings first so heavy items get the easy slots.
    plan_items = sorted(plan_items, key=lambda it: it[0].duration_minutes, reverse=True)

    for meeting, att_queue in plan_items:
        cands = candidate_days(meeting, year, month, work_days)
        if not cands:
            unplaced.append((meeting.name, meeting.count))
            continue
        ai = 0
        if meeting.schedule_type == 4:
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
# LEGACY single-calendar planner (kept for old tests / flat-staff configs)
# --------------------------------------------------------------------------- #
def build_plan(cfg: dict, existing_busy: list[Interval]):
    staff = staff_emails(cfg)
    plan_items: list[tuple[Meeting, list[list[str]]]] = []
    for meeting in resolve_meetings(cfg):
        if meeting.category == "DS" and staff:
            att = [[staff[i % len(staff)]] for i in range(meeting.count)]
        elif meeting.invite_all_staff and staff:
            att = [list(staff) for _ in range(meeting.count)]
        else:
            att = [[] for _ in range(meeting.count)]
        plan_items.append((meeting, att))
    return _place_meetings(cfg, plan_items, existing_busy, owner=organizer(cfg))


# --------------------------------------------------------------------------- #
# NEW multi-calendar planner
# --------------------------------------------------------------------------- #
def _meeting_from_template(t: dict, count: int) -> Meeting:
    return Meeting(
        name=t["name"], category=t["category"],
        schedule_type=int(t["schedule_type"]),
        duration_minutes=int(t["duration_minutes"]),
        count=count,
        invite_all_staff=bool(t.get("invite_all_staff", False)),
        population=t["population"],
    )


def build_all_plans(cfg: dict, busy_by_owner: dict[str, list[Interval]]):
    """Plan every calendar. `busy_by_owner` maps each owner email (the manager
    and each coach) to that owner's existing busy intervals for the month.

    Returns {owner_email: (proposals, unplaced)}.

      * manager owner gets the "manager" and "coaches" population meetings;
      * each coach owner gets the "agents" population meetings for THEIR agents.
    """
    mgr = organizer(cfg)
    templates = meeting_templates(cfg)
    out: dict[str, tuple[list[Proposal], list]] = {}

    # ---- Manager's own calendar: manager-pop + coaches-pop ----------------- #
    coach_list = coaches(cfg)
    mgr_items: list[tuple[Meeting, list[list[str]]]] = []
    for t in templates:
        pop = t["population"]
        if pop == "manager":
            count = int(t.get("count", 1))
            if count <= 0:
                continue
            if t.get("invite_all_staff") and coach_list:
                att = [list(coach_list) for _ in range(count)]
            else:
                att = [[] for _ in range(count)]
            mgr_items.append((_meeting_from_template(t, count), att))
        elif pop == "coaches":
            count = int(t.get("per_staff", 1)) * len(coach_list)
            if count <= 0:
                continue
            att = [[coach_list[i % len(coach_list)]] for i in range(count)]
            mgr_items.append((_meeting_from_template(t, count), att))
    if mgr:
        out[mgr] = _place_meetings(cfg, mgr_items, busy_by_owner.get(mgr, []), owner=mgr)

    # ---- Each coach's calendar: agents-pop for that coach's agents --------- #
    for coach in coach_list:
        my_agents = agents_of(cfg, coach)
        coach_items: list[tuple[Meeting, list[list[str]]]] = []
        for t in templates:
            if t["population"] != "agents":
                continue
            count = int(t.get("per_staff", 1)) * len(my_agents)
            if count <= 0:
                continue
            att = [[my_agents[i % len(my_agents)]] for i in range(count)]
            coach_items.append((_meeting_from_template(t, count), att))
        out[coach] = _place_meetings(cfg, coach_items, busy_by_owner.get(coach, []), owner=coach)

    return out


def flatten_proposals(plans: dict[str, tuple[list[Proposal], list]]) -> list[Proposal]:
    """All proposals across all owners, sorted by start, for a unified preview."""
    rows: list[Proposal] = []
    for proposals, _ in plans.values():
        rows.extend(proposals)
    rows.sort(key=lambda p: (p.start, p.owner))
    return rows


# --------------------------------------------------------------------------- #
# Conflict check  (unchanged; run per owner)
# --------------------------------------------------------------------------- #
def has_conflicts(proposals: list[Proposal], existing_busy: list[Interval]) -> list[str]:
    issues = []
    allivals = [Interval(p.start, p.end) for p in proposals]
    names = [p.name for p in proposals]
    for i in range(len(allivals)):
        for j in range(i + 1, len(allivals)):
            a, b = allivals[i], allivals[j]
            if a.start < b.end and b.start < a.end:
                issues.append(f"'{names[i]}' overlaps '{names[j]}'")
    for p, iv in zip(proposals, allivals):
        for b in existing_busy:
            if iv.start < b.end and b.start < iv.end:
                issues.append(f"'{p.name}' overlaps an existing calendar item")
                break
    return issues
