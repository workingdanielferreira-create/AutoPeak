"""
outlook_client.py
=================
Talks to the locally-installed *classic* Outlook desktop app through COM
automation (pywin32). This needs no Azure app registration, no admin consent,
and no stored credentials -- it simply drives the Outlook session the user is
already signed into. That is what makes a true double-click experience possible.

If COM is unavailable (e.g. the machine only has the "new Outlook" which does
not expose COM), every function degrades gracefully and the caller can fall
back to `export_ics`, which writes a .ics file the user imports manually.

UNCHANGED for the agents/coaches build: get_busy() and commit() still only
ever touch the MANAGER's own calendar -- that's the one calendar COM can
reliably drive locally. Coach calendars are handled entirely through
bridge.py + Power Automate; this file never talks to them. Callers (app.py)
are responsible for only passing manager-owned proposals into commit().
"""

from __future__ import annotations
from datetime import datetime, timedelta
import calendar

try:
    import win32com.client  # type: ignore
    import pythoncom        # type: ignore
    _HAVE_COM = True
except Exception:
    _HAVE_COM = False


class OutlookUnavailable(RuntimeError):
    pass


def com_available() -> bool:
    return _HAVE_COM


def _app():
    if not _HAVE_COM:
        raise OutlookUnavailable(
            "Classic Outlook COM automation is not available on this machine."
        )
    pythoncom.CoInitialize()
    return win32com.client.Dispatch("Outlook.Application")


# --------------------------------------------------------------------------- #
# Read existing calendar items for the target month
# --------------------------------------------------------------------------- #
def get_busy(year: int, month: int):
    """Return a list of (start, end) datetime tuples for existing items in the
    given month, expanding recurring appointments. Always the signed-in
    manager's own calendar -- see module docstring."""
    from scheduler import Interval

    app = _app()
    ns = app.GetNamespace("MAPI")
    cal = ns.GetDefaultFolder(9)  # 9 = olFolderCalendar
    items = cal.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")

    last_day = calendar.monthrange(year, month)[1]
    start = datetime(year, month, 1, 0, 0)
    end = datetime(year, month, last_day, 23, 59)

    restriction = (
        f"[Start] <= '{end.strftime('%m/%d/%Y %H:%M %p')}' AND "
        f"[End] >= '{start.strftime('%m/%d/%Y %H:%M %p')}'"
    )
    busy = []
    for it in items.Restrict(restriction):
        try:
            if int(it.BusyStatus) == 0:   # 0 = Free, ignore "free" blocks
                continue
            s = datetime.fromtimestamp(it.Start.timestamp())
            e = datetime.fromtimestamp(it.End.timestamp())
            busy.append(Interval(s, e))
        except Exception:
            continue
    return busy


# --------------------------------------------------------------------------- #
# Write the approved proposals back to the calendar
# --------------------------------------------------------------------------- #
def commit(proposals, category_prefix=True, reminder_minutes=15,
           busy_status=2, send_invites=True):
    """Create each proposal in Outlook. If a proposal has attendees and
    send_invites is True, it is created as a meeting request and sent so the
    staff members receive an invitation; otherwise it is saved as a plain
    appointment. busy_status: 2 = Busy.

    Pass ONLY manager-owned proposals here (owner == cfg's organizer) -- this
    always writes to the signed-in user's own calendar, regardless of what a
    proposal's `owner` field says."""
    app = _app()
    created, invited = 0, 0
    for p in proposals:
        appt = app.CreateItem(1)  # 1 = olAppointmentItem
        subject = f"[{p.category}] {p.name}" if category_prefix else p.name
        appt.Subject = subject
        appt.Start = p.start.strftime("%Y-%m-%d %H:%M")
        appt.Duration = int((p.end - p.start).total_seconds() // 60)
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
        appt.BusyStatus = busy_status
        appt.Body = (f"Auto-scheduled by AutoPeak.\n"
                     f"Category: {p.category} | Schedule type: {p.schedule_type}")

        attendees = getattr(p, "attendees", []) or []
        if attendees and send_invites:
            appt.MeetingStatus = 1  # olMeeting -> enables invitations
            for email in attendees:
                rcpt = appt.Recipients.Add(email)
                rcpt.Type = 1       # olRequired
            appt.Recipients.ResolveAll()
            appt.Send()             # actually dispatches the invitations
            invited += 1
        else:
            appt.Save()
        created += 1
    return created, invited


# --------------------------------------------------------------------------- #
# Fallback: write a standard .ics the user can import into any calendar
# --------------------------------------------------------------------------- #
def export_ics(proposals, path: str, organizer: str | None = None):
    """organizer is the default; if a proposal has its own `owner` set (the
    agents/coaches build), that takes precedence per-event so a mixed-owner
    export still tags each meeting with the right host."""
    def fmt(dt: datetime) -> str:
        return dt.strftime("%Y%m%dT%H%M%S")

    lines = ["BEGIN:VCALENDAR", "VERSION:2.0",
             "PRODID:-//AutoPeak//EN", "CALSCALE:GREGORIAN",
             "METHOD:REQUEST"]
    stamp = datetime.now().strftime("%Y%m%dT%H%M%SZ")
    for i, p in enumerate(proposals):
        lines += [
            "BEGIN:VEVENT",
            f"UID:autopeak-{i}-{fmt(p.start)}@local",
            f"DTSTAMP:{stamp}",
            f"DTSTART:{fmt(p.start)}",
            f"DTEND:{fmt(p.end)}",
            f"SUMMARY:[{p.category}] {p.name}",
            f"DESCRIPTION:Auto-scheduled. Type {p.schedule_type}.",
        ]
        event_organizer = getattr(p, "owner", None) or organizer
        if event_organizer:
            lines.append(f"ORGANIZER:mailto:{event_organizer}")
        for email in (getattr(p, "attendees", []) or []):
            lines.append(
                "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;"
                f"RSVP=TRUE:mailto:{email}")
        lines += [
            "BEGIN:VALARM", "TRIGGER:-PT15M", "ACTION:DISPLAY",
            "DESCRIPTION:Reminder", "END:VALARM",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\r\n".join(lines))
    return path
