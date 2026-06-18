# Monthly Meeting Scheduler

A small Windows tool that reads your Outlook calendar for a chosen month,
proposes a balanced, conflict-free set of meetings, lets you preview and tweak
them, then writes them back to Outlook on your approval.

## Files
- `app.py` – the GUI (double-click entry point)
- `scheduler.py` – the scheduling engine (no Outlook needed; this is the brain)
- `outlook_client.py` – reads/writes classic Outlook via COM, plus .ics export
- `config.json` – everything you can change without touching code

## Run it during testing (Python installed)
```
pip install pywin32
python app.py
```

## Build the double-click .exe (do this once, on a Windows PC)
```
pip install pywin32 pyinstaller
pyinstaller --onefile --windowed --name MonthlyScheduler app.py
```
The exe appears in `dist\MonthlyScheduler.exe`. Put `config.json` in the same
folder as the exe. From then on it's just double-click — no install.

> Note: a one-file exe with no code-signing certificate can trip corporate
> antivirus/SmartScreen on first run. If that's a blocker, ask IT to whitelist
> it, or use `--onedir` instead of `--onefile` (a folder rather than a single
> file, which flags less often).

## Staff list and invitations
`config.json` now has a `staff` list of **email addresses** — one entry per
direct report. The count is derived from that list (no more "staff count"
field). In the app, click **Staff emails…** to edit the list, one per line.

Each dynamic (DS) meeting is created **once per staff member and sent to that
person as a meeting invitation** when you Commit. The Invitee column in the
preview shows who each meeting goes to. The fixed Focus KPI meeting has no
invitee by default; set `"invite_all_staff": true` on it in `config.json` to
invite the whole team instead.

> Sending invites uses Outlook's `Send()`, so they go out immediately on
> Commit — there's a confirmation prompt first. Exporting to .ics only adds the
> meetings to your own calendar; it does **not** send invitations.

## How scheduling works
- **Schedule types**: 1 = by the 10th · 2 = 11th–20th · 3 = 21st–end ·
  4 = once per week spread across up to 4 weeks (weeks start Sunday) ·
  5 = anytime in the month.
- **Dynamic counts**: each DS meeting = `per_staff × staff_count`.
- **Balance**: meetings go to the least-loaded eligible day, and within a day
  pick the slot farthest from anything already booked, so they spread out.
- **No clashes**: existing Outlook items (and meetings placed earlier in the
  same run) are treated as busy; a `min_gap_minutes` buffer is added around each.

## Outlook compatibility
COM automation works with **classic** Outlook for Windows. The "new Outlook"
does not expose COM — on that, the tool still previews and you use **Export
.ics** then File ▸ Open ▸ Import. (A Microsoft Graph version is the alternative
if you need new-Outlook/web write access; that needs an Azure app registration.)
