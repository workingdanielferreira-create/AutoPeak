# AutoPeak — Agents/Coaches Build Spec (Zero-IT)

This is the concrete build for splitting staff into **agents** and **coaches** so
that agent-SKEP meetings land on each **coach's** calendar (coach as organiser),
the manager's calendar only ever carries the manager-level meetings, and **no
step anywhere requires IT/admin approval**.

The guiding constraint throughout: the people using this should touch as few
buttons as possible, and those buttons should be the same every month.

---

## 1. The principle that makes it zero-IT

Two facts do the heavy lifting:

1. **Python never talks to the cloud.** The engine only reads/writes *local
   files* in a folder that happens to be OneDrive-synced. No SharePoint/Graph
   auth, so no app registration and no consent prompt.
2. **Power Automate does all cloud work on connections the user authorises
   themselves.** The Office 365 Outlook and SharePoint/OneDrive connectors are
   *standard* connectors included with F3 — building flows on your own
   connection needs no admin. (Your tenant already runs standard flows.)

The two halves meet at a **drop folder**: a shared SharePoint document library
(synced to the manager's PC). The engine drops files in; flows pick them up and
do the calendar work; flows drop results back.

```
 Manager PC (Python + desktop Outlook)         Cloud (Power Automate, std connectors)
 ───────────────────────────────────           ─────────────────────────────────────
 read manager busy (existing COM)
 read CoachAgentMapping.xlsx (local)
 write request_<month>.json  ───────────────►  Flow A: read coach calendars,
 read  busy_<month>.json     ◄───────────────          write busy_<month>.json
 build_all_plans() → preview
 COM-commit manager-owned meetings
 write proposals_<month>.json ──────────────►  Flow B: create events on coach
 read  results_<month>.json   ◄──────────────          calendars, write results
```

Manager-owned meetings (Focus KPI + the coach-level SKEPs) are still written
straight to the manager's own calendar by the **existing COM path** — it's
proven, and the manager is the correct organiser for those. Only the
**coach-hosted agent meetings** travel through the bridge.

---

## 2. What each person actually does

**Coaches — one-time only, then never again:**
1. In Outlook on the web: share their calendar with the manager, permission
   **"Can edit"**. (~5 clicks. This is a normal user action; no admin.)
2. Keep their agent list current in one shared spreadsheet (their own rows).

That's it. Coaches click nothing each month and install nothing.

**Manager — every month, ~4 clicks:**
1. Open AutoPeak, choose the month.
2. Click **Fetch availability** (tool reads its own calendar via Outlook and
   waits for the coach-availability file).
3. Click **Generate**, glance at the preview (grouped by host calendar), tweak
   if needed.
4. Click **Publish** — manager meetings go onto the manager's calendar; agent
   meetings are created on each coach's calendar with invites sent. A status
   line reports how many landed where.

The mechanics (files, flows, SharePoint) are invisible behind those buttons.

---

## 3. The drop folder & file contracts

One shared document library, e.g. `Documents/AutoPeak/Drop`, synced to the
manager's PC. All files are JSON (trivial for Power Automate's **Parse JSON** to
read — no CSV parsing pain) except the mapping, which is an Excel file coaches
can edit comfortably.

| File | Direction | Purpose |
|---|---|---|
| `CoachAgentMapping.xlsx` | coaches → tool | columns `CoachEmail`, `AgentEmail`; the tool reads it locally with openpyxl |
| `request_<month>.json` | tool → Flow A | `{"month":"2026-07","owners":["mgr@…","coach1@…",…]}` |
| `busy_<month>.json` | Flow A → tool | `{"coach1@…":[["2026-07-06T09:00","2026-07-06T12:00"],…], …}` |
| `proposals_<month>.json` | tool → Flow B | array of proposal objects (below); coach-owned only |
| `results_<month>.json` | Flow B → tool | `[{"id":"p12","status":"created"},{"id":"p13","status":"failed","error":"…"}]` |

**Proposal object** (matches `Proposal` in `scheduler.py`):

```json
{
  "id": "p001",
  "owner": "coach1@foundever.com",
  "name": "Initial Agent SKEP",
  "category": "DS",
  "schedule_type": 1,
  "start": "2026-07-03T13:00:00",
  "end":   "2026-07-03T13:30:00",
  "attendee": "agent1@foundever.com",
  "status": "Pending"
}
```

---

## 4. Power Automate — PRIMARY path (manager-owned, shared calendars)

Both flows use the manager's own connections (Office 365 Outlook + SharePoint).
Nothing here needs admin.

### Flow A — "AutoPeak – Export Busy"
| # | Action | Notes |
|---|---|---|
| 1 | **When a file is created** (SharePoint) in `…/Drop` | filter to names starting `request_` |
| 2 | **Get file content** | the request JSON |
| 3 | **Parse JSON** | schema = `{month, owners[]}` |
| 4 | **Get calendars (V2)** | returns the manager's own calendar **and** every coach calendar shared with edit rights — that's exactly why the "Can edit" share matters |
| 5 | **Apply to each** `owners` | |
| 5a | — **Filter array** on step 4 output | match the calendar whose owner/name = current owner; take its calendar id |
| 5b | — **Get events (V4)** | that calendar id; window = month start→end; pull `start`, `end`, `showAs` |
| 5c | — **Select** | keep events where `showAs ≠ free` and not cancelled → `[start,end]` pairs; append to an object keyed by owner |
| 6 | **Create file** `busy_<month>.json` | the assembled `{owner:[[s,e]…]}` |

### Flow B — "AutoPeak – Create Events"
| # | Action | Notes |
|---|---|---|
| 1 | **When a file is created** (SharePoint) in `…/Drop` | filter to `proposals_` |
| 2 | **Get file content** → **Parse JSON** | array of proposal objects |
| 3 | **Get calendars (V2)** | once, to resolve owner→calendar id |
| 4 | **Apply to each** proposal where `status = Pending` | |
| 4a | — resolve `owner` → calendar id (Filter array on step 3) | |
| 4b | — **Create event (V4)** | Calendar id = resolved; Subject `[{category}] {name}`; Start/End with time zone; **Required attendees** = `attendee`; reminder 15 min; body note |
| 4c | — **Append to array** | `{id, status:"created"}` or `{id,status:"failed",error:…}` (wrap 4b in a scope with **Configure run after** to catch failures) |
| 5 | **Create file** `results_<month>.json` | the results array |

> **The one thing to validate before trusting Flow B — see §7.** When the
> manager's connection creates an event on a coach's *shared* calendar, the
> coach should come out as the **organiser** (so it lives only on the coach's
> calendar). Microsoft's shared-calendar behaviour says the owner becomes the
> organiser, but the connector layer is worth proving once.

---

## 5. Power Automate — FALLBACK path (per-coach, own calendar)

Use this only if the §7 test shows the manager ends up as organiser. Here each
coach owns the writes to their **own** calendar, so the coach is *unavoidably*
the organiser, and **no calendar sharing is needed at all**. Each coach imports
one provided flow once and authorises their own Outlook connection.

To avoid two coaches writing the same file at once, the busy/results files are
**per-coach** (`busy_<month>_<alias>.json`), and the tool merges them.

### Flow per coach — "AutoPeak – My Calendar"
| # | Action | Notes |
|---|---|---|
| 1 | **When a file is created** in `…/Drop` | one trigger handles both phases by switching on the file name (`request_` vs `proposals_`) |
| 2 | **If** name starts `request_` → **Get events (V4)** on *my* calendar for the month → **Create file** `busy_<month>_<me>.json` | availability export |
| 3 | **If** name starts `proposals_` → **Parse JSON** → **Filter** to `owner = me` and `status = Pending` → **Create event (V4)** on *my* calendar (organiser = me) inviting `attendee` → **Create file** `results_<month>_<me>.json` | event creation |

Everything else (engine, preview, the manager's own COM commit) is identical to
the primary path — only the owner of the create step moves. That sameness is
deliberate: switching paths is a config flip, not a redesign.

---

## 6. Code changes

### `scheduler.py` (delivered, tested)
- `Meeting` gains `population` — `"manager" | "coaches" | "agents"`.
- `Proposal` gains `owner` (the host calendar / organiser email); `as_row()`
  now leads with the host so the preview can group by calendar.
- New config accessors: `organizer`, `coaches`, `agents_by_coach`, `agents_of`,
  `all_owners`, `meeting_templates`.
- Placement logic extracted into `_place_meetings(...)` and reused by both
  planners (behaviour identical to before).
- **New entry point** `build_all_plans(cfg, busy_by_owner) -> {owner:(proposals,
  unplaced)}`: plans the manager's calendar (manager- + coaches-population
  meetings) and each coach's calendar (that coach's agents) **independently**,
  each against its own busy list. `flatten_proposals(plans)` gives one sorted
  list for the preview.
- The old `build_plan` and `has_conflicts` are untouched, so existing tests and
  any flat-`staff` config keep working. Run `has_conflicts` per owner.

### `outlook_client.py`
- Keep `get_busy` / `commit` for the **manager's own** calendar (COM, proven).
- `commit` should be called with only the manager-owned subset of proposals.
- Add a small `bridge.py` (local-file I/O only): `write_request`, `read_busy`,
  `write_proposals`, `read_results`, plus `read_mapping(xlsx_path)` via openpyxl.
  No network code lives here.

### `app.py`
- Staff editor → two lists (coaches; agents are read from the mapping sheet).
- Preview table gains a leading **Host calendar** column and groups rows by it,
  so the manager can see at a glance what lands where.
- **Publish** splits proposals by `owner`: manager-owned → `commit()` (COM);
  coach-owned → `bridge.write_proposals()`. Then poll `read_results()`.

### `config.json` (new shape)
```json
{
  "month": "2026-07",
  "organizer": "manager@foundever.com",
  "coaches": ["coach1@foundever.com", "coach2@foundever.com"],
  "agents_by_coach": {},
  "agent_mapping_xlsx": "Drop/CoachAgentMapping.xlsx",
  "drop_folder": "C:/Users/<manager>/Foundever/AutoPeak - Drop",
  "write_path": "primary",
  "work_days": [0,1,2,3,4],
  "time_window": {"start":"09:00","end":"17:00"},
  "slot_granularity_minutes": 30,
  "min_gap_minutes": 15,
  "meetings": [
    {"name":"Focus KPI Meeting","category":"FS","population":"manager","schedule_type":4,"duration_minutes":60,"count":1,"invite_all_staff":false},
    {"name":"CPA","category":"DS","population":"agents","schedule_type":4,"duration_minutes":30,"per_staff":1},
    {"name":"Initial Agent SKEP","category":"DS","population":"agents","schedule_type":1,"duration_minutes":30,"per_staff":1},
    {"name":"1st Agent SKEP follow up","category":"DS","population":"agents","schedule_type":2,"duration_minutes":30,"per_staff":1},
    {"name":"2nd Agent SKEP follow up","category":"DS","population":"agents","schedule_type":3,"duration_minutes":30,"per_staff":1},
    {"name":"Initial Coach SKEP","category":"DS","population":"coaches","schedule_type":2,"duration_minutes":45,"per_staff":1},
    {"name":"Coach SKEP follow up","category":"DS","population":"coaches","schedule_type":3,"duration_minutes":45,"per_staff":1},
    {"name":"Coach SKEP Audits","category":"DS","population":"coaches","schedule_type":5,"duration_minutes":30,"per_staff":1}
  ]
}
```
`agents_by_coach` is left empty in the file and filled at runtime from the
mapping sheet. **Two populations to confirm before first real run:** `CPA` is
set to `agents` (it could be a coach-level meeting — change to `coaches` if so),
and `Focus KPI` is `manager` with no invitee (set `invite_all_staff:true` to
invite all coaches).

---

## 7. The 15-minute validation test (decides primary vs fallback — needs no IT)

1. One volunteer coach shares their calendar with the manager as **"Can edit"**
   (Outlook web).
2. Manager builds a throwaway flow: **Get calendars (V2)** → pick the coach's
   calendar id → **Create event (V4)** on it, with a test agent as required
   attendee. Run it.
3. Check four things:
   - the event appears on the **coach's** calendar;
   - the **coach** is the organiser (not the manager);
   - the test attendee receives a real invite;
   - the **manager's** calendar is untouched.

All four pass → build the **primary** path (§4). Organiser comes back as the
manager → build the **fallback** path (§5). Either way you've spent 15 minutes
and zero approvals to know which one to commit to.

---

## 8. Suggested build order
1. Run the §7 test → lock primary vs fallback.
2. Create the Drop library + `CoachAgentMapping.xlsx`; have coaches fill rows
   (and, primary path, share calendars).
3. Drop in the new `scheduler.py` (done) and add `bridge.py`.
4. Build the chosen flow(s) from §4 or §5.
5. Wire `app.py`: two-list editor, host column, split Publish, results polling.
6. Trial run on a **test month** before a live month — Create still sends real
   invites the moment Publish writes the proposals file.
