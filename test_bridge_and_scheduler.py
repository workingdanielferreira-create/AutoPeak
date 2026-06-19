"""Synthetic, offline test of bridge.py + scheduler.build_all_plans together.
No Outlook, no Power Automate, no real network -- just proves the file
contracts and the multi-calendar planner behave as designed before anyone
touches a real mailbox.
"""
import json
import os
import shutil
import tempfile
from datetime import datetime

import openpyxl

import bridge as B
import scheduler as S

tmp = tempfile.mkdtemp()
drop = os.path.join(tmp, "Drop")
os.makedirs(drop)

# ---- 1. Build a fake CoachAgentMapping.xlsx -------------------------------
wb = openpyxl.Workbook()
ws = wb.active
ws.append(["CoachEmail", "AgentEmail"])
rows = [
    ("coach1@foundever.com", "agent1@foundever.com"),
    ("coach1@foundever.com", "agent2@foundever.com"),
    ("coach1@foundever.com", "agent3@foundever.com"),
    ("coach2@foundever.com", "agent4@foundever.com"),
    ("coach2@foundever.com", "agent5@foundever.com"),
]
for r in rows:
    ws.append(r)
mapping_path = os.path.join(drop, "CoachAgentMapping.xlsx")
wb.save(mapping_path)

cfg = {
    "month": "2026-07",
    "organizer": "manager@foundever.com",
    "coaches": [],
    "agents_by_coach": {},
    "agent_mapping_xlsx": mapping_path,
    "drop_folder": drop,
    "write_path": "primary",
    "work_days": [0, 1, 2, 3, 4],
    "time_window": {"start": "09:00", "end": "17:00"},
    "slot_granularity_minutes": 30,
    "min_gap_minutes": 15,
    "meetings": json.load(open("config.json"))["meetings"],
}

# ---- 2. merge_mapping_into_cfg should populate coaches + agents_by_coach --
B.merge_mapping_into_cfg(cfg)
assert set(cfg["coaches"]) == {"coach1@foundever.com", "coach2@foundever.com"}, cfg["coaches"]
assert cfg["agents_by_coach"]["coach1@foundever.com"] == [
    "agent1@foundever.com", "agent2@foundever.com", "agent3@foundever.com"]
assert cfg["agents_by_coach"]["coach2@foundever.com"] == ["agent4@foundever.com", "agent5@foundever.com"]
print("OK: merge_mapping_into_cfg ->", cfg["coaches"], cfg["agents_by_coach"])

# ---- 3. write_request -> simulate Flow A writing busy_<month>.json -------
owners = S.all_owners(cfg)
req_path = B.write_request(cfg, owners)
assert os.path.exists(req_path)
print("OK: write_request ->", json.load(open(req_path)))

# Pretend Flow A ran and dropped this:
busy_payload = {
    "coach1@foundever.com": [["2026-07-06T09:00:00", "2026-07-06T11:00:00"]],
    "coach2@foundever.com": [],
}
with open(os.path.join(drop, "busy_2026-07.json"), "w") as f:
    json.dump(busy_payload, f)

coach_busy = B.read_busy(cfg, cfg["coaches"])
assert coach_busy["coach1@foundever.com"][0][0] == datetime(2026, 7, 6, 9, 0)
print("OK: read_busy ->", {k: len(v) for k, v in coach_busy.items()})

# ---- 4. build_all_plans across manager + 2 coaches ------------------------
busy_by_owner = {cfg["organizer"]: []}
for coach, pairs in coach_busy.items():
    busy_by_owner[coach] = [S.Interval(s, e) for s, e in pairs]

plans = S.build_all_plans(cfg, busy_by_owner)
assert cfg["organizer"] in plans
assert "coach1@foundever.com" in plans and "coach2@foundever.com" in plans

mgr_props, mgr_unplaced = plans[cfg["organizer"]]
c1_props, c1_unplaced = plans["coach1@foundever.com"]
c2_props, c2_unplaced = plans["coach2@foundever.com"]

# Manager calendar should carry Focus KPI (4x, invite_all coaches) + 3x Coach SKEPs per coach (2 coaches)
mgr_names = sorted(p.name for p in mgr_props)
assert mgr_names.count("Focus KPI Meeting") == 4, mgr_names
assert mgr_names.count("Initial Coach SKEP") == 2, mgr_names  # one per coach
for p in mgr_props:
    if p.name == "Focus KPI Meeting":
        assert set(p.attendees) == {"coach1@foundever.com", "coach2@foundever.com"}

# Coach1 calendar should carry CPA + 3 agent-SKEP types x 3 agents = 12 instances
c1_names = sorted(p.name for p in c1_props)
assert c1_names.count("CPA") == 3, c1_names           # one per agent (3 agents)
assert c1_names.count("Initial Agent SKEP") == 3, c1_names
assert all(p.owner == "coach1@foundever.com" for p in c1_props)
agent_attendees_c1 = sorted({p.attendees[0] for p in c1_props if p.attendees})
assert agent_attendees_c1 == ["agent1@foundever.com", "agent2@foundever.com", "agent3@foundever.com"]

# Coach2 calendar should carry CPA + agent-SKEPs x 2 agents
c2_names = sorted(p.name for p in c2_props)
assert c2_names.count("CPA") == 2, c2_names

# No proposal should ever land on the wrong calendar
for owner, (props, _) in plans.items():
    assert all(p.owner == owner for p in props)

flat = S.flatten_proposals(plans)
assert len(flat) == len(mgr_props) + len(c1_props) + len(c2_props)
print(f"OK: build_all_plans -> manager={len(mgr_props)} coach1={len(c1_props)} coach2={len(c2_props)} "
      f"unplaced(mgr/c1/c2)={len(mgr_unplaced)}/{len(c1_unplaced)}/{len(c2_unplaced)}")

# ---- 5. has_conflicts per owner -> should be clean ------------------------
for owner, (props, _) in plans.items():
    issues = S.has_conflicts(props, busy_by_owner.get(owner, []))
    assert not issues, (owner, issues)
print("OK: no conflicts in any of the 3 calendars")

# ---- 6. write_proposals for coach-owned only, then simulate Flow B -------
coach_props_all = c1_props + c2_props
path, ids = B.write_proposals(cfg, coach_props_all)
written = json.load(open(path))
assert len(written) == len(ids) == len(coach_props_all)
assert written[0]["owner"] in ("coach1@foundever.com", "coach2@foundever.com")
print(f"OK: write_proposals -> {len(ids)} proposals written, sample id={ids[0]}")

# Pretend Flow B created everything successfully:
results_payload = [{"id": i, "status": "created"} for i in ids]
with open(os.path.join(drop, "results_2026-07.json"), "w") as f:
    json.dump(results_payload, f)

results = B.wait_for_results(cfg, cfg["coaches"], ids, timeout_seconds=5, poll_seconds=1)
assert len(results) == len(ids)
assert all(r["status"] == "created" for r in results)
print(f"OK: wait_for_results -> {len(results)}/{len(ids)} confirmed created")

# ---- 7. fallback path file naming -----------------------------------------
cfg_fb = dict(cfg)
cfg_fb["write_path"] = "fallback"
with open(os.path.join(drop, "busy_2026-07_coach1.json"), "w") as f:
    json.dump([["2026-07-06T09:00:00", "2026-07-06T11:00:00"]], f)
with open(os.path.join(drop, "busy_2026-07_coach2.json"), "w") as f:
    json.dump([], f)
fb_busy = B.read_busy(cfg_fb, cfg["coaches"])
assert fb_busy["coach1@foundever.com"][0][0] == datetime(2026, 7, 6, 9, 0)
print("OK: fallback per-coach busy file naming works ->", {k: len(v) for k, v in fb_busy.items()})

shutil.rmtree(tmp)
print("\nALL SYNTHETIC TESTS PASSED")
