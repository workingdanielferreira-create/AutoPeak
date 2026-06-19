"""
app.py
======
The double-click entry point. A self-contained Tkinter GUI (standard library
only, plus the local bridge.py, so PyInstaller still bundles it cleanly).

Flow (agents/coaches version):
  1. "Coaches…" lists which coaches exist (agents themselves live in the
     shared CoachAgentMapping.xlsx, not typed in here).
  2. "Fetch availability" reads the manager's own calendar locally (COM) and
     asks the cloud side (Power Automate) for every coach's busy time via the
     Drop-folder bridge, then waits briefly for the answer.
  3. "Generate Preview" builds one plan per calendar owner (manager + each
     coach) and shows them together in one table with a Host column.
  4. Double-click any row to retime it; multi-select + Delete to remove rows.
  5. "Publish" sends manager-owned meetings straight to Outlook via COM, and
     writes coach-owned meetings to proposals_<month>.json for the flow to
     create on the right coach calendar. "Check publish results" polls for
     what actually landed.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, timedelta

import scheduler as S
import outlook_client as OC
import bridge as B

HERE = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG_PATH = os.path.join(HERE, "config.json")

DEFAULT_CONFIG = {
    "month": datetime.now().strftime("%Y-%m"),
    "organizer": "",
    "coaches": [],
    "agents_by_coach": {},
    "agent_mapping_xlsx": "CoachAgentMapping.xlsx",
    "drop_folder": "",
    "write_path": "primary",
    "work_days": [0, 1, 2, 3, 4],
    "time_window": {"start": "09:00", "end": "17:00"},
    "slot_granularity_minutes": 30,
    "min_gap_minutes": 15,
    "meetings": [
        {"name": "Focus KPI Meeting",       "category": "FS", "population": "manager", "schedule_type": 4, "duration_minutes": 60, "count": 4, "invite_all_staff": True},
        {"name": "CPA",                      "category": "DS", "population": "agents",  "schedule_type": 4, "duration_minutes": 30, "per_staff": 1},
        {"name": "Initial Agent SKEP",       "category": "DS", "population": "agents",  "schedule_type": 1, "duration_minutes": 30, "per_staff": 1},
        {"name": "1st Agent SKEP follow up", "category": "DS", "population": "agents",  "schedule_type": 2, "duration_minutes": 30, "per_staff": 1},
        {"name": "2nd Agent SKEP follow up", "category": "DS", "population": "agents",  "schedule_type": 3, "duration_minutes": 30, "per_staff": 1},
        {"name": "Initial Coach SKEP",       "category": "DS", "population": "coaches", "schedule_type": 2, "duration_minutes": 45, "per_staff": 1},
        {"name": "Coach SKEP follow up",     "category": "DS", "population": "coaches", "schedule_type": 3, "duration_minutes": 45, "per_staff": 1},
        {"name": "Coach SKEP Audits",        "category": "DS", "population": "coaches", "schedule_type": 5, "duration_minutes": 30, "per_staff": 1},
    ],
}


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AutoPeak — Monthly Meeting Scheduler")
        self.geometry("1080x640")
        self.cfg = load_config()
        self.coaches: list[str] = list(self.cfg.get("coaches", []))
        self._refresh_mapping(silent=True)

        self.proposals: list[S.Proposal] = []
        self.mgr_busy: list[S.Interval] = []
        self.coach_busy: dict[str, list[tuple[datetime, datetime]]] = {}
        self._last_ids: list[str] = []
        self._build_ui()

    # ----------------------------------------------------------------- UI
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Month (YYYY-MM)").grid(row=0, column=0, sticky="w")
        self.month_var = tk.StringVar(value=self.cfg["month"])
        ttk.Entry(top, textvariable=self.month_var, width=10).grid(row=0, column=1, padx=6)

        self.coach_btn = ttk.Button(top, text=self._coach_btn_label(),
                                     command=self.edit_coaches)
        self.coach_btn.grid(row=0, column=2, columnspan=2, padx=6)

        ttk.Label(top, text="Day window").grid(row=0, column=4, sticky="w")
        self.start_var = tk.StringVar(value=self.cfg["time_window"]["start"])
        self.end_var = tk.StringVar(value=self.cfg["time_window"]["end"])
        ttk.Entry(top, textvariable=self.start_var, width=6).grid(row=0, column=5)
        ttk.Label(top, text="to").grid(row=0, column=6)
        ttk.Entry(top, textvariable=self.end_var, width=6).grid(row=0, column=7, padx=6)

        ttk.Label(top, text="Include weekends").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.weekend_var = tk.BooleanVar(value=5 in self.cfg.get("work_days", []) or 6 in self.cfg.get("work_days", []))
        ttk.Checkbutton(top, variable=self.weekend_var).grid(row=1, column=1, sticky="w", pady=(8, 0))

        ttk.Button(top, text="Fetch availability", command=self.fetch_availability).grid(row=1, column=2, pady=(8, 0))
        ttk.Button(top, text="Generate Preview", command=self.generate).grid(row=1, column=3, pady=(8, 0))
        ttk.Button(top, text="Edit meeting types…", command=self.edit_meetings).grid(row=1, column=5, columnspan=2, pady=(8, 0))

        self.mapping_status = tk.StringVar(value=self._mapping_status_text())
        ttk.Label(top, textvariable=self.mapping_status, foreground="#777").grid(
            row=2, column=0, columnspan=8, sticky="w", pady=(6, 0))

        # status / conflict light
        self.status = tk.StringVar(value="Set your options, then Fetch availability, then Generate Preview.")
        bar = ttk.Frame(self, padding=(10, 0))
        bar.pack(fill="x")
        ttk.Label(bar, textvariable=self.status, foreground="#555").pack(side="left")

        # table — Host column first, matches Proposal.as_row()
        cols = ("host", "date", "start", "end", "dur", "name", "cat", "type", "invitee")
        headers = ("Host calendar", "Date", "Start", "End", "Dur", "Meeting", "Cat", "Type", "Invitee")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="extended")
        for c, h in zip(cols, headers):
            self.tree.heading(c, text=h)
            if c == "host":
                w = 200
            elif c == "name":
                w = 220
            elif c == "invitee":
                w = 200
            else:
                w = 70
            self.tree.column(c, width=w, anchor="w" if c in ("host", "name", "invitee") else "center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree.bind("<Double-1>", self.edit_row)

        # bottom buttons
        bot = ttk.Frame(self, padding=10)
        bot.pack(fill="x")
        ttk.Button(bot, text="Delete selected", command=self.delete_selected).pack(side="left")
        ttk.Button(bot, text="Re-balance", command=self.generate).pack(side="left", padx=6)
        ttk.Button(bot, text="Export .ics", command=self.export_ics).pack(side="right")
        ttk.Button(bot, text="Check publish results", command=self.check_results).pack(side="right", padx=6)
        ttk.Button(bot, text="Publish", command=self.publish).pack(side="right", padx=6)

    # ---------------------------------------------------------------- mapping
    def _coach_btn_label(self):
        return f"Coaches ({len(self.coaches)})…"

    def _mapping_status_text(self):
        n_coaches = len(self.cfg.get("agents_by_coach", {}))
        n_agents = sum(len(v) for v in self.cfg.get("agents_by_coach", {}).values())
        path = self.cfg.get("agent_mapping_xlsx", "CoachAgentMapping.xlsx")
        if n_coaches:
            return f"Mapping: {n_coaches} coach(es) / {n_agents} agent(s) loaded from {path}"
        return f"Mapping: no rows loaded yet from {path} — agent meetings won't be placed until it has data."

    def _refresh_mapping(self, silent: bool = False):
        try:
            B.merge_mapping_into_cfg(self.cfg)
            self.coaches = list(self.cfg.get("coaches", []))
        except B.BridgeError as e:
            if not silent:
                messagebox.showwarning("Mapping not loaded", str(e))

    # ---------------------------------------------------------------- logic
    def _gather_cfg(self):
        wd = [0, 1, 2, 3, 4] + ([5, 6] if self.weekend_var.get() else [])
        self.cfg["month"] = self.month_var.get().strip()
        self.cfg["coaches"] = list(self.coaches)
        self.cfg["work_days"] = wd
        self.cfg["time_window"] = {"start": self.start_var.get().strip(),
                                   "end": self.end_var.get().strip()}
        save_config(self.cfg)
        return self.cfg

    def edit_coaches(self):
        """Edit the list of coach email addresses (one per line). Agents are
        NOT entered here — they come from the shared CoachAgentMapping.xlsx."""
        win = tk.Toplevel(self)
        win.title("Coach emails — one per line")
        win.geometry("440x380")
        win.transient(self)
        ttk.Label(win, padding=10,
                  text="Enter each coach's email on its own line.\n"
                       "Coach-level meetings (Coach SKEPs, etc.) are created\n"
                       "once per coach on the manager's calendar.\n\n"
                       "Each coach's AGENTS come from the shared\n"
                       "CoachAgentMapping.xlsx — not from here.").pack(anchor="w")
        txt = tk.Text(win, height=12)
        txt.pack(fill="both", expand=True, padx=10)
        txt.insert("1.0", "\n".join(self.coaches))

        def save_and_close():
            raw = [l.strip() for l in txt.get("1.0", "end").splitlines()]
            emails, bad = [], []
            for e in raw:
                if not e:
                    continue
                (emails if "@" in e and "." in e.split("@")[-1] else bad).append(e)
            if bad and not messagebox.askyesno(
                    "Check addresses",
                    "These don't look like emails:\n\n" + "\n".join(bad) +
                    "\n\nKeep them anyway?"):
                return
            self.coaches = emails + bad
            self.coach_btn.config(text=self._coach_btn_label())
            self._gather_cfg()
            win.destroy()

        bf = ttk.Frame(win, padding=10)
        bf.pack(fill="x")
        ttk.Button(bf, text="Save", command=save_and_close).pack(side="right")
        ttk.Button(bf, text="Cancel", command=win.destroy).pack(side="right", padx=6)

    def fetch_availability(self):
        cfg = self._gather_cfg()
        self._refresh_mapping()
        self.mapping_status.set(self._mapping_status_text())
        save_config(cfg)

        try:
            year, month = (int(x) for x in cfg["month"].split("-"))
        except Exception:
            messagebox.showerror("Bad month", "Use the format YYYY-MM, e.g. 2026-07.")
            return

        # Manager's own busy — local, instant, via COM.
        if OC.com_available():
            try:
                self.mgr_busy = OC.get_busy(year, month)
            except Exception as e:
                self.mgr_busy = []
                messagebox.showwarning("Outlook read failed",
                    f"Could not read your calendar ({e}). Planning around an empty calendar.")
        else:
            self.mgr_busy = []
            messagebox.showwarning("Outlook COM not found",
                "Classic desktop Outlook COM isn't available here, so your own "
                "calendar will be treated as empty for planning purposes.")

        if not self.coaches:
            messagebox.showwarning("No coaches", "Add at least one coach first.")
            return
        if not cfg.get("drop_folder"):
            messagebox.showerror("No Drop folder set",
                "config.json's 'drop_folder' is empty. Point it at your "
                "OneDrive-synced AutoPeak Drop folder before fetching availability.")
            return

        owners = S.all_owners(cfg)
        try:
            B.write_request(cfg, owners)
        except B.BridgeError as e:
            messagebox.showerror("Couldn't write request", str(e))
            return

        self.status.set(f"Request sent for {len(self.coaches)} coach calendar(s) — waiting on the flow…")
        self.update_idletasks()
        try:
            self.coach_busy = B.wait_for_busy(cfg, self.coaches, timeout_seconds=45, poll_seconds=3)
            self.status.set(f"Got availability for {len(self.coach_busy)} coach calendar(s). Ready to Generate Preview.")
        except B.BridgeError as e:
            self.coach_busy = {}
            messagebox.showwarning("Still waiting",
                f"{e}\n\nThe request file was written. If this is the first run, check "
                "the flow's run history in Power Automate, then click Fetch availability again.")
            self.status.set("Availability not back yet — try Fetch availability again shortly.")

    def generate(self):
        cfg = self._gather_cfg()
        try:
            year, month = (int(x) for x in cfg["month"].split("-"))
        except Exception:
            messagebox.showerror("Bad month", "Use the format YYYY-MM, e.g. 2026-07.")
            return

        if not self.coaches:
            messagebox.showwarning("No coaches",
                "No coaches are set, so only manager-population meetings will be placed.")

        busy_by_owner: dict[str, list[S.Interval]] = {}
        organizer = S.organizer(cfg)
        if organizer:
            busy_by_owner[organizer] = list(self.mgr_busy)
        for coach in self.coaches:
            pairs = self.coach_busy.get(coach, [])
            busy_by_owner[coach] = [S.Interval(s, e) for s, e in pairs]
        self.busy_by_owner = busy_by_owner

        plans = S.build_all_plans(cfg, busy_by_owner)
        self.proposals = S.flatten_proposals(plans)
        self._refresh_table()

        unplaced_total = sum(len(u) for _, u in plans.values())
        issues_total = sum(len(S.has_conflicts(props, busy_by_owner.get(owner, [])))
                            for owner, (props, _) in plans.items())
        msg = f"{len(self.proposals)} meetings proposed across {len(plans)} calendar(s)."
        if unplaced_total:
            msg += f"  ⚠ {unplaced_total} item(s) could not fit."
        msg += "  No conflicts." if not issues_total else f"  ⚠ {issues_total} conflicts!"
        self.status.set(msg)

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for p in self.proposals:
            self.tree.insert("", "end", values=p.as_row())

    def _grouped_conflicts(self) -> int:
        groups: dict[str, list[S.Proposal]] = {}
        for p in self.proposals:
            groups.setdefault(p.owner, []).append(p)
        total = 0
        for owner, props in groups.items():
            total += len(S.has_conflicts(props, self.busy_by_owner.get(owner, [])))
        return total

    def delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idxs = sorted((self.tree.index(i) for i in sel), reverse=True)
        for i in idxs:
            del self.proposals[i]
        self._refresh_table()
        self.status.set(f"{len(self.proposals)} meetings remaining.")

    def edit_row(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        p = self.proposals[idx]
        new = simpledialog.askstring(
            "Edit time",
            f"{p.name}  (host: {p.owner})\nEnter new start as YYYY-MM-DD HH:MM",
            initialvalue=p.start.strftime("%Y-%m-%d %H:%M"))
        if not new:
            return
        try:
            start = datetime.strptime(new.strip(), "%Y-%m-%d %H:%M")
        except Exception:
            messagebox.showerror("Bad time", "Use YYYY-MM-DD HH:MM")
            return
        dur = int((p.end - p.start).total_seconds() // 60)
        p.start = start
        p.end = start + timedelta(minutes=dur)
        self.proposals.sort(key=lambda x: (x.start, x.owner))
        self._refresh_table()
        issues = self._grouped_conflicts()
        self.status.set("No conflicts." if not issues else f"⚠ {issues} conflicts after edit!")

    def edit_meetings(self):
        messagebox.showinfo(
            "Meeting types",
            "Meeting names, durations, schedule types, per-staff counts and "
            "population ('manager' | 'coaches' | 'agents') live in config.json "
            "next to this program:\n\n" + CONFIG_PATH +
            "\n\nEdit it, then click Generate Preview again.")

    # ---------------------------------------------------------------- publish
    def publish(self):
        if not self.proposals:
            return
        organizer = S.organizer(self.cfg)
        mgr_props = [p for p in self.proposals if p.owner == organizer]
        coach_props = [p for p in self.proposals if p.owner != organizer]

        issues = self._grouped_conflicts()
        if issues and not messagebox.askyesno(
                "Conflicts present", f"{issues} conflicts detected. Publish anyway?"):
            return

        mgr_invites = sum(1 for p in mgr_props if p.attendees)
        if not messagebox.askyesno("Confirm Publish",
                f"{len(mgr_props)} meeting(s) go straight onto YOUR calendar "
                f"({mgr_invites} send invitations immediately).\n\n"
                f"{len(coach_props)} meeting(s) will be created on "
                f"{len({p.owner for p in coach_props})} coach calendar(s) via "
                f"the Power Automate flow, with invites to their agents.\n\n"
                "This cannot be undone in bulk. Continue?"):
            return

        created = invited = 0
        if mgr_props:
            if not OC.com_available():
                messagebox.showwarning("Outlook not available",
                    "Classic Outlook COM isn't available here — manager-owned "
                    "meetings were NOT created. Coach-owned meetings will still "
                    "be sent to the flow below.")
            else:
                try:
                    created, invited = OC.commit(mgr_props)
                except Exception as e:
                    messagebox.showerror("Manager commit failed", str(e))

        self._last_ids = []
        if coach_props:
            try:
                path, ids = B.write_proposals(self.cfg, coach_props)
                self._last_ids = ids
                self.status.set(
                    f"Manager: {created} created, {invited} invited. "
                    f"Wrote {len(ids)} coach-owned proposal(s) to {os.path.basename(path)} — "
                    "click 'Check publish results' once the flow has run.")
            except B.BridgeError as e:
                messagebox.showerror("Couldn't write coach proposals", str(e))
                return
        else:
            self.status.set(f"Manager: {created} created, {invited} invited. No coach-owned meetings to publish.")

        messagebox.showinfo("Published",
            f"Manager calendar: {created} created, {invited} invitations sent.\n"
            f"Coach calendars: {len(coach_props)} proposal(s) handed off to the flow."
            + ("\n\nUse 'Check publish results' to confirm they landed." if coach_props else ""))

    def check_results(self):
        if not self._last_ids:
            messagebox.showinfo("Nothing to check", "Publish coach-owned meetings first.")
            return
        cfg = self.cfg
        results = B.wait_for_results(cfg, self.coaches, self._last_ids, timeout_seconds=30, poll_seconds=3)
        by_id = {r["id"]: r for r in results}
        created = [i for i in self._last_ids if by_id.get(i, {}).get("status") == "created"]
        failed = [(i, by_id[i].get("error", "unknown error")) for i in self._last_ids
                  if by_id.get(i, {}).get("status") == "failed"]
        pending = [i for i in self._last_ids if i not in by_id]

        lines = [f"Created: {len(created)}", f"Pending: {len(pending)}", f"Failed: {len(failed)}"]
        if failed:
            lines.append("")
            lines.extend(f"  {i}: {err}" for i, err in failed)
        messagebox.showinfo("Publish results", "\n".join(lines))
        if pending:
            self.status.set(f"{len(pending)} coach-owned meeting(s) still pending — check again shortly.")

    def export_ics(self):
        if not self.proposals:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".ics", filetypes=[("iCalendar", "*.ics")],
            initialfile=f"schedule_{self.cfg['month']}.ics")
        if path:
            OC.export_ics(self.proposals, path, organizer=self.cfg.get("organizer") or None)
            messagebox.showinfo("Exported",
                f"Saved {path}\n\nImport via File ▸ Open ▸ Import. Note: importing "
                "an .ics adds the meetings to the importer's own calendar but does "
                "NOT send invitations. For real invitations, use Publish.")


if __name__ == "__main__":
    App().mainloop()
