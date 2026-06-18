"""
app.py
======
The double-click entry point. A self-contained Tkinter GUI (standard library
only, so PyInstaller bundles it into one .exe with no extra runtime).

Flow:
  1. Adjust staff count / month / time window / work days.
  2. "Generate Preview" reads the live Outlook calendar, builds a plan, and
     shows every proposed meeting in an editable table.
  3. Double-click any row to change its time or delete it; the conflict light
     re-checks instantly.
  4. "Commit to Outlook" writes them, or "Export .ics" saves a file to import.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, timedelta

import scheduler as S
import outlook_client as OC

HERE = os.path.dirname(os.path.abspath(sys.argv[0]))
CONFIG_PATH = os.path.join(HERE, "config.json")

DEFAULT_CONFIG = {
    "month": datetime.now().strftime("%Y-%m"),
    "staff": [],
    "organizer": "",
    "work_days": [0, 1, 2, 3, 4],
    "time_window": {"start": "09:00", "end": "17:00"},
    "slot_granularity_minutes": 30,
    "min_gap_minutes": 15,
    "meetings": [
        {"name": "Focus KPI Meeting",       "category": "FS", "schedule_type": 4, "duration_minutes": 60, "count": 1},
        {"name": "CPA",                      "category": "DS", "schedule_type": 4, "duration_minutes": 30, "per_staff": 1},
        {"name": "Initial Agent SKEP",       "category": "DS", "schedule_type": 1, "duration_minutes": 30, "per_staff": 1},
        {"name": "1st Agent SKEP follow up", "category": "DS", "schedule_type": 2, "duration_minutes": 30, "per_staff": 1},
        {"name": "2nd Agent SKEP follow up", "category": "DS", "schedule_type": 3, "duration_minutes": 30, "per_staff": 1},
        {"name": "Initial Coach SKEP",       "category": "DS", "schedule_type": 2, "duration_minutes": 45, "per_staff": 1},
        {"name": "Coach SKEP follow up",     "category": "DS", "schedule_type": 3, "duration_minutes": 45, "per_staff": 1},
        {"name": "Coach SKEP Audits",        "category": "DS", "schedule_type": 5, "duration_minutes": 30, "per_staff": 1},
    ],
}


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Monthly Meeting Scheduler")
        self.geometry("980x620")
        self.cfg = load_config()
        self.staff: list[str] = S.staff_emails(self.cfg)
        self.proposals: list[S.Proposal] = []
        self.existing = []
        self._build_ui()

    # ----------------------------------------------------------------- UI
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Month (YYYY-MM)").grid(row=0, column=0, sticky="w")
        self.month_var = tk.StringVar(value=self.cfg["month"])
        ttk.Entry(top, textvariable=self.month_var, width=10).grid(row=0, column=1, padx=6)

        self.staff_btn = ttk.Button(top, text=self._staff_btn_label(),
                                     command=self.edit_staff)
        self.staff_btn.grid(row=0, column=2, columnspan=2, padx=6)

        ttk.Label(top, text="Day window").grid(row=0, column=4, sticky="w")
        self.start_var = tk.StringVar(value=self.cfg["time_window"]["start"])
        self.end_var = tk.StringVar(value=self.cfg["time_window"]["end"])
        ttk.Entry(top, textvariable=self.start_var, width=6).grid(row=0, column=5)
        ttk.Label(top, text="to").grid(row=0, column=6)
        ttk.Entry(top, textvariable=self.end_var, width=6).grid(row=0, column=7, padx=6)

        ttk.Label(top, text="Include weekends").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.weekend_var = tk.BooleanVar(value=5 in self.cfg.get("work_days", []) or 6 in self.cfg.get("work_days", []))
        ttk.Checkbutton(top, variable=self.weekend_var).grid(row=1, column=1, sticky="w", pady=(8, 0))

        ttk.Button(top, text="Generate Preview", command=self.generate).grid(row=1, column=3, pady=(8, 0))
        ttk.Button(top, text="Edit meeting types…", command=self.edit_meetings).grid(row=1, column=5, columnspan=2, pady=(8, 0))

        # status / conflict light
        self.status = tk.StringVar(value="Set your options, then Generate Preview.")
        bar = ttk.Frame(self, padding=(10, 0))
        bar.pack(fill="x")
        ttk.Label(bar, textvariable=self.status, foreground="#555").pack(side="left")

        # table
        cols = ("date", "start", "end", "dur", "name", "cat", "type", "invitee")
        headers = ("Date", "Start", "End", "Dur", "Meeting", "Cat", "Type", "Invitee")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="extended")
        for c, h in zip(cols, headers):
            self.tree.heading(c, text=h)
            if c == "name":
                w = 240
            elif c == "invitee":
                w = 200
            else:
                w = 70
            self.tree.column(c, width=w, anchor="w" if c in ("name", "invitee") else "center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree.bind("<Double-1>", self.edit_row)

        # bottom buttons
        bot = ttk.Frame(self, padding=10)
        bot.pack(fill="x")
        ttk.Button(bot, text="Delete selected", command=self.delete_selected).pack(side="left")
        ttk.Button(bot, text="Re-balance", command=self.generate).pack(side="left", padx=6)
        ttk.Button(bot, text="Export .ics", command=self.export_ics).pack(side="right")
        ttk.Button(bot, text="Commit to Outlook", command=self.commit).pack(side="right", padx=6)

    # ---------------------------------------------------------------- logic
    def _staff_btn_label(self):
        return f"Staff emails ({len(self.staff)})…"

    def _gather_cfg(self):
        wd = [0, 1, 2, 3, 4] + ([5, 6] if self.weekend_var.get() else [])
        self.cfg["month"] = self.month_var.get().strip()
        self.cfg["staff"] = list(self.staff)
        self.cfg.pop("staff_count", None)
        self.cfg["work_days"] = wd
        self.cfg["time_window"] = {"start": self.start_var.get().strip(),
                                   "end": self.end_var.get().strip()}
        save_config(self.cfg)
        return self.cfg

    def edit_staff(self):
        """Edit the list of direct-report email addresses (one per line)."""
        win = tk.Toplevel(self)
        win.title("Staff emails — one per line")
        win.geometry("420x360")
        win.transient(self)
        ttk.Label(win, padding=10,
                  text="Enter each direct report's email on its own line.\n"
                       "Each dynamic meeting is created once per person and\n"
                       "sent to them as an invitation.").pack(anchor="w")
        txt = tk.Text(win, height=14)
        txt.pack(fill="both", expand=True, padx=10)
        txt.insert("1.0", "\n".join(self.staff))

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
            self.staff = emails + bad
            self.staff_btn.config(text=self._staff_btn_label())
            self._gather_cfg()
            win.destroy()

        bf = ttk.Frame(win, padding=10)
        bf.pack(fill="x")
        ttk.Button(bf, text="Save", command=save_and_close).pack(side="right")
        ttk.Button(bf, text="Cancel", command=win.destroy).pack(side="right", padx=6)

    def generate(self):
        cfg = self._gather_cfg()
        try:
            year, month = (int(x) for x in cfg["month"].split("-"))
        except Exception:
            messagebox.showerror("Bad month", "Use the format YYYY-MM, e.g. 2026-07.")
            return

        if OC.com_available():
            try:
                self.existing = OC.get_busy(year, month)
            except Exception as e:
                self.existing = []
                messagebox.showwarning("Outlook read failed",
                    f"Could not read the calendar ({e}). Planning around an empty calendar.")
        else:
            self.existing = []
            self.status.set("Outlook COM not found — preview only. Use Export .ics to import.")

        if not self.staff:
            messagebox.showwarning("No staff emails",
                "No staff emails are set, so only fixed meetings will be created. "
                "Click 'Staff emails…' to add your direct reports.")

        self.proposals, unplaced = S.build_plan(cfg, self.existing)
        self._refresh_table()
        msg = f"{len(self.proposals)} meetings proposed."
        if unplaced:
            msg += f"  ⚠ {sum(n for _, n in unplaced)} could not fit (no free slots)."
        issues = S.has_conflicts(self.proposals, self.existing)
        msg += "  No conflicts." if not issues else f"  ⚠ {len(issues)} conflicts!"
        self.status.set(msg)

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        for p in self.proposals:
            self.tree.insert("", "end", values=p.as_row())

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
            f"{p.name}\nEnter new start as YYYY-MM-DD HH:MM",
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
        self.proposals.sort(key=lambda x: x.start)
        self._refresh_table()
        issues = S.has_conflicts(self.proposals, self.existing)
        self.status.set("No conflicts." if not issues else f"⚠ {len(issues)} conflicts after edit!")

    def edit_meetings(self):
        # Lightweight editor: opens config.json location for manual editing.
        messagebox.showinfo(
            "Meeting types",
            "Meeting names, durations, schedule types and per-staff counts "
            "live in config.json next to this program:\n\n" + CONFIG_PATH +
            "\n\nEdit it, then click Generate Preview again.")

    def commit(self):
        if not self.proposals:
            return
        if not OC.com_available():
            messagebox.showinfo("Outlook not available",
                "Classic Outlook COM isn't available here. Use Export .ics instead.")
            return
        issues = S.has_conflicts(self.proposals, self.existing)
        if issues and not messagebox.askyesno(
                "Conflicts present", f"{len(issues)} conflicts detected. Commit anyway?"):
            return
        with_invites = sum(1 for p in self.proposals if p.attendees)
        if not messagebox.askyesno("Confirm",
                f"Add {len(self.proposals)} meetings to your Outlook calendar?\n\n"
                f"{with_invites} of them will send invitations to staff "
                f"right away. This cannot be undone in bulk."):
            return
        try:
            created, invited = OC.commit(self.proposals)
            messagebox.showinfo(
                "Done",
                f"Created {created} items in Outlook.\n"
                f"{invited} invitations sent to staff.")
        except Exception as e:
            messagebox.showerror("Commit failed", str(e))

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
                "an .ics adds the meetings to your calendar but does NOT send "
                "invitations. To actually invite staff, use Commit to Outlook.")


if __name__ == "__main__":
    App().mainloop()
