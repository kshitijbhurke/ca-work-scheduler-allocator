"""
Work Scheduler
--------------
Desktop app (Tkinter) with cascading dropdowns driven by 'Work Schedule Master.xlsx':

    1. Client (Client List)  ->  shows Client Short Name
    2. Assignment (Assignment Master)
    3. Work Type / Area (Work Type Master, filtered by Assignment)
    4. Frequency (Work Type Master, filtered by Assignment + Work Type / Area)
    5. Period Covered (From / To dates, picked from a calendar)

Scheduled entries are saved to 'Work Schedule.xlsx' (sheet 'Scheduler') in the same folder,
so the master workbook is only ever read, never modified.
"""

import calendar
import os
import sys
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, font as tkfont, ttk, messagebox

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE_DIR = Path(__file__).resolve().parent
MASTER_FILE = BASE_DIR / "Work Schedule Master.xlsx"
SCHEDULE_FILE = BASE_DIR / "Work Schedule.xlsx"
SCHEDULE_SHEET = "Scheduler"

SCHEDULE_COLUMNS = [
    "Sr No", "Date Added", "Client Full Name", "Client Short Name", "Client Group",
    "Assignment", "Work Type / Area", "Frequency", "Financial Year",
    "Period From", "Period To",
    "Manager", "Manager Short Name", "Due Date", "Target Date", "Work Status",
]
DATE_COLUMNS = ("Period From", "Period To", "Due Date", "Target Date")
WORK_STATUSES = ("WIP", "Complete")   # blank = work not started
CENTERED_COLUMNS = ("Sr No", "Frequency", "Financial Year", "Period From", "Period To",
                    "Due Date", "Target Date", "Work Status")
COLUMN_WIDTHS = {
    "Sr No": 7, "Date Added": 18, "Client Full Name": 50, "Client Short Name": 16,
    "Client Group": 24, "Assignment": 28, "Work Type / Area": 48, "Frequency": 18, "Financial Year": 14,
    "Period From": 14, "Period To": 14, "Manager": 22, "Manager Short Name": 16, "Due Date": 14, "Target Date": 14,
    "Work Status": 14,
}
NOT_STARTED = "Not started (blank)"
NO_MANAGER = "— None —"
DATE_FMT = "%d-%b-%Y"            # 01-Apr-2026
OLD_DATE_FMTS = ("%d-%m-%Y",)    # accepted when reading older files
NA = "NA"
CHECK_COL, CHECKED, UNCHECKED = "Select", "☑", "☐"

# Frequencies that are split across a Financial Year (April–March): months per line item.
FY_FREQUENCIES = {
    "monthly": 1,
    "quarterly": 3,
    "half yearly": 6, "halfyearly": 6, "six monthly": 6, "semi annual": 6,
    "annual": 12, "yearly": 12,
}
NA_FREQUENCIES = {"na", "n/a", "n.a.", "not applicable"}

# ---- palette ---------------------------------------------------------------- #
BG = "#EEF2F7"
CARD = "#FFFFFF"
NAVY = "#0F2A47"
NAVY_2 = "#1B3E66"
ON_NAVY_MUTED = "#A9C1DE"
ACCENT = "#2563EB"
ACCENT_DARK = "#1D4ED8"
ACCENT_SOFT = "#DBEAFE"
HOVER_SOFT = "#EFF6FF"
SUCCESS = "#16A34A"
SUCCESS_DARK = "#15803D"
AMBER = "#F59E0B"
AMBER_SOFT = "#FEF3C7"
AMBER_DARK = "#92400E"
WARN = "#B45309"
DANGER = "#DC2626"
TEXT = "#0F172A"
MUTED = "#64748B"
FAINT = "#CBD5E1"
BORDER = "#D7DEE8"
ZEBRA = "#F7F9FC"
FONT = "Segoe UI"


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _clean(value):
    """Normalise a cell value: strip stray spaces; blank -> None."""
    if value is None:
        return None
    text = " ".join(str(value).split())
    return text or None


def to_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in (DATE_FMT, *OLD_DATE_FMTS):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def fmt_cell(value):
    if isinstance(value, (date, datetime)):
        return to_date(value).strftime(DATE_FMT)
    return "" if value is None else value


def frequency_kind(freq):
    """Classify a Frequency from the master.

    Returns ("fy", months_per_item) for Monthly / Quarterly / Half Yearly / Annual,
    ("na", None) for NA, and ("manual", None) for everything else
    (e.g. 'Annual Dec End', 'As per Engagement', 'Pre Decided').
    """
    f = " ".join(str(freq or "").lower().replace("-", " ").split())
    if f in NA_FREQUENCIES:
        return "na", None
    if f in FY_FREQUENCIES:
        return "fy", FY_FREQUENCIES[f]
    return "manual", None


def fy_label(start_year):
    return f"FY {start_year}-{str(start_year + 1)[-2:]}"


def fy_start_year(label):
    return int(label.split()[1][:4])


def current_fy_start(today=None):
    today = today or date.today()
    return today.year if today.month >= 4 else today.year - 1


def fy_periods(start_year, months):
    """Split FY start_year (1 Apr start_year – 31 Mar next year) into equal periods."""
    periods = []
    year, month = start_year, 4
    for _ in range(12 // months):
        end_index = (month - 1) + (months - 1)
        end_year, end_month = year + end_index // 12, end_index % 12 + 1
        periods.append((date(year, month, 1),
                        date(end_year, end_month, calendar.monthrange(end_year, end_month)[1])))
        month += months
        if month > 12:
            month -= 12
            year += 1
    return periods


def _open_workbook_shared(path):
    """Open a workbook read-only, even while it is open (locked) in Excel."""
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        if sys.platform != "win32":
            raise
        # Excel holds the file locked; Windows' CopyFile can still copy it.
        import ctypes
        import tempfile
        tmp = Path(tempfile.gettempdir()) / f"work_scheduler_{path.stem}.xlsx"
        if not ctypes.windll.kernel32.CopyFileW(str(path), str(tmp), False):
            raise
        return openpyxl.load_workbook(tmp, read_only=True, data_only=True)


def _read_table(ws, required_headers):
    """Locate the header row anywhere on the sheet and return rows as dicts.

    The master sheets don't all start at A1 (headers sit in row 2-4, from column B),
    so we scan for the row that contains every required header.
    """
    rows = list(ws.iter_rows(values_only=True))
    for idx, row in enumerate(rows):
        cleaned = [_clean(v) for v in row]
        if all(h in cleaned for h in required_headers):
            col = {h: cleaned.index(h) for h in required_headers}
            out = []
            for data_row in rows[idx + 1:]:
                record = {h: _clean(data_row[c]) if c < len(data_row) else None
                          for h, c in col.items()}
                if any(record.values()):
                    out.append(record)
            return out
    raise ValueError(
        f"Sheet '{ws.title}' has no header row with: {', '.join(required_headers)}")


def _unique(seq):
    """De-duplicate while keeping the order used in the master sheet."""
    seen, out = set(), []
    for item in seq:
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return out


def _read_managers(ws):
    """Return (manager names, {name: short name}); the short name column is optional."""
    for name_col in ("Manager Name", "Manager"):
        for short_col in ("Mgr Short Name", "Manager Short Name", None):
            headers = [name_col] + ([short_col] if short_col else [])
            try:
                rows = _read_table(ws, headers)
            except ValueError:
                continue
            names = _unique(r[name_col] for r in rows)
            shorts = {r[name_col]: r.get(short_col) for r in rows if r[name_col]}
            return names, shorts
    return [], {}


# --------------------------------------------------------------------------- #
# Master data
# --------------------------------------------------------------------------- #
class MasterData:
    def __init__(self, path: Path):
        wb = _open_workbook_shared(path)
        try:
            client_headers = ["Client Full Name", "Client Short Name"]
            try:
                clients = _read_table(wb["Client List"], [*client_headers, "Client Group"])
            except ValueError:   # older workbook without the 'Client Group' column
                clients = _read_table(wb["Client List"], client_headers)
            assignments = _read_table(wb["Assignment Master"], ["Assignment"])
            work_types = _read_table(wb["Work Type Master"], ["Assignment", "Work Type / Area"])
            # Frequencies come from 'Frequency Master' (it lists every frequency per work
            # type); older workbooks without that sheet fall back to 'Work Type Master'.
            self.frequency_sheet = ("Frequency Master" if "Frequency Master" in wb.sheetnames
                                    else "Work Type Master")
            frequencies = _read_table(wb[self.frequency_sheet],
                                      ["Assignment", "Work Type / Area", "Frequency"])
            # Managers come from the 'Manager Master' sheet (column 'Manager Name')
            self.managers = []
            self.manager_short = {}   # Manager Name -> Mgr Short Name
            sheet = next((n for n in ("Manager Master", "Manager List") if n in wb.sheetnames),
                         None)
            if sheet:
                self.managers, self.manager_short = _read_managers(wb[sheet])
        finally:
            wb.close()

        # Client full name -> short name / group
        self.clients = {}
        self.client_groups = {}
        for r in clients:
            if r["Client Full Name"]:
                self.clients.setdefault(r["Client Full Name"], r["Client Short Name"] or "")
                self.client_groups.setdefault(r["Client Full Name"], r.get("Client Group"))

        self.assignments = _unique(r["Assignment"] for r in assignments)
        self.work_type_rows = [r for r in work_types if r["Assignment"] and r["Work Type / Area"]]
        self.frequency_rows = [r for r in frequencies
                               if r["Assignment"] and r["Work Type / Area"] and r["Frequency"]]

    def work_types_for(self, assignment):
        return _unique(r["Work Type / Area"] for r in self.work_type_rows
                       if r["Assignment"] == assignment)

    def frequencies_for(self, assignment, work_type):
        return _unique(r["Frequency"] for r in self.frequency_rows
                       if r["Assignment"] == assignment and r["Work Type / Area"] == work_type)


# --------------------------------------------------------------------------- #
# Schedule storage
# --------------------------------------------------------------------------- #
def load_schedule():
    if not SCHEDULE_FILE.exists():
        return []
    wb = openpyxl.load_workbook(SCHEDULE_FILE, data_only=True)
    try:
        ws = wb[SCHEDULE_SHEET] if SCHEDULE_SHEET in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        return []
    # Read by header name so files saved by older versions still load correctly.
    headers = [_clean(h) for h in rows[0]]
    entries = []
    for row in rows[1:]:
        if any(v is not None for v in row):
            e = {h: v for h, v in zip(headers, row) if h}
            for c in DATE_COLUMNS:
                v = e.get(c)
                e[c] = NA if isinstance(v, str) and v.strip().upper() == NA else to_date(v)
            # Accept 'wip' / 'complete' typed in Excel in any case; anything else = blank
            status = _clean(e.get("Work Status"))
            e["Work Status"] = next((s for s in WORK_STATUSES
                                     if status and s.lower() == status.lower()), None)
            entries.append(e)
    return entries


def write_schedule(path, entries, columns=SCHEDULE_COLUMNS):
    """Write entries to an Excel file, with the given columns in the given order."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SCHEDULE_SHEET
    ws.append(list(columns))
    header_fill = PatternFill("solid", fgColor="0F2A47")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for e in entries:
        ws.append([e.get(c) for c in columns])
    for i, c in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = COLUMN_WIDTHS[c]
        for (cell,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
            if c in DATE_COLUMNS:
                cell.number_format = "DD-MMM-YYYY"
            if c in CENTERED_COLUMNS:
                cell.alignment = Alignment(horizontal="center")
    if "Work Status" in columns:
        # Drop-down when the sheet is edited directly in Excel (blank = not started)
        status_col = get_column_letter(columns.index("Work Status") + 1)
        dv = DataValidation(type="list", formula1=f'"{",".join(WORK_STATUSES)}"',
                            allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{status_col}2:{status_col}{max(len(entries) + 1, 2) + 500}")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def save_schedule(entries):
    for i, e in enumerate(entries, start=1):
        e["Sr No"] = i
    write_schedule(SCHEDULE_FILE, entries)


# --------------------------------------------------------------------------- #
# Widgets
# --------------------------------------------------------------------------- #
class DatePicker(tk.Toplevel):
    """Pop-up month calendar shown under a DateField."""

    def __init__(self, anchor, selected, on_pick, min_date, S):
        super().__init__(anchor)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=BORDER)
        self.S, self.selected, self.on_pick, self.min_date = S, selected, on_pick, min_date
        base = selected or min_date or date.today()
        self.year, self.month = base.year, base.month

        body = tk.Frame(self, bg="white")
        body.pack(padx=1, pady=1)

        head = tk.Frame(body, bg=NAVY)
        head.pack(fill="x")
        for text, delta, side in (("«", -12, "left"), ("‹", -1, "left"),
                                  ("»", 12, "right"), ("›", 1, "right")):
            self._nav(head, text, delta).pack(side=side)
        self.title_lbl = tk.Label(head, bg=NAVY, fg="white", font=(FONT, 11, "bold"))
        self.title_lbl.pack(side="left", expand=True)

        self.days = tk.Frame(body, bg="white", padx=S(8), pady=S(6))
        self.days.pack()

        foot = tk.Frame(body, bg="white")
        foot.pack(fill="x", padx=S(10), pady=(0, S(8)))
        today = date.today()
        if not (min_date and today < min_date):
            self._link(foot, "Today", lambda: self._pick(today)).pack(side="left")
        self._link(foot, "Clear", lambda: self._pick(None), MUTED).pack(side="right")

        self._render()
        self._place_under(anchor)
        self.bind("<Escape>", lambda _e: self.close())
        self.bind("<ButtonPress-1>", self._click_outside, add="+")
        self.grab_set()
        self.focus_force()

    def _nav(self, parent, text, delta):
        lbl = tk.Label(parent, text=text, bg=NAVY, fg="white", font=(FONT, 12, "bold"),
                       padx=self.S(9), pady=self.S(5), cursor="hand2")
        lbl.bind("<Button-1>", lambda _e: self._shift(delta))
        lbl.bind("<Enter>", lambda _e: lbl.configure(bg=NAVY_2))
        lbl.bind("<Leave>", lambda _e: lbl.configure(bg=NAVY))
        return lbl

    @staticmethod
    def _link(parent, text, command, fg=ACCENT):
        lbl = tk.Label(parent, text=text, bg="white", fg=fg, cursor="hand2",
                       font=(FONT, 9, "bold", "underline"))
        lbl.bind("<Button-1>", lambda _e: command())
        return lbl

    def _render(self):
        for w in self.days.winfo_children():
            w.destroy()
        self.title_lbl.configure(text=f"{calendar.month_name[self.month]} {self.year}")
        for i, name in enumerate(("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")):
            tk.Label(self.days, text=name, bg="white", fg=MUTED, width=4,
                     font=(FONT, 9, "bold")).grid(row=0, column=i, pady=(0, self.S(4)))
        today = date.today()
        weeks = calendar.Calendar(firstweekday=0).monthdatescalendar(self.year, self.month)
        for r, week in enumerate(weeks, start=1):
            for c, d in enumerate(week):
                disabled = bool(self.min_date and d < self.min_date)
                in_month = d.month == self.month
                bg, fg, font = "white", TEXT, (FONT, 10)
                if c >= 5:
                    fg = "#C2410C"
                if not in_month:
                    fg = FAINT
                if d == today:
                    fg, font = ACCENT, (FONT, 10, "bold")
                if disabled:
                    fg = "#E2E8F0"
                if d == self.selected:
                    bg, fg, font = ACCENT, "white", (FONT, 10, "bold")
                lbl = tk.Label(self.days, text=d.day, bg=bg, fg=fg, font=font,
                               width=4, pady=self.S(4))
                lbl.grid(row=r, column=c, padx=1, pady=1)
                if disabled:
                    continue
                lbl.configure(cursor="hand2")
                lbl.bind("<Button-1>", lambda _e, d=d: self._pick(d))
                if d != self.selected:
                    lbl.bind("<Enter>", lambda _e, l=lbl: l.configure(bg=HOVER_SOFT))
                    lbl.bind("<Leave>", lambda _e, l=lbl, b=bg: l.configure(bg=b))

    def _place_under(self, anchor):
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        x = anchor.winfo_rootx()
        y = anchor.winfo_rooty() + anchor.winfo_height() + self.S(4)
        if y + h > self.winfo_screenheight() - self.S(48):
            y = anchor.winfo_rooty() - h - self.S(4)
        x = max(0, min(x, self.winfo_screenwidth() - w - self.S(8)))
        self.geometry(f"+{x}+{y}")

    def _shift(self, delta):
        total = self.year * 12 + (self.month - 1) + delta
        self.year, month0 = divmod(total, 12)
        self.month = month0 + 1
        self._render()

    def _click_outside(self, event):
        try:
            x0, y0 = self.winfo_rootx(), self.winfo_rooty()
            inside = (x0 <= event.x_root < x0 + self.winfo_width()
                      and y0 <= event.y_root < y0 + self.winfo_height())
        except tk.TclError:
            return
        if not inside:
            self.close()

    def _pick(self, d):
        self.close()
        self.on_pick(d)

    def close(self):
        try:
            self.grab_release()
            self.destroy()
        except tk.TclError:
            pass


class DateField(tk.Frame):
    """Read-only date box; clicking it opens a DatePicker."""

    def __init__(self, parent, on_change, S, min_date=None):
        super().__init__(parent, bg="white", highlightthickness=1,
                         highlightbackground=BORDER, highlightcolor=BORDER, cursor="hand2")
        self.S, self._on_change, self._min_date = S, on_change, min_date
        self.value = None
        self._picker = None
        self.text = tk.Label(self, bg="white", font=(FONT, 10), anchor="w",
                             padx=S(10), pady=S(7), cursor="hand2")
        self.text.pack(side="left", fill="x", expand=True)
        self.icon = tk.Label(self, text="📅", bg="white", font=("Segoe UI Emoji", 11),
                             padx=S(8), cursor="hand2")
        self.icon.pack(side="right")
        for w in (self, self.text, self.icon):
            w.bind("<Button-1>", self._open)
            w.bind("<Enter>", lambda _e: self.configure(highlightbackground=ACCENT))
            w.bind("<Leave>", lambda _e: self.configure(highlightbackground=BORDER))
        self.set(None, notify=False)

    def set(self, d, notify=True):
        self.value = d
        self.text.configure(text=d.strftime(DATE_FMT) if d else "Select date",
                            fg=TEXT if d else "#94A3B8")
        if notify:
            self._on_change()

    def _open(self, _event=None):
        if self._picker is not None and self._picker.winfo_exists():
            return
        min_d = self._min_date() if self._min_date else None
        self._picker = DatePicker(self, self.value, self.set, min_d, self.S)


class ExportDialog(tk.Toplevel):
    """Pop-up to choose which fields go into the Excel export."""

    def __init__(self, parent, row_count, selected, on_export, S):
        super().__init__(parent, bg=CARD)
        self.title("Export to Excel")
        self.resizable(False, False)
        self.transient(parent)
        self.on_export = on_export

        inner = tk.Frame(self, bg=CARD)
        inner.pack(padx=S(24), pady=S(20))
        tk.Label(inner, text="Export to Excel", bg=CARD, fg=TEXT,
                 font=(FONT, 14, "bold")).pack(anchor="w")
        tk.Label(inner, text=f"Choose the fields to export.  {row_count} "
                             f"row{'' if row_count == 1 else 's'} shown in the Scheduler "
                             "will be exported.",
                 bg=CARD, fg=MUTED, font=(FONT, 9)).pack(anchor="w", pady=(S(2), S(12)))

        grid = tk.Frame(inner, bg=CARD)
        grid.pack(fill="x")
        self.vars = {c: tk.BooleanVar(value=c in selected) for c in SCHEDULE_COLUMNS}
        per_column = (len(SCHEDULE_COLUMNS) + 1) // 2
        for i, (name, var) in enumerate(self.vars.items()):
            tk.Checkbutton(grid, text=name, variable=var, command=self._sync, bg=CARD, fg=TEXT,
                           activebackground=CARD, font=(FONT, 10), anchor="w").grid(
                row=i % per_column, column=i // per_column, sticky="w", padx=(0, S(24)))

        links = tk.Frame(inner, bg=CARD)
        links.pack(fill="x", pady=(S(10), 0))
        ttk.Button(links, text="Select All", style="Small.TButton",
                   command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(links, text="Clear All", style="Small.TButton",
                   command=lambda: self._set_all(False)).pack(side="left", padx=(S(8), 0))
        self.count_var = tk.StringVar()
        tk.Label(links, textvariable=self.count_var, bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(side="right")

        tk.Frame(inner, bg=BORDER, height=1).pack(fill="x", pady=(S(12), S(14)))
        btns = tk.Frame(inner, bg=CARD)
        btns.pack(fill="x")
        ttk.Button(btns, text="Cancel", style="Ghost.TButton",
                   command=self.destroy).pack(side="right")
        self.export_btn = ttk.Button(btns, text="Export", style="Accent.TButton",
                                     command=self._export)
        self.export_btn.pack(side="right", padx=(0, S(8)))

        self.bind("<Escape>", lambda _e: self.destroy())
        self._sync()
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_reqwidth()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_reqheight()) // 3
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        self.focus_set()

    def chosen(self):
        return [c for c, var in self.vars.items() if var.get()]

    def _set_all(self, value):
        for var in self.vars.values():
            var.set(value)
        self._sync()

    def _sync(self):
        n = len(self.chosen())
        self.count_var.set(f"{n} of {len(self.vars)} fields selected")
        self.export_btn.state(["!disabled"] if n else ["disabled"])

    def _export(self):
        columns = self.chosen()
        self.destroy()
        self.on_export(columns)


class Step(tk.Frame):
    """Numbered step in the entry form: badge + title, with a body frame below."""

    def __init__(self, parent, number, title, hint, S):
        super().__init__(parent, bg=CARD)
        self.number, self.S = number, S
        head = tk.Frame(self, bg=CARD)
        head.pack(fill="x")
        self.badge = tk.Canvas(head, width=S(26), height=S(26), bg=CARD, highlightthickness=0)
        self.badge.pack(side="left")
        self.title_lbl = tk.Label(head, text=title, bg=CARD, fg=TEXT, font=(FONT, 11, "bold"))
        self.title_lbl.pack(side="left", padx=(S(10), 0))
        self.hint_lbl = tk.Label(head, text=hint, bg=CARD, fg=MUTED, font=(FONT, 9))
        self.hint_lbl.pack(side="left", padx=(S(8), 0))
        self.body = tk.Frame(self, bg=CARD)
        self.body.pack(fill="x", padx=(S(36), 0), pady=(S(4), 0))
        self._done = None
        self.set_done(False)

    def set_title(self, title, hint):
        self.title_lbl.configure(text=title)
        self.hint_lbl.configure(text=hint)

    def set_done(self, done):
        if done == self._done:
            return
        self._done = done
        n = self.S(26)
        self.badge.delete("all")
        self.badge.create_oval(1, 1, n - 1, n - 1, fill=SUCCESS if done else ACCENT, outline="")
        self.badge.create_text(n / 2, n / 2, text="✓" if done else str(self.number),
                               fill="white", font=(FONT, 10, "bold"))


# --------------------------------------------------------------------------- #
# App
# --------------------------------------------------------------------------- #
TABLE_COLUMNS = [  # (column, heading, width, anchor)
    ("Sr No", "#", 46, "center"),
    ("Client Short Name", "Client", 80, "w"),
    ("Client Group", "Client Group", 150, "w"),
    ("Assignment", "Assignment", 130, "w"),
    ("Work Type / Area", "Work Type / Area", 220, "w"),
    ("Work Status", "Status", 85, "center"),
    ("Frequency", "Frequency", 110, "w"),
    ("Financial Year", "FY", 80, "center"),
    ("Period From", "Period From", 95, "center"),
    ("Period To", "Period To", 95, "center"),
    ("Due Date", "Due Date", 95, "center"),
    ("Target Date", "Target Date", 95, "center"),
    ("Manager", "Manager", 140, "w"),
    ("Manager Short Name", "Mgr Short Name", 110, "center"),
    ("Client Full Name", "Client Full Name", 320, "w"),
    ("Date Added", "Date Added", 130, "w"),
]


class SchedulerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Work Scheduler")
        self.configure(bg=BG)
        self._scale = self.winfo_fpixels("1i") / 96
        self.geometry(f"{self.S(1360)}x{self.S(820)}")
        self.minsize(self.S(1100), self.S(720))
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

        self.master_data = None
        self.entries = []   # submitted (saved in Work Schedule.xlsx)
        self.pending = []   # added but not yet submitted
        self.export_fields = list(SCHEDULE_COLUMNS)

        self.client_var = tk.StringVar()
        self.assignment_var = tk.StringVar()
        self.work_type_var = tk.StringVar()
        self.frequency_var = tk.StringVar()
        self.short_name_var = tk.StringVar()
        self.client_group_var = tk.StringVar()
        self.duration_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.search_var = tk.StringVar()
        self.master_info_var = tk.StringVar()
        self.fy_var = tk.StringVar()
        self.fy_preview_var = tk.StringVar()
        self._master_mtime = None

        self._init_style()
        self._build_header()
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=self.S(24), pady=self.S(20))
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)
        self._build_form(main)
        self._build_table(main)

        self._load_master(initial=True)
        self._load_entries()
        self._refresh_pending()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def S(self, px):
        """Scale a pixel size for the screen's DPI."""
        return int(round(px * self._scale))

    # ---- styling -------------------------------------------------------- #
    def _init_style(self):
        S = self.S
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure(".", font=(FONT, 10), background=CARD, foreground=TEXT)

        st.configure("TCombobox", padding=S(5), fieldbackground="white", background="white",
                     foreground=TEXT, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                     arrowcolor=NAVY, arrowsize=S(14))
        st.map("TCombobox",
               fieldbackground=[("readonly", "white")],
               selectbackground=[("readonly", "white")],
               selectforeground=[("readonly", TEXT)],
               bordercolor=[("focus", ACCENT), ("hover", ACCENT)],
               lightcolor=[("focus", ACCENT)], darkcolor=[("focus", ACCENT)])
        self.option_add("*TCombobox*Listbox.font", (FONT, 10))
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*TCombobox*Listbox.selectForeground", "white")

        st.configure("Accent.TButton", background=ACCENT, foreground="white",
                     font=(FONT, 10, "bold"), padding=(S(18), S(10)), borderwidth=0,
                     focuscolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT, bordercolor=ACCENT)
        dis, act = ("disabled", "#A9BEF0"), ("active", ACCENT_DARK)
        st.map("Accent.TButton", background=[dis, act], lightcolor=[dis, act],
               darkcolor=[dis, act], bordercolor=[dis, act], foreground=[("disabled", "white")])

        st.configure("Success.TButton", background=SUCCESS, foreground="white",
                     font=(FONT, 10, "bold"), padding=(S(18), S(9)), borderwidth=0,
                     focuscolor=SUCCESS, lightcolor=SUCCESS, darkcolor=SUCCESS,
                     bordercolor=SUCCESS)
        dis, act = ("disabled", "#A7D7B5"), ("active", SUCCESS_DARK)
        st.map("Success.TButton", background=[dis, act], lightcolor=[dis, act],
               darkcolor=[dis, act], bordercolor=[dis, act], foreground=[("disabled", "white")])

        for name, fg, hover_bg, hover_border in (("Ghost", NAVY, HOVER_SOFT, ACCENT),
                                                 ("Danger", DANGER, "#FEF2F2", DANGER)):
            st.configure(f"{name}.TButton", background="white", foreground=fg,
                         font=(FONT, 10), padding=(S(14), S(9)), bordercolor=BORDER,
                         lightcolor="white", darkcolor="white", focuscolor="white")
            st.map(f"{name}.TButton", background=[("active", hover_bg)],
                   lightcolor=[("active", hover_bg)], darkcolor=[("active", hover_bg)],
                   bordercolor=[("active", hover_border)])
        st.configure("Small.TButton", background="white", foreground=ACCENT,
                     font=(FONT, 9), padding=(S(8), S(3)), bordercolor=BORDER,
                     lightcolor="white", darkcolor="white", focuscolor="white")
        st.map("Small.TButton", background=[("active", HOVER_SOFT)],
               lightcolor=[("active", HOVER_SOFT)], darkcolor=[("active", HOVER_SOFT)],
               bordercolor=[("active", ACCENT)])

        st.configure("Sched.Treeview", background="white", fieldbackground="white",
                     foreground=TEXT, rowheight=S(34), borderwidth=0, font=(FONT, 10))
        st.map("Sched.Treeview", background=[("selected", ACCENT_SOFT)],
               foreground=[("selected", TEXT)])
        st.configure("Sched.Treeview.Heading", background=NAVY, foreground="white",
                     font=(FONT, 10, "bold"), relief="flat", padding=(S(8), S(9)),
                     borderwidth=0, lightcolor=NAVY, darkcolor=NAVY, bordercolor=NAVY)
        st.map("Sched.Treeview.Heading", background=[("active", NAVY_2)])
        st.layout("Sched.Treeview", [("Sched.Treeview.treearea", {"sticky": "nswe"})])

        for orient in ("Vertical", "Horizontal"):
            st.configure(f"{orient}.TScrollbar", background=FAINT, troughcolor=ZEBRA,
                         bordercolor=ZEBRA, lightcolor=FAINT, darkcolor=FAINT,
                         arrowcolor=MUTED, gripcount=0)
            st.map(f"{orient}.TScrollbar", background=[("active", "#94A3B8")])

    # ---- layout --------------------------------------------------------- #
    def _card(self, parent):
        return tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)

    def _build_header(self):
        S = self.S
        bar = tk.Frame(self, bg=NAVY)
        bar.pack(fill="x")
        tk.Frame(self, bg=ACCENT, height=S(4)).pack(fill="x")

        left = tk.Frame(bar, bg=NAVY)
        left.pack(side="left", padx=S(28), pady=S(18))
        tk.Label(left, text="Work Scheduler", bg=NAVY, fg="white",
                 font=(FONT, 22, "bold")).pack(anchor="w")
        tk.Label(left, text="Plan client assignments, work areas, frequencies and coverage periods",
                 bg=NAVY, fg=ON_NAVY_MUTED, font=(FONT, 10)).pack(anchor="w", pady=(S(2), 0))
        tk.Label(left, textvariable=self.master_info_var, bg=NAVY, fg="#7DD3A8",
                 font=(FONT, 9)).pack(anchor="w", pady=(S(4), 0))

        stats = tk.Frame(bar, bg=NAVY)
        stats.pack(side="right", padx=S(28))
        self.stat_vars = {}
        for key, label in (("entries", "Scheduled Entries"), ("clients", "Clients Covered"),
                           ("wip", "Work in Progress"), ("complete", "Completed")):
            chip = tk.Frame(stats, bg=NAVY_2, padx=S(20), pady=S(8))
            chip.pack(side="left", padx=(S(10), 0))
            var = tk.StringVar(value="0")
            self.stat_vars[key] = var
            tk.Label(chip, textvariable=var, bg=NAVY_2, fg="white",
                     font=(FONT, 18, "bold")).pack(anchor="w")
            tk.Label(chip, text=label.upper(), bg=NAVY_2, fg=ON_NAVY_MUTED,
                     font=(FONT, 8, "bold")).pack(anchor="w")

    def _build_form(self, parent):
        S = self.S
        card = self._card(parent)
        card.grid(row=0, column=0, sticky="ns")
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="both", expand=True, padx=S(24), pady=S(20))

        # Footer is packed first (at the bottom) so the buttons are always visible.
        footer = tk.Frame(inner, bg=CARD)
        footer.pack(side="bottom", fill="x")

        title_row = tk.Frame(inner, bg=CARD)
        title_row.pack(fill="x")
        tk.Label(title_row, text="New Schedule Entry", bg=CARD, fg=TEXT,
                 font=(FONT, 14, "bold")).pack(side="left")
        ttk.Button(title_row, text="⟳ Reload", style="Small.TButton",
                   command=self._load_master).pack(side="right")
        tk.Label(inner, text="Complete each step to unlock the next.",
                 bg=CARD, fg=MUTED, font=(FONT, 9)).pack(anchor="w", pady=(S(2), S(12)))

        steps = tk.Frame(inner, bg=CARD)
        steps.pack(fill="x")
        steps.columnconfigure(0, weight=1)
        self.steps = []

        def step(title, hint):
            s = Step(steps, len(self.steps) + 1, title, hint, S)
            s.grid(row=len(self.steps), column=0, sticky="we", pady=(0, S(12)))
            self.steps.append(s)
            return s

        list_font = tkfont.Font(family=FONT, size=10)

        def combo(s, var, handler):
            style = f"Step{len(self.steps)}.TCombobox"
            cb = ttk.Combobox(s.body, textvariable=var, state="readonly", width=32, style=style)

            def widen_popdown():
                # The field is narrow; let the drop-down list grow to fit long names.
                longest = max((list_font.measure(v) for v in cb["values"]), default=0)
                extra = max(0, longest + S(40) - cb.winfo_width())
                ttk.Style(self).configure(style, postoffset=(0, 0, extra, 0))

            cb.configure(postcommand=widen_popdown)
            cb.pack(fill="x")
            cb.bind("<<ComboboxSelected>>", handler)
            return cb

        s = step("Client", "from Client List")
        self.client_cb = combo(s, self.client_var, self._on_client)
        self.short_row = tk.Frame(s.body, bg=CARD)
        tk.Label(self.short_row, text="Short Name", bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(side="left")
        tk.Label(self.short_row, textvariable=self.short_name_var, bg=ACCENT_SOFT,
                 fg=ACCENT_DARK, font=(FONT, 10, "bold"), padx=S(10), pady=S(2)).pack(
            side="left", padx=(S(8), 0))
        self.group_row = tk.Frame(s.body, bg=CARD)
        tk.Label(self.group_row, text="Client Group", bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(side="left")
        tk.Label(self.group_row, textvariable=self.client_group_var, bg=CARD, fg=ACCENT_DARK,
                 font=(FONT, 9, "bold")).pack(side="left", padx=(S(8), 0))

        s = step("Assignment", "from Assignment Master")
        self.assignment_cb = combo(s, self.assignment_var, self._on_assignment)

        s = step("Work Type / Area", "for Assignment")
        self.work_type_cb = combo(s, self.work_type_var, self._on_work_type)

        s = step("Frequency", "for Work Type")
        self.frequency_cb = combo(s, self.frequency_var, self._on_frequency)

        s = step("Period Covered", "required")
        self.period_step = s

        # (a) Financial Year – Monthly / Quarterly / Half Yearly / Annual
        self.fy_frame = tk.Frame(s.body, bg=CARD)
        self.fy_cb = ttk.Combobox(self.fy_frame, textvariable=self.fy_var, state="readonly",
                                  width=32)
        start = current_fy_start()
        self.fy_cb["values"] = [fy_label(y) for y in range(start - 2, start + 5)]
        self.fy_cb.pack(fill="x")
        self.fy_cb.bind("<<ComboboxSelected>>", lambda _e: self._sync())
        tk.Label(self.fy_frame, textvariable=self.fy_preview_var, bg=CARD, fg=ACCENT_DARK,
                 font=(FONT, 9, "bold"), justify="left").pack(anchor="w", pady=(S(6), 0))

        # (c) NA – period not applicable
        self.na_frame = tk.Frame(s.body, bg=CARD)
        tk.Label(self.na_frame, text=NA, bg=AMBER_SOFT, fg=AMBER_DARK, font=(FONT, 10, "bold"),
                 padx=S(10), pady=S(2)).pack(side="left")
        tk.Label(self.na_frame, text="Period From / To will be recorded as NA", bg=CARD,
                 fg=MUTED, font=(FONT, 9)).pack(side="left", padx=(S(8), 0))

        # (b) From / To dates – 'Annual Dec End', 'As per Engagement' and other frequencies
        self.date_frame = tk.Frame(s.body, bg=CARD)
        period = tk.Frame(self.date_frame, bg=CARD)
        period.pack(fill="x")
        period.columnconfigure((0, 1), weight=1, uniform="p")
        for col, text in enumerate(("From", "To")):
            tk.Label(period, text=text.upper(), bg=CARD, fg=MUTED,
                     font=(FONT, 8, "bold")).grid(row=0, column=col, sticky="w",
                                                  padx=(0 if col == 0 else S(6), 0))
        self.from_field = DateField(period, self._on_period, S)
        self.from_field.grid(row=1, column=0, sticky="we", padx=(0, S(6)), pady=(S(3), 0))
        self.to_field = DateField(period, self._on_period, S,
                                  min_date=lambda: self.from_field.value)
        self.to_field.grid(row=1, column=1, sticky="we", padx=(S(6), 0), pady=(S(3), 0))
        tk.Label(self.date_frame, textvariable=self.duration_var, bg=CARD, fg=ACCENT_DARK,
                 font=(FONT, 9, "bold")).pack(anchor="w", pady=(S(6), 0))

        self.status_lbl = tk.Label(footer, textvariable=self.status_var, bg=CARD, fg=WARN,
                                   font=(FONT, 9), wraplength=S(290), justify="left")
        self.status_lbl.pack(anchor="w")

        tk.Frame(footer, bg=BORDER, height=1).pack(fill="x", pady=(S(8), S(14)))
        btns = tk.Frame(footer, bg=CARD)
        btns.pack(fill="x")
        self.add_btn = ttk.Button(btns, text="＋  Add to Pending", style="Accent.TButton",
                                  command=self._add_entry)
        self.add_btn.pack(side="left", fill="x", expand=True)
        ttk.Button(btns, text="Clear", style="Ghost.TButton",
                   command=self._reset_form).pack(side="left", padx=(S(8), 0))

    def _make_tree(self, parent, height=None, style="Sched.Treeview"):
        """Bordered Treeview with scrollbars and an 'empty' overlay label."""
        S = self.S
        table = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        cols = [CHECK_COL] + [c for c, *_ in TABLE_COLUMNS]
        tree = ttk.Treeview(table, columns=cols, show="headings", selectmode="extended",
                            style=style, **({"height": height} if height else {}))
        # Tick-box column: click a box to tick/untick that row; click the header box
        # to tick/untick every row currently shown.
        tree.heading(CHECK_COL, text=UNCHECKED, anchor="center",
                     command=lambda: self._toggle_all(tree))
        tree.column(CHECK_COL, width=S(40), minwidth=S(40), anchor="center", stretch=False)
        for col, heading, width, anchor in TABLE_COLUMNS:
            tree.heading(col, text=heading, anchor=anchor)
            tree.column(col, width=S(width), minwidth=S(40), anchor=anchor, stretch=False)
        tree.tag_configure("odd", background=ZEBRA)
        # Created after 'odd' so they take priority over the zebra stripe.
        tree.tag_configure("wip", background="#FFF6E0")
        tree.tag_configure("complete", background="#E8F8EE")
        tree.tag_configure("overdue", foreground="#B91C1C")
        tree.bind("<Button-1>", lambda e: self._on_tree_click(tree, e))
        tree.bind("<<TreeviewSelect>>", lambda _e: self._update_checks(tree))
        tree.bind("<Control-a>", lambda _e: (tree.selection_set(tree.get_children()), "break")[1])
        ysb = ttk.Scrollbar(table, orient="vertical", command=tree.yview)
        xsb = ttk.Scrollbar(table, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="we")
        tree.empty_lbl = tk.Label(table, bg="white", fg=MUTED, font=(FONT, 10),
                                  justify="center")
        return table, tree

    def _fill_tree(self, tree, rows, empty_text, query=""):
        """Load rows into a tree (filtered by query); returns how many are shown.

        Every word in the query must appear somewhere in the row, so e.g.
        "BRMCPL FY 2026-27 Monthly" narrows to exactly those rows.
        """
        tree.delete(*tree.get_children())
        words = query.split()
        shown = 0
        for i, e in enumerate(rows):
            e["Sr No"] = i + 1
            values = [fmt_cell(e.get(c)) for c, *_ in TABLE_COLUMNS]
            text = " ".join(str(v) for v in values).lower()
            if words and not all(w in text for w in words):
                continue
            tags = ["odd"] if shown % 2 else []
            status = e.get("Work Status")
            if status == "WIP":
                tags.append("wip")
            elif status == "Complete":
                tags.append("complete")
            due = e.get("Due Date")
            if isinstance(due, date) and due < date.today() and status != "Complete":
                tags.append("overdue")
            tree.insert("", "end", iid=str(i), values=[UNCHECKED] + values, tags=tags)
            shown += 1
        if shown:
            tree.empty_lbl.place_forget()
        else:
            tree.empty_lbl.configure(text=empty_text)
            tree.empty_lbl.place(in_=tree, relx=0.5, rely=0.5, anchor="center")
        self._update_checks(tree)
        return shown

    def _build_update_bar(self, parent):
        """Manager / Due Date / Target Date / Work Status editor for the selected rows."""
        S = self.S
        bar = tk.Frame(parent, bg=HOVER_SOFT, highlightthickness=1, highlightbackground=ACCENT_SOFT)
        bar.pack(fill="x", pady=(0, S(10)))
        row = tk.Frame(bar, bg=HOVER_SOFT)
        row.pack(fill="x", padx=S(14), pady=(S(10), S(4)))

        def field(label, widget_factory, col):
            tk.Label(row, text=label.upper(), bg=HOVER_SOFT, fg=MUTED,
                     font=(FONT, 8, "bold")).grid(row=0, column=col, sticky="w", padx=(0, S(12)))
            w = widget_factory(row)
            w.grid(row=1, column=col, sticky="we", padx=(0, S(12)), pady=(S(2), 0))
            return w

        self.mgr_var = tk.StringVar()
        self.status_edit_var = tk.StringVar()
        self.mgr_cb = field("Manager", lambda p: ttk.Combobox(p, textvariable=self.mgr_var,
                                                              state="readonly", width=22), 0)
        # Short name follows the Manager picked, from the 'Manager Master' sheet
        self.mgr_short_var = tk.StringVar()
        self.mgr_var.trace_add("write", lambda *_: self.mgr_short_var.set(
            self._manager_short(self.mgr_var.get()) or ""))
        field("Mgr Short Name", lambda p: tk.Label(
            p, textvariable=self.mgr_short_var, bg=ACCENT_SOFT, fg=ACCENT_DARK,
            font=(FONT, 10, "bold"), width=8, pady=S(5)), 1)
        self.due_field = field("Due Date", lambda p: DateField(p, lambda: None, S), 2)
        self.target_field = field("Target Date", lambda p: DateField(p, lambda: None, S), 3)
        self.status_cb = field("Work Status", lambda p: ttk.Combobox(
            p, textvariable=self.status_edit_var, state="readonly", width=18,
            values=[NOT_STARTED, *WORK_STATUSES]), 4)
        for w in (self.due_field, self.target_field):
            w.text.configure(width=12)
        self.apply_btn = ttk.Button(row, text="Apply to Selected", style="Accent.TButton",
                                    command=self._apply_updates, width=22)
        self.apply_btn.grid(row=1, column=5, sticky="w", pady=(S(2), 0))

        self.update_hint_var = tk.StringVar()
        tk.Label(bar, textvariable=self.update_hint_var, bg=HOVER_SOFT, fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(fill="x", padx=S(14), pady=(0, S(8)))
        self._prefilled_iid = None

    def _set_update_fields(self, e=None):
        e = e or {}
        self.mgr_var.set((e.get("Manager") or NO_MANAGER) if e else "")
        self.due_field.set(to_date(e.get("Due Date")), notify=False)
        self.target_field.set(to_date(e.get("Target Date")), notify=False)
        self.status_edit_var.set((e.get("Work Status") or NOT_STARTED) if e else "")

    def _sync_update_bar(self):
        sel = self.tree.selection()
        n = len(sel)
        if n == 1 and sel[0] != self._prefilled_iid:
            # One row: show its current values so they can be edited.
            self._set_update_fields(self.entries[int(sel[0])])
            self._prefilled_iid = sel[0]
        elif n != 1 and self._prefilled_iid is not None:
            self._set_update_fields()
            self._prefilled_iid = None
        self.apply_btn.configure(text=f"Apply to Selected ({n})" if n else "Apply to Selected")
        self.apply_btn.state(["!disabled"] if n else ["disabled"])
        self.update_hint_var.set(
            "Tick rows (or Ctrl / Shift-click) to set Manager, dates and Work Status, or to "
            "delete them.  Tip: search e.g. “BRMCPL FY 2026-27 Monthly”, then Select All."
            if not n else
            "Editing 1 row – change any field and click Apply (an empty field clears it)."
            if n == 1 else
            f"Updating {n} rows – only the fields you fill in are changed; the rest are left "
            "as they are.  Choose “Not started (blank)” to clear Work Status.")

    def _manager_short(self, name):
        return self.master_data.manager_short.get(name) if self.master_data else None

    def _sync_from_masters(self):
        """Bring rows' Client Group and manager short name in line with the masters.

        Returns True if any row changed.
        """
        changed = False
        for e in self.entries:
            client = e.get("Client Full Name")
            if client in self.master_data.client_groups:
                group = self.master_data.client_groups[client]
                if e.get("Client Group") != group:
                    e["Client Group"] = group
                    changed = True
            name = e.get("Manager")
            if name in self.master_data.manager_short:
                short = self._manager_short(name)
                if e.get("Manager Short Name") != short:
                    e["Manager Short Name"] = short
                    changed = True
        return changed

    def _refresh_manager_choices(self):
        """Manager drop-down = names in the 'Manager Master' sheet (plus a 'none' option)."""
        if not hasattr(self, "mgr_cb"):
            return
        master = self.master_data.managers if self.master_data else []
        self.mgr_cb["values"] = [NO_MANAGER, *master]

    def _apply_updates(self):
        sel = list(self.tree.selection())
        if not sel:
            return
        rows = [self.entries[int(i)] for i in sel]
        mgr = self.mgr_var.get()
        mgr_value = None if mgr in ("", NO_MANAGER) else mgr
        due, target = self.due_field.value, self.target_field.value
        status = self.status_edit_var.get()
        status_value = None if status in ("", NOT_STARTED) else status

        if len(rows) == 1:
            changes = {"Manager": mgr_value, "Manager Short Name": self._manager_short(mgr_value),
                       "Due Date": due, "Target Date": target,
                       "Work Status": status_value}
        else:
            changes = {}
            if mgr:
                changes["Manager"] = mgr_value
                changes["Manager Short Name"] = self._manager_short(mgr_value)
            if due:
                changes["Due Date"] = due
            if target:
                changes["Target Date"] = target
            if status:
                changes["Work Status"] = status_value
            if not changes:
                self._status("Fill in at least one field to apply to the selected rows.")
                return

        if due and target and target > due and not messagebox.askyesno(
                "Target after Due Date",
                f"Target Date {target:%d-%b-%Y} is after Due Date {due:%d-%b-%Y}.\n\n"
                "Apply anyway?", icon="warning"):
            return

        backup = [dict(r) for r in rows]
        for r in rows:
            r.update(changes)
        if not self._persist():
            for r, b in zip(rows, backup):
                r.clear()
                r.update(b)
            return
        self._refresh_table()
        visible = [i for i in sel if self.tree.exists(i)]
        self.tree.selection_set(visible)
        if visible:
            self.tree.see(visible[0])
        self._status(f"✓ Updated {self._plural(len(rows))}: "
                     + ", ".join(changes) + ".", SUCCESS)

    # ---- multi-select ---------------------------------------------------- #
    @staticmethod
    def _on_tree_click(tree, event):
        """A click in the tick-box column toggles just that row (no Ctrl needed)."""
        if tree.identify_region(event.x, event.y) != "cell" or tree.identify_column(event.x) != "#1":
            return None
        iid = tree.identify_row(event.y)
        if iid:
            tree.selection_toggle(iid)
            tree.focus(iid)
            tree.focus_set()
        return "break"

    @staticmethod
    def _toggle_all(tree):
        items = tree.get_children()
        if items and set(tree.selection()) >= set(items):
            tree.selection_remove(items)
        else:
            tree.selection_add(items)

    def _update_checks(self, tree):
        items = tree.get_children()
        sel = set(tree.selection())
        for iid in items:
            tree.set(iid, CHECK_COL, CHECKED if iid in sel else UNCHECKED)
        all_ticked = bool(items) and sel >= set(items)
        tree.heading(CHECK_COL, text=CHECKED if all_ticked else UNCHECKED)
        if tree is getattr(self, "tree", None):
            self._update_count()
            self.select_all_btn.configure(text="Unselect All" if all_ticked else "Select All")
            self.delete_btn.configure(text=f"Delete Selected ({len(sel)})" if sel
                                      else "Delete Selected")
            self.delete_btn.state(["!disabled"] if sel else ["disabled"])
            self._sync_update_bar()

    def _update_count(self):
        total = len(self.entries)
        shown = len(self.tree.get_children())
        n_sel = len(self.tree.selection())
        self.count_var.set(f"Showing {shown} of {total} entr{'y' if total == 1 else 'ies'}"
                           f"  ·  {n_sel} selected")

    def _build_table(self, parent):
        S = self.S
        right = tk.Frame(parent, bg=BG)
        right.grid(row=0, column=1, sticky="nsew", padx=(S(20), 0))

        # ---- Pending Submission ---- #
        card = tk.Frame(right, bg=CARD, highlightthickness=2, highlightbackground=AMBER)
        card.pack(fill="x", pady=(0, S(16)))
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="both", expand=True, padx=S(24), pady=S(16))

        top = tk.Frame(inner, bg=CARD)
        top.pack(fill="x", pady=(0, S(10)))
        tk.Label(top, text="Pending Submission", bg=CARD, fg=TEXT,
                 font=(FONT, 14, "bold")).pack(side="left")
        self.pending_badge = tk.Label(top, bg=AMBER_SOFT, fg=AMBER_DARK,
                                      font=(FONT, 10, "bold"), padx=S(10), pady=S(1))
        self.pending_badge.pack(side="left", padx=(S(10), 0))
        tk.Label(top, text="Not saved until you click Submit", bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(side="left", padx=(S(10), 0))

        self.submit_btn = ttk.Button(top, text="✓  Submit to Scheduler", style="Success.TButton",
                                     command=self._submit_pending)
        self.submit_btn.pack(side="right")
        ttk.Button(top, text="Remove Selected", style="Danger.TButton",
                   command=self._remove_pending).pack(side="right", padx=(0, S(8)))

        table, self.pending_tree = self._make_tree(inner, height=4)
        table.pack(fill="x")
        self.pending_table = table   # hidden while the Pending list is empty

        # ---- Scheduler ---- #
        card = self._card(right)
        card.pack(fill="both", expand=True)
        inner = tk.Frame(card, bg=CARD)
        inner.pack(fill="both", expand=True, padx=S(24), pady=S(20))

        top = tk.Frame(inner, bg=CARD)
        top.pack(fill="x", pady=(0, S(14)))
        titles = tk.Frame(top, bg=CARD)
        titles.pack(side="left")
        tk.Label(titles, text="Scheduler", bg=CARD, fg=TEXT,
                 font=(FONT, 14, "bold")).pack(anchor="w")
        self.count_var = tk.StringVar()
        tk.Label(titles, textvariable=self.count_var, bg=CARD, fg=MUTED,
                 font=(FONT, 9)).pack(anchor="w")

        ttk.Button(top, text="Export to Excel", style="Ghost.TButton",
                   command=self._open_export_dialog).pack(side="right")
        self.delete_btn = ttk.Button(top, text="Delete Selected", style="Danger.TButton",
                                     command=self._delete_selected, width=20)
        self.delete_btn.pack(side="right", padx=(0, S(8)))
        self.select_all_btn = ttk.Button(top, text="Select All", style="Ghost.TButton",
                                         command=lambda: self._toggle_all(self.tree), width=12)
        self.select_all_btn.pack(side="right", padx=(0, S(8)))
        search = tk.Frame(top, bg="white", highlightthickness=1, highlightbackground=BORDER,
                          highlightcolor=ACCENT)
        search.pack(side="right", padx=(0, S(12)))
        tk.Label(search, text="🔍", bg="white", fg=MUTED,
                 font=("Segoe UI Emoji", 10)).pack(side="left", padx=(S(8), 0))
        entry = tk.Entry(search, textvariable=self.search_var, relief="flat", bg="white",
                         fg=TEXT, font=(FONT, 10), width=24, insertbackground=TEXT)
        entry.pack(side="left", padx=S(6), pady=S(8))
        self.search_var.trace_add("write", lambda *_: self._refresh_table())

        self._build_update_bar(inner)
        table, self.tree = self._make_tree(inner)
        table.pack(fill="both", expand=True)
        self.tree.bind("<Delete>", lambda _e: self._delete_selected())
        self.pending_tree.bind("<Delete>", lambda _e: self._remove_pending())

    # ---- data loading -------------------------------------------------- #
    def _load_master(self, initial=False, auto=False):
        """(Re)read the master workbook. Runs at every start-up, on Reload, and
        automatically whenever the workbook is saved while the app is open."""
        try:
            mtime = MASTER_FILE.stat().st_mtime
            self.master_data = MasterData(MASTER_FILE)
        except PermissionError:
            if auto:  # Excel/OneDrive may still be writing it – try again on the next check
                return
            messagebox.showerror("Work Scheduler",
                                 f"Can't read {MASTER_FILE.name} – close it in Excel and retry.")
            if initial:
                self.destroy()
                sys.exit(1)
            return
        except Exception as exc:  # noqa: BLE001 - show any master-data problem to the user
            if auto:
                self._master_mtime = MASTER_FILE.stat().st_mtime  # don't re-alert every check
                self._status(f"Master workbook changed but could not be read: {exc}")
                return
            messagebox.showerror("Work Scheduler", f"Could not load master workbook:\n{exc}")
            if initial:
                self.destroy()
                sys.exit(1)
            return

        self._master_mtime = mtime
        self.master_info_var.set(
            f"Master lists loaded {datetime.now():%d-%m-%Y %H:%M}  ·  "
            f"{len(self.master_data.clients)} clients, "
            f"{len(self.master_data.assignments)} assignments, "
            f"{len(self.master_data.work_type_rows)} work types, "
            f"{len(self.master_data.managers)} managers  ·  auto-updates on change")
        self.client_cb["values"] = list(self.master_data.clients)
        self._refresh_manager_choices()
        self.mgr_var.set(self.mgr_var.get())   # re-show the short name from the new master
        if self._sync_from_masters():
            self._refresh_table()
        if initial:
            self._reset_form()
            self._watch_master()
            return
        kept = self._reapply_selection()
        self._status(("Master workbook changed – lists updated automatically."
                      if auto else "Master lists reloaded.")
                     + ("" if kept else " Your selection is no longer in the masters and "
                                        "was cleared."), SUCCESS if kept else WARN)

    def _reapply_selection(self):
        """After a reload, keep the current form choices if they are still valid."""
        c, a = self.client_var.get(), self.assignment_var.get()
        w, f = self.work_type_var.get(), self.frequency_var.get()
        md = self.master_data
        self.client_var.set("")
        self._clear_from(self.assignment_cb, self.work_type_cb, self.frequency_cb)
        if not c:
            self._sync()
            return True
        if c not in md.clients:
            self.short_name_var.set("")
            self._sync()
            return False
        self.client_var.set(c)
        self._on_client()
        for value, valid, cb, handler in (
                (a, lambda: md.assignments, self.assignment_cb, self._on_assignment),
                (w, lambda: md.work_types_for(a), self.work_type_cb, self._on_work_type),
                (f, lambda: md.frequencies_for(a, w), self.frequency_cb, self._on_frequency)):
            if not value:
                return True
            if value not in valid():
                return False
            cb.set(value)
            handler()
        return True

    def _watch_master(self):
        """Check every few seconds whether the master workbook was saved."""
        try:
            if MASTER_FILE.stat().st_mtime != self._master_mtime:
                self._load_master(auto=True)
        except OSError:
            pass
        self.after(3000, self._watch_master)

    def _load_entries(self):
        try:
            self.entries = load_schedule()
        except PermissionError:
            messagebox.showwarning("Work Scheduler",
                                   f"{SCHEDULE_FILE.name} is open in Excel – showing an empty list.")
            self.entries = []
        self._sync_from_masters()
        self._refresh_table()

    def _refresh_pending(self):
        self._fill_tree(self.pending_tree, self.pending,
                        "No pending entries.  Fill the form and click  ＋ Add to Pending.")
        n = len(self.pending)
        if n:
            self.pending_table.pack(fill="x")
        else:
            self.pending_table.pack_forget()
        self.pending_badge.configure(text=f"{n} pending")
        self.submit_btn.state(["!disabled"] if n else ["disabled"])

    def _refresh_table(self):
        query = self.search_var.get().strip().lower()
        total = len(self.entries)
        shown = self._fill_tree(
            self.tree, self.entries,
            "No matching entries." if total else
            "Nothing submitted yet.\nAdd entries to Pending, then click  ✓ Submit to Scheduler.",
            query)

        self._update_count()
        self.stat_vars["entries"].set(str(total))
        self.stat_vars["clients"].set(str(len({e.get("Client Full Name") for e in self.entries})))
        for key, status in (("wip", "WIP"), ("complete", "Complete")):
            self.stat_vars[key].set(str(sum(e.get("Work Status") == status for e in self.entries)))
        self._refresh_manager_choices()

    # ---- form state ---------------------------------------------------- #
    def _status(self, text, color=WARN):
        self.status_var.set(text)
        self.status_lbl.configure(fg=color)

    def _sync(self):
        """Show/hide steps, update badges and enable the Add button from current values."""
        c, a = self.client_var.get(), self.assignment_var.get()
        w, f = self.work_type_var.get(), self.frequency_var.get()
        fr, to = self.from_field.value, self.to_field.value

        mode, months = frequency_kind(f) if f else (None, None)
        period_done = {"fy": bool(self.fy_var.get()), "manual": bool(fr and to),
                       "na": True}.get(mode, False)

        has_work_types = bool(a and self.work_type_cb["values"])
        visible = [True, bool(c), bool(c and has_work_types), bool(c and a and w),
                   bool(c and a and w and f)]
        done = [bool(c), bool(c and a), bool(c and a and w), bool(c and a and w and f),
                period_done]

        # Step 5 changes with the Frequency: Financial Year / From–To dates / NA
        for frame, m in ((self.fy_frame, "fy"), (self.date_frame, "manual"),
                         (self.na_frame, "na")):
            if mode == m:
                frame.pack(fill="x")
            else:
                frame.pack_forget()
        if mode == "fy":
            n = 12 // months
            self.period_step.set_title("Financial Year",
                                       f"creates {n} line item{'s' if n > 1 else ''}")
            fy = self.fy_var.get()
            if fy:
                periods = fy_periods(fy_start_year(fy), months)
                first, last = periods[0], periods[-1]
                rng = lambda p: f"{p[0]:%d-%b-%Y} → {p[1]:%d-%b-%Y}"
                self.fy_preview_var.set(rng(first) if n == 1 else
                                        f"{rng(first)}\n{'…  ' if n > 2 else ''}{rng(last)}")
            else:
                self.fy_preview_var.set("")
        elif mode == "na":
            self.period_step.set_title("Period Covered", "not applicable")
        else:
            self.period_step.set_title("Period Covered", "select From and To dates")
        for s, vis, ok in zip(self.steps, visible, done):
            if vis:
                s.grid()
            else:
                s.grid_remove()
            s.set_done(vis and ok)

        if c:
            self.short_row.pack(anchor="w", pady=(self.S(6), 0))
            self.group_row.pack(anchor="w", pady=(self.S(3), 0))
        else:
            self.short_row.pack_forget()
            self.group_row.pack_forget()

        if fr and to:
            days = (to - fr).days + 1
            self.duration_var.set(f"{fr:%d %b %Y}  →  {to:%d %b %Y}   ·   "
                                  f"{days} day{'s' if days != 1 else ''}")
        else:
            self.duration_var.set("")

        self.add_btn.state(["!disabled"] if all(v and d for v, d in zip(visible, done))
                           else ["disabled"])

    def _clear_from(self, *cbs):
        for cb in cbs:
            cb.set("")
            cb["values"] = []

    def _reset_form(self):
        self.client_var.set("")
        self._clear_from(self.assignment_cb, self.work_type_cb, self.frequency_cb)
        self.from_field.set(None, notify=False)
        self.to_field.set(None, notify=False)
        self.fy_var.set("")
        self._status("")
        self._sync()

    def _on_client(self, _event=None):
        self.short_name_var.set(self.master_data.clients.get(self.client_var.get(), ""))
        self.client_group_var.set(self.master_data.client_groups.get(self.client_var.get()) or "")
        self._clear_from(self.assignment_cb, self.work_type_cb, self.frequency_cb)
        self.assignment_cb["values"] = self.master_data.assignments
        self._status("")
        self._sync()

    def _on_assignment(self, _event=None):
        self._clear_from(self.work_type_cb, self.frequency_cb)
        work_types = self.master_data.work_types_for(self.assignment_var.get())
        self.work_type_cb["values"] = work_types
        self._status("" if work_types else
                     "No Work Type / Area is mapped to this Assignment in 'Work Type Master'.")
        self._sync()

    def _on_work_type(self, _event=None):
        self._clear_from(self.frequency_cb)
        # Always left blank – the user picks the Frequency themselves.
        self.frequency_cb["values"] = self.master_data.frequencies_for(
            self.assignment_var.get(), self.work_type_var.get())
        self._sync()

    def _on_frequency(self, _event=None):
        self._sync()

    def _on_period(self):
        fr, to = self.from_field.value, self.to_field.value
        if fr and to and to < fr:
            self.to_field.set(None, notify=False)
            self._status("'To' date was cleared – it must be on or after the 'From' date.")
        else:
            self._status("")
        self._sync()

    # ---- actions ------------------------------------------------------- #
    def _build_line_items(self):
        """Turn the form into one or more schedule rows; None if the period is incomplete."""
        base = {
            "Date Added": datetime.now().strftime("%d-%m-%Y %H:%M"),
            "Client Full Name": self.client_var.get(),
            "Client Short Name": self.short_name_var.get(),
            "Client Group": self.client_group_var.get() or None,
            "Assignment": self.assignment_var.get(),
            "Work Type / Area": self.work_type_var.get(),
            "Frequency": self.frequency_var.get(),
            "Financial Year": "",
        }
        mode, months = frequency_kind(base["Frequency"])
        if mode == "fy":
            fy = self.fy_var.get()
            if not fy:
                messagebox.showwarning("Financial Year required",
                                       "Select the Financial Year before adding.")
                return None
            return [dict(base, **{"Financial Year": fy, "Period From": a, "Period To": b})
                    for a, b in fy_periods(fy_start_year(fy), months)]
        if mode == "na":
            return [dict(base, **{"Period From": NA, "Period To": NA})]
        fr, to = self.from_field.value, self.to_field.value
        if not (fr and to):
            messagebox.showwarning("Period Covered required",
                                   "Select both the 'From' and 'To' dates before adding.")
            return None
        return [dict(base, **{"Period From": fr, "Period To": to})]

    def _add_entry(self):
        items = self._build_line_items()
        if not items:
            return
        new, dups = [], []
        for item in items:
            if self._is_duplicate(item, self.pending + new):
                dups.append((item, "Pending list"))
            elif self._is_duplicate(item, self.entries):
                dups.append((item, "Scheduler"))
            else:
                new.append(item)
        if dups:
            messagebox.showwarning(
                "Duplicate entry",
                f"{self._plural(len(dups))} {'is' if len(dups) == 1 else 'are'} already "
                "scheduled and " + ("was" if len(dups) == 1 else "were") + " not added:\n\n"
                + "\n".join(f"{self._describe(e)}   [in {where}]" for e, where in dups)
                + (f"\n\n{self._plural(len(new))} added to Pending." if new else ""))
        if not new:
            self._status("Nothing added – duplicate of an existing entry.")
            return

        self.pending.extend(new)
        self._refresh_pending()
        last = self.pending_tree.get_children()
        if last:
            self.pending_tree.see(last[-1])
        # Keep the client, FY and dates so several work types can be added quickly.
        self._clear_from(self.work_type_cb, self.frequency_cb)
        self.assignment_cb.set("")
        self._status(f"Added {self._plural(len(new))} to Pending: {new[0]['Client Short Name']}"
                     f" – {new[0]['Work Type / Area']}.  Click Submit to save.", SUCCESS)
        self._sync()

    @staticmethod
    def _is_duplicate(entry, rows):
        keys = ("Client Full Name", "Assignment", "Work Type / Area", "Frequency",
                "Period From", "Period To")
        return any(all(e.get(k) == entry[k] for k in keys) for e in rows)

    def _remove_pending(self):
        sel = self.pending_tree.selection()
        if not sel:
            self._status("Select one or more Pending rows to remove.")
            return
        drop = {int(i) for i in sel}
        self.pending = [e for i, e in enumerate(self.pending) if i not in drop]
        self._refresh_pending()

    @staticmethod
    def _describe(e):
        return (f"• {e.get('Client Short Name')} | {e.get('Assignment')} | "
                f"{e.get('Work Type / Area')} | {e.get('Frequency')} | "
                f"{fmt_cell(e.get('Period From'))} to {fmt_cell(e.get('Period To'))}")

    @staticmethod
    def _plural(n):
        return f"{n} entr{'y' if n == 1 else 'ies'}"

    def _submit_pending(self):
        if not self.pending:
            return
        no_period = [e for e in self.pending if not (e.get("Period From") and e.get("Period To"))]
        if no_period:
            messagebox.showwarning(
                "Period Covered required",
                "These entries have no Period Covered. Remove them and add them again "
                "with 'From' and 'To' dates before submitting:\n\n"
                + "\n".join(self._describe(e) for e in no_period))
            return
        if not messagebox.askyesno("Submit to Scheduler",
                                   f"Submit {self._plural(len(self.pending))} to the Scheduler?"
                                   f"\n\nThey will be saved to {SCHEDULE_FILE.name}."):
            return

        to_add, skipped = [], []
        for e in self.pending:
            (skipped if self._is_duplicate(e, self.entries + to_add) else to_add).append(e)

        if to_add:
            before = len(self.entries)
            self.entries.extend(to_add)
            if not self._persist():
                del self.entries[before:]
                return
        self.pending = []
        self._refresh_pending()
        self._refresh_table()
        last = self.tree.get_children()
        if last:
            self.tree.see(last[-1])

        if skipped:
            messagebox.showwarning(
                "Duplicate entries skipped",
                f"{self._plural(len(skipped))} {'was' if len(skipped) == 1 else 'were'} "
                "already in the Scheduler and skipped:\n\n"
                + "\n".join(self._describe(e) for e in skipped)
                + f"\n\n{self._plural(len(to_add))} submitted.")
        self._status(f"✓ {self._plural(len(to_add))} submitted to the Scheduler"
                     + (f", {len(skipped)} duplicate{'s' if len(skipped) > 1 else ''} skipped."
                        if skipped else "."),
                     SUCCESS if to_add else WARN)

    def _on_close(self):
        n = len(self.pending)
        if n and not messagebox.askyesno(
                "Work Scheduler",
                f"You have {n} pending entr{'y' if n == 1 else 'ies'} that "
                f"{'has' if n == 1 else 'have'} not been submitted.\n\n"
                "Close anyway and discard them?", icon="warning"):
            return
        self.destroy()

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            self._status("Select one or more rows in the Scheduler to delete.")
            return
        rows = [self.entries[int(i)] for i in sel]
        listing = "\n".join(self._describe(e) for e in rows[:8])
        if len(rows) > 8:
            listing += f"\n… and {len(rows) - 8} more"
        if not messagebox.askyesno("Delete from Scheduler",
                                   f"Permanently delete {self._plural(len(rows))} from the "
                                   f"Scheduler?\n\n{listing}", icon="warning"):
            return
        backup = list(self.entries)
        drop = {int(i) for i in sel}
        self.entries = [e for i, e in enumerate(self.entries) if i not in drop]
        if self._persist():
            self._refresh_table()
        else:
            self.entries = backup

    def _persist(self):
        try:
            save_schedule(self.entries)
            return True
        except PermissionError:
            messagebox.showerror("Work Scheduler",
                                 f"Can't save – {SCHEDULE_FILE.name} is open in Excel. "
                                 "Close it and try again.")
            return False

    def _shown_rows(self):
        return [self.entries[int(i)] for i in self.tree.get_children()]

    def _open_export_dialog(self):
        rows = self._shown_rows()
        if not rows:
            messagebox.showinfo("Export to Excel", "There are no rows in the Scheduler to export.")
            return
        ExportDialog(self, len(rows), self.export_fields, self._export_excel, self.S)

    def _export_excel(self, columns):
        self.export_fields = columns   # remembered for the next export in this session
        path = filedialog.asksaveasfilename(
            parent=self, title="Save Excel export", initialdir=str(BASE_DIR),
            initialfile=f"Work Schedule Export {datetime.now():%d-%m-%Y %H%M}.xlsx",
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        # Sr No runs 1, 2, 3 ... within the export, even when a search has narrowed the rows.
        rows = [dict(e, **{"Sr No": i}) for i, e in enumerate(self._shown_rows(), start=1)]
        try:
            write_schedule(path, rows, columns)
        except PermissionError:
            messagebox.showerror("Export to Excel",
                                 f"Can't save – {Path(path).name} is open in Excel. "
                                 "Close it and try again.")
            return
        self._status(f"✓ Exported {self._plural(len(rows))} to {Path(path).name}.", SUCCESS)
        os.startfile(path)  # Windows only


if __name__ == "__main__":
    if sys.platform == "win32":
        try:  # crisp text on high-DPI screens
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    SchedulerApp().mainloop()
