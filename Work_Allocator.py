"""
Work Allocator
--------------
Desktop app (Tkinter) that allocates scheduled work to Resources, in the layout of the
'WWA Format' sheet of 'Work Schedule Master.xlsx':

    Rows     Resources from 'Resource Master' (alphabetical by default)
    Columns  Every date of the selected month
    Cells    Pending tasks from 'Work Schedule.xlsx', chosen in a pop-up by Client Short Name,
             Assignment and Financial Year

Allocations are saved to 'Work Allocation.xlsx' with one sheet per month in the WWA layout and
a hidden sheet listing the tasks behind each cell, so the master workbook and the schedule are
only ever read, never modified.
"""

import calendar
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

BASE_DIR = Path(__file__).resolve().parent
MASTER_FILE = BASE_DIR / "Work Schedule Master.xlsx"
SCHEDULE_FILE = BASE_DIR / "Work Schedule.xlsx"
SCHEDULE_SHEET = "Scheduler"
DATE_FMT = "%d-%b-%Y"            # 01-Apr-2026
OLD_DATE_FMTS = ("%d-%m-%Y",)    # accepted when reading older files
NA = "NA"
BG = "#EEF2F7"
CARD = "#FFFFFF"
NAVY = "#0F2A47"
NAVY_2 = "#1B3E66"
ON_NAVY_MUTED = "#A9C1DE"
ACCENT = "#2563EB"
ACCENT_SOFT = "#DBEAFE"
HOVER_SOFT = "#EFF6FF"
AMBER_SOFT = "#FEF3C7"
AMBER_DARK = "#92400E"
WARN = "#B45309"
DANGER = "#DC2626"
TEXT = "#0F172A"
MUTED = "#64748B"
FAINT = "#CBD5E1"
BORDER = "#D7DEE8"
FONT = "Segoe UI"


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

ALLOCATION_FILE = BASE_DIR / "Work Allocation.xlsx"
RESOURCE_SHEET = "Resource Master"
MANAGER_SHEET = "Manager Master"
APP_NAME = "Work Allocator"
SEPARATOR = " - "

# Layout of the 'WWA Format' sheet
TITLE_ROW, DAY_ROW, HEADER_ROW, FIRST_DATA_ROW = 5, 6, 7, 8
FIRST_DATE_COL = 5                 # column E
LEFT_HEADERS = ("Sr\nNo", "Resource", "M/F", "Manager")
LEFT_WIDTHS = (4.6, 29.6, 6.1, 9.2)
WORKDAY_WIDTH, OFF_DAY_WIDTH = 19.1, 13.7
XL_NAVY, XL_OFF_DAY = "002060", "F6C6AD"
XL_FONT = "Book Antiqua"

OFF_DAY_BG = "#FBE3D6"
TODAY_BG = "#E0EAFF"
CLEAR_LABEL = "— Clear —"
DOUBLE_CLICK_MS = 300              # a single click waits this long before opening the pop-up
NO_PENDING = "There are no tasks pending under this selection"
COMPLETE = "complete"
CHECKED, UNCHECKED = "☑", "☐"

# Hidden sheet holding the tasks behind each cell. These fields identify a schedule line.
TASKS_SHEET = "Allocated Tasks"
TASK_FIELDS = ("Client Short Name", "Assignment", "Work Type / Area", "Frequency",
               "Financial Year", "Period From", "Period To")

# Task list in the allocation pop-up: (heading, width, anchor)
PICK_COLUMNS = (
    ("Select", 60, "center"), ("Work Type / Area", 360, "w"), ("Frequency", 110, "center"),
    ("Period From", 100, "center"), ("Period To", 100, "center"),
    ("Target Date", 100, "center"), ("Due Date", 100, "center"),
    ("Work Status", 100, "center"), ("Manager", 80, "center"),
)

# Layout of the 'Resource Day Details Format' sheet: (heading, width, anchor)
DAY_TITLE_FMT = "%b %d, %Y [%A]"   # Nov 03, 2026 [Tuesday]
DAY_COLUMNS = (
    ("Client Short Name", 160, "w"), ("Manager", 80, "center"), ("Assignment", 170, "w"),
    ("Work Type / Area", 390, "w"), ("Frequency", 110, "center"), ("FY", 90, "center"),
    ("Period From", 100, "center"), ("Period To", 100, "center"),
    ("Target Date", 100, "center"), ("Due Date", 100, "center"),
    ("Client Full Name", 340, "w"), ("Manager Name", 170, "w"),
)
CLEAR_COLUMNS = (
    ("Select", 60, "center"), ("Resource", 210, "w"), ("Date", 100, "center"),
    ("Day", 100, "center"), ("Client Short Name", 150, "w"), ("Assignment", 170, "w"),
    ("Work Type / Area", 320, "w"), ("Period Covered", 190, "center"),
)
DAY_XL_WIDTHS = (20.9, 10.8, 18.9, 25.2, 17.6, 14.0, 13.8, 17.2, 16.2, 17.0, 29.8, 36.8)
DAY_PDF_WIDTHS = (56, 42, 62, 100, 50, 46, 52, 52, 52, 52, 100, 70)
DETAIL_WIDTH = 36                  # day columns in a detailed weekly export

TAB_NAMES = ("Work Allocation", "Resource's Schedule (Day)", "Weekly Schedule")
ALLOCATE_TAB, DAY_TAB, WEEK_TAB = range(3)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def month_dates(year, month):
    return [date(year, month, d) for d in range(1, calendar.monthrange(year, month)[1] + 1)]


def is_off_day(d):
    """Sundays and the 1st and 3rd Saturdays, as shaded in the 'WWA Format' sheet."""
    return d.weekday() == 6 or (d.weekday() == 5 and (d.day - 1) // 7 in (0, 2))


def shift_month(key, delta):
    year, month0 = divmod(key[0] * 12 + (key[1] - 1) + delta, 12)
    return year, month0 + 1


def sheet_name(key):
    return f"{calendar.month_abbr[key[1]]} {key[0]}"


def sort_key(name):
    return name.casefold()


# --------------------------------------------------------------------------- #
# Master data
# --------------------------------------------------------------------------- #
def read_masters():
    """Return (resources, managers): {full name: M/F} and [(short name, full name)]."""
    wb = _open_workbook_shared(MASTER_FILE)
    try:
        rows = _read_table(wb[RESOURCE_SHEET], ["Resource Full Name", "M/F"])
        names, shorts = _read_managers(wb[MANAGER_SHEET]) if MANAGER_SHEET in wb.sheetnames \
            else ([], {})
    finally:
        wb.close()
    resources = {r["Resource Full Name"]: r["M/F"] or "" for r in rows if r["Resource Full Name"]}
    return resources, [(shorts.get(n) or n, n) for n in names]


def work_label(entry):
    """'Client Short Name - Assignment' for a schedule row; None if either is blank."""
    client, assignment = _clean(entry.get("Client Short Name")), _clean(entry.get("Assignment"))
    return f"{client}{SEPARATOR}{assignment}" if client and assignment else None


def task_key(entry):
    """The fields that identify a schedule line, as text."""
    return tuple(_clean(fmt_cell(entry.get(f))) or "" for f in TASK_FIELDS)


def task_label(key):
    return f"{key[0]}{SEPARATOR}{key[1]}"


def cell_text(tasks):
    """What a WWA cell holds: one 'Client Short Name - Assignment' line per pair."""
    return "\n".join(_unique(task_label(t) for t in tasks))


def financial_year(entry):
    return _clean(entry.get("Financial Year")) or NA


def is_complete(entry):
    return (_clean(entry.get("Work Status")) or "").casefold() == COMPLETE


def read_schedule():
    """Every row of the schedule as a dict, in schedule order, read by header name."""
    if not SCHEDULE_FILE.exists():
        return []
    wb = _open_workbook_shared(SCHEDULE_FILE)
    try:
        ws = wb[SCHEDULE_SHEET] if SCHEDULE_SHEET in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        return []
    headers = [_clean(h) for h in rows[0]]
    return [{h: v for h, v in zip(headers, row) if h}
            for row in rows[1:] if any(v is not None for v in row)]


def work_options(schedule):
    """Unique 'Client Short Name - Assignment' values, A to Z."""
    return sorted(_unique(work_label(e) for e in schedule), key=sort_key)


# --------------------------------------------------------------------------- #
# Allocation storage
# --------------------------------------------------------------------------- #
class MonthData:
    """One month of allocations, keyed by Resource Full Name."""

    def __init__(self):
        self.cells = {}      # (resource, date) -> text shown in the WWA cell
        self.tasks = {}      # (resource, date) -> [task_key, ...]
        self.managers = {}   # resource -> manager short name ("" = none)
        self.gender = {}     # resource -> M/F as saved


def load_allocations():
    """Read every month sheet of 'Work Allocation.xlsx' into {(year, month): MonthData}."""
    if not ALLOCATION_FILE.exists():
        return {}
    wb = openpyxl.load_workbook(ALLOCATION_FILE, data_only=True)
    months = {}
    try:
        for ws in wb.worksheets:
            if ws.title == TASKS_SHEET:
                continue
            header = [c.value for c in ws[HEADER_ROW]] if ws.max_row >= HEADER_ROW else []
            dates = {i: v.date() for i, v in enumerate(header) if isinstance(v, datetime)}
            if not dates:
                continue
            first = next(iter(dates.values()))
            data = MonthData()
            for row in ws.iter_rows(min_row=FIRST_DATA_ROW, values_only=True):
                resource = _clean(row[1])
                if not resource:
                    continue
                data.gender[resource] = _clean(row[2]) or ""
                data.managers[resource] = _clean(row[3]) or ""
                for i, d in dates.items():
                    work = _clean(row[i]) if i < len(row) else None
                    if work:
                        data.cells[resource, d] = work
            months[first.year, first.month] = data
        if TASKS_SHEET in wb.sheetnames:
            for row in wb[TASKS_SHEET].iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue
                resource, day = _clean(row[0]), date.fromisoformat(str(row[1]))
                data = months.get((day.year, day.month))
                if data and (resource, day) in data.cells:
                    key = tuple(_clean(v) or "" for v in row[2:2 + len(TASK_FIELDS)])
                    data.tasks.setdefault((resource, day), []).append(key)
        for data in months.values():
            for cell, tasks in data.tasks.items():
                data.cells[cell] = cell_text(tasks)
    finally:
        wb.close()
    return months


def write_tasks(ws, months):
    """Fill the hidden sheet that links each cell to its schedule lines."""
    ws.append(["Resource", "Date", *TASK_FIELDS])
    for key in sorted(months):
        for (resource, day), tasks in sorted(months[key].tasks.items(), key=lambda i: i[0]):
            for task in tasks:
                ws.append([resource, day.isoformat(), *task])
    ws.sheet_state = "hidden"


def task_detail(task):
    """Client Short Name, Assignment, Work Type / Area and period covered of one task."""
    period = task[5] if task[5] == task[6] else f"{task[5]} to {task[6]}"
    return "\n".join(filter(None, (task_label(task), task[2], period)))


def month_title(key):
    return f"Weekly Work Allocation - {calendar.month_name[key[1]]} {key[0]}"


def week_title(dates):
    return (f"Weekly Work Allocation - {dates[0].strftime(DATE_FMT)} "
            f"to {dates[-1].strftime(DATE_FMT)}")


def week_of(day):
    """Monday to Sunday of the week containing day."""
    monday = day - timedelta(days=day.weekday())
    return [monday + timedelta(days=i) for i in range(7)]


def day_title(day):
    return f"Schedule for the day - {day.strftime(DAY_TITLE_FMT)}"


def write_day(path, resource, day, rows):
    """Save a day's tasks in the 'Resource Day Details Format' layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resource Day Details"
    side = Side(style="thin")
    box = Border(left=side, right=side, top=side, bottom=side)
    bold = Font(name=XL_FONT, size=11, bold=True)
    ws["A1"], ws["A2"], ws["B2"] = "Resource Day Details", resource, day_title(day)
    ws["A1"].font = ws["A2"].font = ws["B2"].font = bold
    for col, ((title, _w, anchor), width) in enumerate(zip(DAY_COLUMNS, DAY_XL_WIDTHS), start=1):
        cell = ws.cell(3, col, title)
        cell.font = Font(name=XL_FONT, size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=XL_NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = box
        ws.column_dimensions[get_column_letter(col)].width = width
        for r, row in enumerate(rows, start=4):
            cell = ws.cell(r, col, row[col - 1] or None)
            cell.font, cell.border = Font(name=XL_FONT, size=11), box
            cell.alignment = Alignment(horizontal="left" if anchor == "w" else "center",
                                       vertical="center", wrap_text=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    wb.save(path)


def write_week(path, dates, rows, detailed):
    """Save a week in the WWA layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Schedule"
    write_wwa(ws, week_title(dates), dates, rows, DETAIL_WIDTH if detailed else None)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    wb.save(path)


def write_pdf(path, title, subtitle, table, widths, off_columns=(), font_size=8):
    """Save a landscape A4 table. table[0] is the heading row."""
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    margin = 24
    usable = landscape(A4)[0] - 2 * margin
    scale = usable / sum(widths)
    body = ParagraphStyle("body", fontName="Helvetica", fontSize=font_size,
                          leading=font_size + 2)
    head = ParagraphStyle("head", parent=body, fontName="Helvetica-Bold", alignment=1,
                          textColor=colors.white)
    para = lambda text, style: Paragraph(escape(str(text or "")).replace("\n", "<br/>"), style)
    data = [[para(v, head) for v in table[0]]] + [[para(v, body) for v in r] for r in table[1:]]
    grid = Table(data, colWidths=[w * scale for w in widths], repeatRows=1, splitInRow=1)
    style = [("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
             ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#" + XL_NAVY)),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
    style += [("BACKGROUND", (c, 1), (c, -1), colors.HexColor("#" + XL_OFF_DAY))
              for c in off_columns]
    grid.setStyle(TableStyle(style))
    heading = ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=11, leading=15)
    SimpleDocTemplate(path, pagesize=landscape(A4), title=title, leftMargin=margin,
                      rightMargin=margin, topMargin=margin, bottomMargin=margin).build(
        [para(title, heading), para(subtitle, heading), Spacer(1, 8), grid])


def day_pdf(path, resource, day, rows):
    table = [[c[0] for c in DAY_COLUMNS], *rows]
    write_pdf(path, resource, day_title(day), table, DAY_PDF_WIDTHS, font_size=7)


def week_pdf(path, dates, rows):
    headings = [h.replace("\n", " ") for h in LEFT_HEADERS]
    table = [headings + [d.strftime("%A\n%d-%b-%y") for d in dates]]
    for sr, (resource, gender, manager, work) in enumerate(rows, start=1):
        table.append([sr, resource, gender, manager, *(work.get(d) for d in dates)])
    off = [len(headings) + i for i, d in enumerate(dates) if is_off_day(d)]
    write_pdf(path, "Weekly Work Allocation",
              f"{dates[0].strftime(DATE_FMT)} to {dates[-1].strftime(DATE_FMT)}",
              table, (24, 92, 30, 42, *[88] * len(dates)), off, font_size=7)


def write_wwa(ws, title, dates, rows, day_width=None):
    """Fill a sheet in the WWA layout. rows = [(resource, M/F, manager, {date: work})]."""
    side = Side(style="thin")
    box = Border(left=side, right=side, top=side, bottom=side)
    centre = Alignment(horizontal="center", vertical="center", wrap_text=True)
    navy = PatternFill("solid", fgColor=XL_NAVY)
    off_day = PatternFill("solid", fgColor=XL_OFF_DAY)
    font = Font(name=XL_FONT, size=11)

    ws.cell(TITLE_ROW, 1, title).font = Font(name=XL_FONT, size=11, bold=True)
    for col, (title, width) in enumerate(zip(LEFT_HEADERS, LEFT_WIDTHS), start=1):
        cell = ws.cell(HEADER_ROW, col, title)
        cell.font = Font(name=XL_FONT, size=11, bold=True, color="FFFFFF")
        cell.fill, cell.alignment, cell.border = navy, centre, box
        ws.column_dimensions[get_column_letter(col)].width = width
    for col, d in enumerate(dates, start=FIRST_DATE_COL):
        day = ws.cell(DAY_ROW, col, d.strftime("%A"))
        day.font, day.alignment = font, Alignment(horizontal="center")
        cell = ws.cell(HEADER_ROW, col, datetime(d.year, d.month, d.day))
        cell.font = Font(name=XL_FONT, size=11, color="FFFFFF")
        cell.fill, cell.alignment, cell.border = navy, centre, box
        cell.number_format = "d-mmm-yy"
        ws.column_dimensions[get_column_letter(col)].width = (
            day_width or (OFF_DAY_WIDTH if is_off_day(d) else WORKDAY_WIDTH))
    ws.row_dimensions[HEADER_ROW].height = 28.8

    for sr, (resource, gender, manager, work) in enumerate(rows, start=1):
        row = FIRST_DATA_ROW + sr - 1
        for col, value in enumerate((sr, resource, gender, manager), start=1):
            cell = ws.cell(row, col, value or None)
            cell.font, cell.border = font, box
            cell.alignment = (Alignment(horizontal="left", vertical="center", wrap_text=True)
                              if col == 2 else centre)
        for col, d in enumerate(dates, start=FIRST_DATE_COL):
            cell = ws.cell(row, col, work.get(d))
            cell.font, cell.alignment, cell.border = font, centre, box
            if is_off_day(d):
                cell.fill = off_day
        # Row height for the longest cell, as Excel doesn't always fit wrapped text on opening
        width = day_width or WORKDAY_WIDTH
        lines = max((sum(1 + len(line) // int(width) for line in text.split("\n"))
                     for text in work.values()), default=1)
        if lines > 1:
            ws.row_dimensions[row].height = 15.5 * lines
    ws.freeze_panes = ws.cell(FIRST_DATA_ROW, FIRST_DATE_COL)


# --------------------------------------------------------------------------- #
# Drop-down shown over a grid cell
# --------------------------------------------------------------------------- #
class CellDropdown(tk.Toplevel):
    """Searchable list that opens at a cell. options = [(label, value)]."""

    def __init__(self, parent, x, y, width, options, current, on_pick, S):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=BORDER)
        self.options, self.on_pick = options, on_pick
        self.shown = []

        body = tk.Frame(self, bg="white")
        body.pack(padx=1, pady=1, fill="both", expand=True)
        self.filter_var = tk.StringVar()
        self.entry = tk.Entry(body, textvariable=self.filter_var, font=(FONT, 10), relief="flat",
                              bg=HOVER_SOFT, fg=TEXT)
        self.entry.pack(fill="x", padx=S(6), pady=S(6), ipady=S(3))
        frame = tk.Frame(body, bg="white")
        frame.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(frame, font=(FONT, 10), relief="flat", borderwidth=0,
                                  highlightthickness=0, activestyle="none", exportselection=False,
                                  height=min(max(len(options), 1), 12), fg=TEXT,
                                  selectbackground=ACCENT, selectforeground="white")
        bar = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=bar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")

        self.filter_var.trace_add("write", lambda *_: self._render())
        self._render(current)
        self.update_idletasks()
        h = self.winfo_reqheight()
        x = max(0, min(x, self.winfo_screenwidth() - width - S(8)))
        if y + h > self.winfo_screenheight() - S(48):
            y = max(0, y - h)
        self.geometry(f"{width}x{h}+{x}+{y}")

        self.listbox.bind("<ButtonRelease-1>", lambda _e: self._choose())
        for widget in (self.entry, self.listbox):
            widget.bind("<Return>", lambda _e: self._choose())
            widget.bind("<Escape>", lambda _e: self.close())
        self.entry.bind("<Down>", lambda _e: self._move(1))
        self.entry.bind("<Up>", lambda _e: self._move(-1))
        self.bind("<ButtonPress-1>", self._click_outside, add="+")
        self.grab_set()
        self.focus_force()
        self.entry.focus_set()

    def _render(self, current=None):
        words = self.filter_var.get().casefold().split()
        self.shown = [o for o in self.options if all(w in o[0].casefold() for w in words)]
        self.listbox.delete(0, "end")
        for label, _value in self.shown:
            self.listbox.insert("end", label)
        index = next((i for i, o in enumerate(self.shown) if current and o[1] == current), 0)
        if self.shown:
            self.listbox.selection_set(index)
            self.listbox.see(index)

    def _move(self, step):
        if not self.shown:
            return "break"
        sel = self.listbox.curselection()
        index = max(0, min((sel[0] if sel else 0) + step, len(self.shown) - 1))
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(index)
        self.listbox.see(index)
        return "break"

    def _choose(self):
        sel = self.listbox.curselection()
        if sel:
            value = self.shown[sel[0]][1]
            self.close()
            self.on_pick(value)

    def _click_outside(self, event):
        x0, y0 = self.winfo_rootx(), self.winfo_rooty()
        if not (x0 <= event.x_root < x0 + self.winfo_width()
                and y0 <= event.y_root < y0 + self.winfo_height()):
            self.close()

    def close(self):
        try:
            self.grab_release()
            self.destroy()
        except tk.TclError:
            pass


# --------------------------------------------------------------------------- #
# Allocation pop-up
# --------------------------------------------------------------------------- #
class AllocateWindow(tk.Toplevel):
    """Choose Client Short Name, Assignment and Financial Year, then tick pending tasks."""

    def __init__(self, parent, resource, day, schedule, chosen, short_names, on_done, S,
                 on_clear=None):
        super().__init__(parent)
        self.title(f"Allocate Work - {resource}")
        self.configure(bg=CARD)
        self.transient(parent)
        self.schedule, self.short_names, self.on_done = schedule, short_names, on_done
        self.on_clear = on_clear
        self.chosen = list(chosen)      # task keys ticked for this cell
        self.shown = []                 # task keys in the list, in row order

        head = tk.Frame(self, bg=CARD)
        head.pack(fill="x", padx=S(20), pady=(S(16), S(10)))
        tk.Label(head, text=resource, bg=CARD, fg=NAVY,
                 font=(FONT, 14, "bold")).pack(side="left")
        tk.Label(head, text=day.strftime(DAY_TITLE_FMT), bg=CARD, fg=TEXT,
                 font=(FONT, 11, "bold")).pack(side="left", padx=(S(24), 0))

        pick = tk.Frame(self, bg=CARD)
        pick.pack(fill="x", padx=S(20), pady=(0, S(12)))
        self.combos = []
        for col, (title, width) in enumerate((("Client Short Name", 24), ("Assignment", 32),
                                              ("Financial Year", 16))):
            tk.Label(pick, text=title, bg=CARD, fg=MUTED, font=(FONT, 9, "bold")) \
                .grid(row=0, column=col, sticky="w", padx=(0, S(16)))
            combo = ttk.Combobox(pick, state="readonly", width=width, font=(FONT, 10))
            combo.grid(row=1, column=col, sticky="w", padx=(0, S(16)), pady=(S(4), 0))
            combo.bind("<<ComboboxSelected>>", lambda _e, c=col: self._on_pick(c))
            self.combos.append(combo)

        table = tk.Frame(self, bg=CARD)
        table.pack(fill="both", expand=True, padx=S(20))
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        ids = [str(i) for i in range(len(PICK_COLUMNS))]
        self.tree = ttk.Treeview(table, columns=ids, show="headings", style="Day.Treeview",
                                 height=10, selectmode="none")
        for cid, (title, width, anchor) in zip(ids, PICK_COLUMNS):
            self.tree.heading(cid, text=title)
            self.tree.column(cid, width=S(width), minwidth=S(50), anchor=anchor, stretch=False)
        yview = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        xview = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yview.set, xscrollcommand=xview.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yview.grid(row=0, column=1, sticky="ns")
        xview.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Button-1>", self._on_tree_click)

        foot = tk.Frame(self, bg=CARD)
        foot.pack(fill="x", padx=S(20), pady=S(14))
        self.count_var = tk.StringVar()
        tk.Label(foot, textvariable=self.count_var, bg=CARD, fg=MUTED,
                 font=(FONT, 10)).pack(side="left")
        ttk.Button(foot, text="✓ Allocate", style="Accent.TButton",
                   command=self._allocate).pack(side="right")
        ttk.Button(foot, text="Cancel", style="Ghost.TButton",
                   command=self.destroy).pack(side="right", padx=(0, S(10)))
        if on_clear:
            ttk.Button(foot, text="Clear Tasks", style="Danger.TButton",
                       command=self._clear).pack(side="right", padx=(0, S(10)))

        self.bind("<Escape>", lambda _e: self.destroy())
        self.combos[0].configure(values=self._values(0))
        if self.chosen:                 # reopen on the selection already allocated
            task = self.chosen[0]
            for col, value in enumerate((task[0], task[1], task[4] or NA)):
                self.combos[col].configure(values=self._values(col))
                self.combos[col].set(value)
            self._render()
        self._update_count()

        self.update_idletasks()
        w = min(S(sum(c[1] for c in PICK_COLUMNS)) + S(60), self.winfo_screenwidth() - S(80))
        h = min(self.winfo_reqheight(), self.winfo_screenheight() - S(120))
        self.geometry(f"{w}x{h}+{(self.winfo_screenwidth() - w) // 2}"
                      f"+{max(0, (self.winfo_screenheight() - h) // 3)}")
        self.grab_set()
        self.focus_force()

    def _matching(self, levels):
        """Schedule lines that match the first `levels` selectors."""
        fields = (lambda e: _clean(e.get("Client Short Name")),
                  lambda e: _clean(e.get("Assignment")), financial_year)
        rows = self.schedule
        for get, combo in list(zip(fields, self.combos))[:levels]:
            rows = [e for e in rows if get(e) == combo.get()]
        return rows

    def _values(self, level):
        get = (lambda e: _clean(e.get("Client Short Name")),
               lambda e: _clean(e.get("Assignment")), financial_year)[level]
        return sorted(_unique(get(e) for e in self._matching(level)), key=sort_key)

    def _on_pick(self, level):
        for col in range(level + 1, len(self.combos)):
            self.combos[col].set("")
            self.combos[col].configure(values=[])
        if level + 1 < len(self.combos):
            self.combos[level + 1].configure(values=self._values(level + 1))
            self._fill([])
        else:
            self._render(announce=True)

    def _render(self, announce=False):
        rows = [e for e in self._matching(len(self.combos)) if not is_complete(e)]
        self._fill(rows)
        if announce and not rows:
            messagebox.showinfo(APP_NAME, NO_PENDING, parent=self)

    def _fill(self, rows):
        self.tree.delete(*self.tree.get_children())
        self.shown = [task_key(e) for e in rows]
        for e, key in zip(rows, self.shown):
            manager = _clean(e.get("Manager"))
            self.tree.insert("", "end", values=(
                CHECKED if key in self.chosen else UNCHECKED,
                *(fmt_cell(e.get(f)) for f in ("Work Type / Area", "Frequency", "Period From",
                                               "Period To", "Target Date", "Due Date",
                                               "Work Status")),
                self.short_names.get(manager) or e.get("Manager Short Name") or ""))

    def _on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        key = self.shown[self.tree.index(item)]
        if key in self.chosen:
            self.chosen.remove(key)
        else:
            self.chosen.append(key)
        self.tree.set(item, "0", CHECKED if key in self.chosen else UNCHECKED)
        self._update_count()

    def _update_count(self):
        n = len(self.chosen)
        self.count_var.set(f"{n} task{'' if n == 1 else 's'} selected for this day")

    def _allocate(self):
        self.destroy()
        self.on_done(self.chosen)

    def _clear(self):
        self.destroy()
        self.on_clear()


# --------------------------------------------------------------------------- #
# Clear allocations
# --------------------------------------------------------------------------- #
class ClearWindow(tk.Toplevel):
    """Tick allocated work in a month and clear it. items = [(resource, date, task or None, work)]."""

    def __init__(self, parent, title, items, on_clear, S):
        super().__init__(parent)
        self.title("Clear Allocations")
        self.configure(bg=CARD)
        self.transient(parent)
        self.items, self.on_clear = items, on_clear
        self.ticked = set()     # indexes into items
        self.shown = []         # indexes into items, in row order

        head = tk.Frame(self, bg=CARD)
        head.pack(fill="x", padx=S(20), pady=(S(16), S(10)))
        tk.Label(head, text=f"Clear Allocations - {title}", bg=CARD, fg=NAVY,
                 font=(FONT, 14, "bold")).pack(side="left")
        self.search_var = tk.StringVar()
        tk.Entry(head, textvariable=self.search_var, font=(FONT, 10), relief="flat",
                 bg=HOVER_SOFT, fg=TEXT, width=30).pack(side="right", ipady=S(4))
        tk.Label(head, text="Search", bg=CARD, fg=MUTED,
                 font=(FONT, 9, "bold")).pack(side="right", padx=(0, S(8)))

        table = tk.Frame(self, bg=CARD)
        table.pack(fill="both", expand=True, padx=S(20))
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        ids = [str(i) for i in range(len(CLEAR_COLUMNS))]
        self.tree = ttk.Treeview(table, columns=ids, show="headings", style="Day.Treeview",
                                 height=14, selectmode="none")
        for cid, (heading, width, anchor) in zip(ids, CLEAR_COLUMNS):
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=S(width), minwidth=S(50), anchor=anchor, stretch=False)
        yview = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        xview = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yview.set, xscrollcommand=xview.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yview.grid(row=0, column=1, sticky="ns")
        xview.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Button-1>", self._on_tree_click)

        foot = tk.Frame(self, bg=CARD)
        foot.pack(fill="x", padx=S(20), pady=S(14))
        ttk.Button(foot, text="Select All", style="Ghost.TButton",
                   command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(foot, text="Select None", style="Ghost.TButton",
                   command=lambda: self._set_all(False)).pack(side="left", padx=(S(10), 0))
        self.count_var = tk.StringVar()
        tk.Label(foot, textvariable=self.count_var, bg=CARD, fg=MUTED,
                 font=(FONT, 10)).pack(side="left", padx=S(16))
        self.clear_btn = ttk.Button(foot, text="Clear Selected", style="Danger.TButton",
                                    command=self._clear)
        self.clear_btn.pack(side="right")
        ttk.Button(foot, text="Cancel", style="Ghost.TButton",
                   command=self.destroy).pack(side="right", padx=(0, S(10)))

        self.search_var.trace_add("write", lambda *_: self._render())
        self.bind("<Escape>", lambda _e: self.destroy())
        self._render()
        self.update_idletasks()
        w = min(S(sum(c[1] for c in CLEAR_COLUMNS)) + S(60), self.winfo_screenwidth() - S(80))
        h = min(self.winfo_reqheight(), self.winfo_screenheight() - S(120))
        self.geometry(f"{w}x{h}+{(self.winfo_screenwidth() - w) // 2}"
                      f"+{max(0, (self.winfo_screenheight() - h) // 3)}")
        self.grab_set()
        self.focus_force()

    @staticmethod
    def _values(item):
        resource, day, task, work = item
        client, assignment, work_type, period = "", work, "", ""
        if task:
            client, assignment, work_type = task[0], task[1], task[2]
            period = task[5] if task[5] == task[6] else f"{task[5]} to {task[6]}"
        return resource, day.strftime(DATE_FMT), day.strftime("%A"), client, assignment, \
            work_type, period

    def _render(self):
        words = self.search_var.get().casefold().split()
        self.tree.delete(*self.tree.get_children())
        self.shown = []
        for i, item in enumerate(self.items):
            values = self._values(item)
            text = " ".join(values).casefold()
            if all(w in text for w in words):
                self.shown.append(i)
                self.tree.insert("", "end", values=(
                    CHECKED if i in self.ticked else UNCHECKED, *values))
        self._update_count()

    def _on_tree_click(self, event):
        row = self.tree.identify_row(event.y)
        if not row:
            return
        i = self.shown[self.tree.index(row)]
        self.ticked.symmetric_difference_update({i})
        self.tree.set(row, "0", CHECKED if i in self.ticked else UNCHECKED)
        self._update_count()

    def _set_all(self, value):
        """Applies to the rows shown, so a search can pick out a Resource or a client."""
        (self.ticked.update if value else self.ticked.difference_update)(self.shown)
        for row in self.tree.get_children():
            self.tree.set(row, "0", CHECKED if value else UNCHECKED)
        self._update_count()

    def _update_count(self):
        n = len(self.ticked)
        self.count_var.set(f"{n} of {len(self.items)} selected")
        self.clear_btn.state(["!disabled"] if n else ["disabled"])

    def _clear(self):
        n = len(self.ticked)
        if n and messagebox.askyesno(
                "Clear Allocations", f"Clear {n} allocation{'' if n == 1 else 's'}?", parent=self):
            chosen = [self.items[i] for i in sorted(self.ticked)]
            self.destroy()
            self.on_clear(chosen)


# --------------------------------------------------------------------------- #
# Day details, in the layout of the 'Resource Day Details Format' sheet
# --------------------------------------------------------------------------- #
class DayDetails(tk.Toplevel):
    """The tasks allocated to a Resource on one day."""

    def __init__(self, parent, resource, day, rows, note, S, on_change=None, on_clear=None):
        super().__init__(parent)
        self.title(f"{resource} - {day.strftime(DAY_TITLE_FMT)}")
        self.configure(bg=CARD)
        self.transient(parent)

        head = tk.Frame(self, bg=CARD)
        head.pack(fill="x", padx=S(20), pady=(S(16), S(10)))
        tk.Label(head, text=resource, bg=CARD, fg=NAVY,
                 font=(FONT, 14, "bold")).pack(side="left")
        tk.Label(head, text=day_title(day), bg=CARD,
                 fg=TEXT, font=(FONT, 11, "bold")).pack(side="left", padx=(S(24), 0))
        ttk.Button(head, text="Close", style="Ghost.TButton",
                   command=self.destroy).pack(side="right")
        stem = f"Day Schedule {resource} {day.isoformat()}"
        ttk.Button(head, text="Export to PDF", style="Ghost.TButton",
                   command=lambda: parent.save_as(self, stem, "pdf", lambda path: day_pdf(
                       path, resource, day, rows))).pack(side="right", padx=(0, S(10)))
        ttk.Button(head, text="Export to Excel", style="Ghost.TButton",
                   command=lambda: parent.save_as(self, stem, "xlsx", lambda path: write_day(
                       path, resource, day, rows))).pack(side="right", padx=(0, S(10)))
        if on_clear:
            ttk.Button(head, text="Clear Tasks", style="Danger.TButton",
                       command=on_clear).pack(side="right", padx=(0, S(10)))
        if on_change:
            ttk.Button(head, text="Change Resource / Day", style="Ghost.TButton",
                       command=on_change).pack(side="right", padx=(0, S(10)))
        if note:
            tk.Label(self, text=note, bg=AMBER_SOFT, fg=AMBER_DARK, font=(FONT, 10),
                     anchor="w", padx=S(12), pady=S(8)).pack(fill="x", padx=S(20),
                                                             pady=(0, S(10)))

        table = tk.Frame(self, bg=CARD)
        table.pack(fill="both", expand=True, padx=S(20), pady=(0, S(20)))
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        ids = [str(i) for i in range(len(DAY_COLUMNS))]
        self.tree = ttk.Treeview(table, columns=ids, show="headings", style="Day.Treeview",
                                 height=min(max(len(rows), 1), 16))
        for cid, (title, width, anchor) in zip(ids, DAY_COLUMNS):
            self.tree.heading(cid, text=title)
            self.tree.column(cid, width=S(width), minwidth=S(60), anchor=anchor, stretch=False)
        for row in rows:
            self.tree.insert("", "end", values=row)
        yview = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        xview = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yview.set, xscrollcommand=xview.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yview.grid(row=0, column=1, sticky="ns")
        xview.grid(row=1, column=0, sticky="ew")

        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        w = min(S(sum(c[1] for c in DAY_COLUMNS)) + S(60), self.winfo_screenwidth() - S(80))
        h = min(self.winfo_reqheight(), self.winfo_screenheight() - S(120))
        self.geometry(f"{w}x{h}+{(self.winfo_screenwidth() - w) // 2}"
                      f"+{max(0, (self.winfo_screenheight() - h) // 3)}")
        self.focus_force()


# --------------------------------------------------------------------------- #
# Calendar pop-up: pick a day (and Resource), or a Monday-to-Sunday week
# --------------------------------------------------------------------------- #
class CalendarPick(tk.Toplevel):
    """resources=None picks a week: clicking any day selects its whole week."""

    def __init__(self, parent, title, on_show, S, resources=None, selected=None):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=CARD)
        self.transient(parent)
        self.resizable(False, False)
        self.S, self.on_show, self.pick_week = S, on_show, resources is None
        self.selected = selected or date.today()
        self.year, self.month = self.selected.year, self.selected.month

        body = tk.Frame(self, bg=CARD)
        body.pack(padx=S(20), pady=S(16))
        self.resource = None
        if resources is not None:
            tk.Label(body, text="Resource", bg=CARD, fg=MUTED, font=(FONT, 9, "bold")) \
                .pack(anchor="w")
            self.resource = ttk.Combobox(body, state="readonly", values=resources, width=34,
                                         font=(FONT, 10))
            self.resource.pack(fill="x", pady=(S(4), S(12)))
            self.resource.bind("<<ComboboxSelected>>", lambda _e: self._sync())
        tk.Label(body, text="Week" if self.pick_week else "Day", bg=CARD, fg=MUTED,
                 font=(FONT, 9, "bold")).pack(anchor="w", pady=(0, S(4)))

        head = tk.Frame(body, bg=NAVY)
        head.pack(fill="x")
        for text, delta, side in (("‹", -1, "left"), ("›", 1, "right")):
            lbl = tk.Label(head, text=text, bg=NAVY, fg="white", font=(FONT, 12, "bold"),
                           padx=S(12), pady=S(5), cursor="hand2")
            lbl.bind("<Button-1>", lambda _e, d=delta: self._shift(d))
            lbl.pack(side=side)
        self.title_lbl = tk.Label(head, bg=NAVY, fg="white", font=(FONT, 11, "bold"))
        self.title_lbl.pack(side="left", expand=True)
        self.days = tk.Frame(body, bg=CARD, pady=S(6))
        self.days.pack()

        self.summary = tk.Label(body, bg=CARD, fg=TEXT, font=(FONT, 10, "bold"))
        self.summary.pack(anchor="w", pady=(S(6), S(10)))
        foot = tk.Frame(body, bg=CARD)
        foot.pack(fill="x")
        self.show_btn = ttk.Button(foot, text="Show", style="Accent.TButton", command=self._show)
        self.show_btn.pack(side="right")
        ttk.Button(foot, text="Cancel", style="Ghost.TButton",
                   command=self.destroy).pack(side="right", padx=(0, S(10)))

        self._render()
        self._sync()
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f"+{(self.winfo_screenwidth() - w) // 2}"
                      f"+{max(0, (self.winfo_screenheight() - h) // 3)}")
        self.grab_set()
        self.focus_force()

    def _in_selection(self, d):
        return d in week_of(self.selected) if self.pick_week else d == self.selected

    def _render(self):
        for w in self.days.winfo_children():
            w.destroy()
        self.title_lbl.configure(text=f"{calendar.month_name[self.month]} {self.year}")
        for i, name in enumerate(("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")):
            tk.Label(self.days, text=name, bg=CARD, fg=MUTED, width=4,
                     font=(FONT, 9, "bold")).grid(row=0, column=i)
        weeks = calendar.Calendar(firstweekday=0).monthdatescalendar(self.year, self.month)
        for r, week in enumerate(weeks, start=1):
            for c, d in enumerate(week):
                bg, fg = CARD, TEXT if d.month == self.month else FAINT
                if d == date.today():
                    fg = ACCENT
                if self._in_selection(d):
                    bg, fg = ACCENT, "white"
                lbl = tk.Label(self.days, text=d.day, bg=bg, fg=fg, width=4, cursor="hand2",
                               font=(FONT, 10), pady=self.S(3))
                lbl.grid(row=r, column=c, padx=0 if self.pick_week else 1, pady=1)
                lbl.bind("<Button-1>", lambda _e, d=d: self._pick(d))

    def _shift(self, delta):
        self.year, self.month = shift_month((self.year, self.month), delta)
        self._render()

    def _pick(self, d):
        self.selected = d
        self.year, self.month = d.year, d.month
        self._render()
        self._sync()

    def _sync(self):
        if self.pick_week:
            week = week_of(self.selected)
            text = f"{week[0].strftime(DATE_FMT)} to {week[-1].strftime(DATE_FMT)}"
        else:
            text = self.selected.strftime(DAY_TITLE_FMT)
        self.summary.configure(text=text)
        ready = self.pick_week or bool(self.resource.get())
        self.show_btn.state(["!disabled"] if ready else ["disabled"])

    def _show(self):
        resource = self.resource.get() if self.resource else None
        if not self.pick_week and not resource:
            return
        self.destroy()
        self.on_show(*([resource, self.selected] if resource else [self.selected]))


# --------------------------------------------------------------------------- #
# Weekly Schedule, in the WWA layout from Monday to Sunday
# --------------------------------------------------------------------------- #
class WeeklySchedule(tk.Toplevel):
    def __init__(self, parent, S):
        super().__init__(parent)
        self.app, self.S = parent, S
        self.title("Weekly Schedule")
        self.configure(bg=BG)
        self.left_widths = (S(40), S(190), S(40), S(70))
        self.detailed = False
        self.mode = "this"
        self.dates = week_of(date.today())
        self.rows = []

        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=S(20), pady=(S(14), S(8)))
        self.buttons = {}
        for mode, text, command in (("this", "This Week", self._this_week),
                                    ("next", "Next Week", self._next_week),
                                    ("custom", "Custom Week", self._custom_week)):
            self.buttons[mode] = ttk.Button(bar, text=text, command=command)
            self.buttons[mode].pack(side="left", padx=(0, S(8)))
        self.range_var = tk.StringVar()
        tk.Label(bar, textvariable=self.range_var, bg=BG, fg=NAVY,
                 font=(FONT, 12, "bold")).pack(side="left", padx=S(12))
        ttk.Button(bar, text="Export to PDF", style="Ghost.TButton",
                   command=lambda: self._export("pdf")).pack(side="right")
        ttk.Button(bar, text="Export to Excel", style="Ghost.TButton",
                   command=lambda: self._export("xlsx")).pack(side="right", padx=(0, S(8)))
        self.detail_btn = ttk.Button(bar, style="Ghost.TButton", command=self._toggle_detail)
        self.detail_btn.pack(side="right", padx=(0, S(8)))

        card = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="both", expand=True, padx=S(20), pady=(0, S(20)))
        card.rowconfigure(1, weight=1)
        card.columnconfigure(0, weight=1)
        self.head_h = S(50)
        self.top = tk.Canvas(card, bg="white", highlightthickness=0, height=self.head_h)
        self.body = tk.Canvas(card, bg="white", highlightthickness=0)
        self.top.grid(row=0, column=0, sticky="nsew")
        self.body.grid(row=1, column=0, sticky="nsew")
        yview = ttk.Scrollbar(card, orient="vertical", command=self.body.yview)
        xview = ttk.Scrollbar(card, orient="horizontal", command=self._xview)
        yview.grid(row=1, column=1, sticky="ns")
        xview.grid(row=2, column=0, sticky="ew")
        self.body.configure(yscrollcommand=yview.set, xscrollcommand=xview.set)
        self.body.bind("<MouseWheel>", lambda e: self.body.yview_scroll(
            -1 if e.delta > 0 else 1, "units"))
        self.body.bind("<Configure>", lambda _e: self.render())
        self.bind("<Escape>", lambda _e: self.destroy())

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{sw - S(80)}x{sh - S(140)}+{S(40)}+{S(30)}")
        self.render()
        self.focus_force()

    def _xview(self, *args):
        self.body.xview(*args)
        self.top.xview(*args)

    def _this_week(self):
        self.mode, self.dates = "this", week_of(date.today())
        self.render()

    def _next_week(self):
        self.mode = "next"
        self.dates = week_of(date.today() + timedelta(days=7))
        self.render()

    def _custom_week(self):
        CalendarPick(self, "Custom Week", self._show_custom, self.S, selected=self.dates[0])

    def _show_custom(self, day):
        self.mode, self.dates = "custom", week_of(day)
        self.render()

    def _toggle_detail(self):
        self.detailed = not self.detailed
        self.render()

    def _export(self, kind):
        stem = f"Weekly Schedule {self.dates[0].isoformat()}"
        if self.detailed:
            stem += " Detailed"
        if kind == "pdf":
            writer = lambda path: week_pdf(path, self.dates, self.rows)
        else:
            writer = lambda path: write_week(path, self.dates, self.rows, self.detailed)
        self.app.save_as(self, stem, kind, writer)

    def render(self):
        S = self.S
        self.rows = self.app.wwa_rows(self.dates, self.detailed)
        self.range_var.set(f"{self.dates[0].strftime(DATE_FMT)} to "
                           f"{self.dates[-1].strftime(DATE_FMT)}")
        for mode, button in self.buttons.items():
            button.configure(style="Accent.TButton" if mode == self.mode else "Ghost.TButton")
        self.detail_btn.configure(text="Show Concise" if self.detailed else "Show Detailed")

        left_w = sum(self.left_widths)
        available = self.body.winfo_width() if self.body.winfo_width() > 1 else S(1200)
        day_w = max(S(130), (available - left_w) // len(self.dates))
        total_w = left_w + day_w * len(self.dates)

        top = self.top
        top.delete("all")
        x = 0
        for title, w in zip(LEFT_HEADERS, self.left_widths):
            top.create_rectangle(x, 0, x + w, self.head_h, fill=NAVY, outline=NAVY_2)
            top.create_text(x + w / 2, self.head_h / 2, text=title, fill="white",
                            justify="center", font=(FONT, 10, "bold"))
            x += w
        today = date.today()
        for d in self.dates:
            top.create_rectangle(x, 0, x + day_w, self.head_h, outline=NAVY_2,
                                 fill=ACCENT if d == today else NAVY)
            top.create_text(x + day_w / 2, self.head_h * 0.3, text=d.strftime("%A"),
                            fill="white" if d == today else ON_NAVY_MUTED, font=(FONT, 9))
            top.create_text(x + day_w / 2, self.head_h * 0.68, fill="white",
                            text=d.strftime("%d-%b-%y").lstrip("0"), font=(FONT, 10, "bold"))
            x += day_w
        top.configure(scrollregion=(0, 0, total_w, self.head_h))

        body = self.body
        body.delete("all")
        y = 0
        for sr, (resource, gender, manager, work) in enumerate(self.rows, start=1):
            # Texts first: the tallest one sets the height of the row
            cells = [(self.left_widths[0], str(sr), None), (self.left_widths[1], resource, None),
                     (self.left_widths[2], gender, None), (self.left_widths[3], manager, None)]
            cells += [(day_w, work.get(d, ""), d) for d in self.dates]
            x, bottom, drawn = 0, y + S(38), []
            for col, (w, text, d) in enumerate(cells):
                item = body.create_text(x + S(6) if col == 1 else x + w / 2, y + S(8), text=text,
                                        anchor="nw" if col == 1 else "n", fill=TEXT,
                                        justify="left" if col == 1 else "center",
                                        width=w - S(12),
                                        font=(FONT, 10, "bold") if col == 1 else (FONT, 9))
                bbox = body.bbox(item)
                if bbox:
                    bottom = max(bottom, bbox[3] + S(8))
                drawn.append((x, w, d))
                x += w
            for x, w, d in drawn:
                fill = "white"
                if d and work.get(d):
                    fill = ACCENT_SOFT
                elif d and is_off_day(d):
                    fill = OFF_DAY_BG
                item = body.create_rectangle(x, y, x + w, bottom, fill=fill, outline=BORDER)
                body.tag_lower(item)
            y = bottom
        body.configure(scrollregion=(0, 0, total_w, y), yscrollincrement=S(30))


# --------------------------------------------------------------------------- #
# Application
# --------------------------------------------------------------------------- #
class AllocatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.configure(bg=BG)
        self._scale = self.winfo_fpixels("1i") / 96
        self.geometry(f"{self.S(1360)}x{self.S(820)}")
        self.minsize(self.S(900), self.S(560))
        try:
            self.state("zoomed")
        except tk.TclError:
            pass

        S = self.S
        self.left_widths = (S(42), S(230), S(46), S(84))
        self.date_width = S(128)
        self.row_height = S(54)
        self.head_height = S(52)

        today = date.today()
        self.month = (today.year, today.month)
        self.months = {}          # (year, month) -> MonthData
        self.resources = {}       # full name -> M/F
        self.managers = []        # [(short name, full name)]
        self.schedule = []        # rows of Work Schedule.xlsx
        self.schedule_keys = set()
        self.work_options = []
        self._click_job = None
        self.sort_descending = False
        self.row_names = []
        self.dirty = False
        self.dropdown = None
        self.details = None
        self._source_mtimes = None
        self._file_mtime = None

        self.month_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.master_info_var = tk.StringVar()
        self.stat_vars = {}

        self.weekly = None
        self.tab_labels = []

        self._init_style()
        self._build_header()
        self._build_tabs()
        self._build_toolbar()
        self._build_grid()

        self._load_sources(initial=True)
        self._load_file()
        self._redraw()
        self.after(100, self._scroll_to_today)
        self.after(3000, self._watch_sources)
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
        st.configure("Accent.TButton", background=ACCENT, foreground="white",
                     font=(FONT, 10, "bold"), padding=(S(18), S(8)), borderwidth=0,
                     focuscolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT, bordercolor=ACCENT)
        dis, act = ("disabled", "#A9BEF0"), ("active", "#1D4ED8")
        st.map("Accent.TButton", background=[dis, act], lightcolor=[dis, act],
               darkcolor=[dis, act], bordercolor=[dis, act], foreground=[("disabled", "white")])
        st.configure("Ghost.TButton", background="white", foreground=NAVY, font=(FONT, 10),
                     padding=(S(12), S(7)), bordercolor=BORDER, lightcolor="white",
                     darkcolor="white", focuscolor="white")
        st.map("Ghost.TButton", background=[("active", HOVER_SOFT)],
               lightcolor=[("active", HOVER_SOFT)], darkcolor=[("active", HOVER_SOFT)],
               bordercolor=[("active", ACCENT)])
        st.configure("Danger.TButton", background="white", foreground=DANGER, font=(FONT, 10),
                     padding=(S(12), S(7)), bordercolor=BORDER, lightcolor="white",
                     darkcolor="white", focuscolor="white")
        st.map("Danger.TButton", background=[("active", "#FEF2F2")],
               lightcolor=[("active", "#FEF2F2")], darkcolor=[("active", "#FEF2F2")],
               bordercolor=[("active", DANGER)], foreground=[("disabled", FAINT)])
        st.configure("TCombobox", padding=S(5), fieldbackground="white", background="white",
                     foreground=TEXT, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                     arrowcolor=NAVY, arrowsize=S(14))
        st.map("TCombobox", fieldbackground=[("readonly", "white")],
               selectbackground=[("readonly", "white")], selectforeground=[("readonly", TEXT)],
               bordercolor=[("focus", ACCENT), ("hover", ACCENT)])
        self.option_add("*TCombobox*Listbox.font", (FONT, 10))
        self.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
        self.option_add("*TCombobox*Listbox.selectForeground", "white")
        st.configure("Day.Treeview", background="white", fieldbackground="white",
                     foreground=TEXT, rowheight=S(30), borderwidth=0, font=(FONT, 10))
        st.map("Day.Treeview", background=[("selected", ACCENT_SOFT)],
               foreground=[("selected", TEXT)])
        st.configure("Day.Treeview.Heading", background=NAVY, foreground="white",
                     font=(FONT, 10, "bold"), relief="flat", padding=(S(8), S(8)),
                     borderwidth=0, lightcolor=NAVY, darkcolor=NAVY, bordercolor=NAVY)
        st.map("Day.Treeview.Heading", background=[("active", NAVY_2)])
        for orient in ("Vertical", "Horizontal"):
            st.configure(f"{orient}.TScrollbar", background=FAINT, troughcolor="#F7F9FC",
                         bordercolor="#F7F9FC", lightcolor=FAINT, darkcolor=FAINT,
                         arrowcolor=MUTED, gripcount=0)
            st.map(f"{orient}.TScrollbar", background=[("active", "#94A3B8")])

    # ---- layout --------------------------------------------------------- #
    def _build_header(self):
        S = self.S
        bar = tk.Frame(self, bg=NAVY)
        bar.pack(fill="x")
        tk.Frame(self, bg=ACCENT, height=S(4)).pack(fill="x")

        left = tk.Frame(bar, bg=NAVY)
        left.pack(side="left", padx=S(28), pady=S(14))
        tk.Label(left, text=APP_NAME, bg=NAVY, fg="white",
                 font=(FONT, 22, "bold")).pack(anchor="w")
        tk.Label(left, text="Allocate scheduled work to Resources, date by date",
                 bg=NAVY, fg=ON_NAVY_MUTED, font=(FONT, 10)).pack(anchor="w", pady=(S(2), 0))
        tk.Label(left, textvariable=self.master_info_var, bg=NAVY, fg="#7DD3A8",
                 font=(FONT, 9)).pack(anchor="w", pady=(S(4), 0))

        stats = tk.Frame(bar, bg=NAVY)
        stats.pack(side="right", padx=S(28))
        for key, label in (("resources", "Resources"), ("work", "Work Options"),
                           ("allocated", "Allocated This Month")):
            chip = tk.Frame(stats, bg=NAVY_2, padx=S(20), pady=S(8))
            chip.pack(side="left", padx=(S(10), 0))
            var = tk.StringVar(value="0")
            self.stat_vars[key] = var
            tk.Label(chip, textvariable=var, bg=NAVY_2, fg="white",
                     font=(FONT, 18, "bold")).pack(anchor="w")
            tk.Label(chip, text=label.upper(), bg=NAVY_2, fg=ON_NAVY_MUTED,
                     font=(FONT, 8, "bold")).pack(anchor="w")

    def _build_tabs(self):
        S = self.S
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=S(24), pady=(S(12), 0))
        for i, name in enumerate(TAB_NAMES):
            lbl = tk.Label(bar, text=name, font=(FONT, 10, "bold"), padx=S(18), pady=S(8),
                           cursor="hand2")
            lbl.pack(side="left", padx=(0, S(4)))
            lbl.bind("<Button-1>", lambda _e, tab=i: self._open_tab(tab))
            self.tab_labels.append(lbl)
        self._set_tab(ALLOCATE_TAB)

    def _set_tab(self, tab):
        for i, lbl in enumerate(self.tab_labels):
            lbl.configure(bg=NAVY if i == tab else "white", fg="white" if i == tab else NAVY)

    def _track_tab(self, window, tab):
        """Show a tab as active while its window is open."""
        self._set_tab(tab)
        window.bind("<Destroy>", lambda e: e.widget is window and self._set_tab(ALLOCATE_TAB),
                    add="+")

    def _open_tab(self, tab):
        if tab == DAY_TAB:
            self._pick_day()
        elif tab == WEEK_TAB:
            self._open_weekly()
        else:
            self._set_tab(ALLOCATE_TAB)

    def _pick_day(self):
        if self.details is not None and self.details.winfo_exists():
            self.details.destroy()
        names = sorted(set(self.resources), key=sort_key)
        self._track_tab(CalendarPick(self, "Resource's Schedule (Day)",
                                     lambda r, d: self._show_day(r, d, from_tab=True), self.S,
                                     resources=names), DAY_TAB)

    def _open_weekly(self):
        if self.weekly is not None and self.weekly.winfo_exists():
            self.weekly.lift()
            self.weekly.focus_force()
            return
        self.weekly = WeeklySchedule(self, self.S)
        self._track_tab(self.weekly, WEEK_TAB)

    def _build_toolbar(self):
        S = self.S
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=S(24), pady=(S(14), S(10)))
        ttk.Button(bar, text="‹", style="Ghost.TButton", width=3,
                   command=lambda: self._go_to(shift_month(self.month, -1))).pack(side="left")
        tk.Label(bar, textvariable=self.month_var, bg=BG, fg=NAVY, width=16,
                 font=(FONT, 14, "bold")).pack(side="left")
        ttk.Button(bar, text="›", style="Ghost.TButton", width=3,
                   command=lambda: self._go_to(shift_month(self.month, 1))).pack(side="left")
        ttk.Button(bar, text="This Month", style="Ghost.TButton",
                   command=self._go_to_today).pack(side="left", padx=(S(10), 0))
        tk.Label(bar, textvariable=self.status_var, bg=BG, fg=MUTED,
                 font=(FONT, 10)).pack(side="left", padx=S(16))

        self.save_btn = ttk.Button(bar, text="✓ Save", style="Accent.TButton", command=self._save)
        self.save_btn.pack(side="right")
        ttk.Button(bar, text="Export to Excel", style="Ghost.TButton",
                   command=self._export).pack(side="right", padx=(0, S(10)))
        ttk.Button(bar, text="Clear Allocations", style="Danger.TButton",
                   command=self._open_clear).pack(side="right", padx=(0, S(10)))
        ttk.Button(bar, text="⟳ Reload", style="Ghost.TButton",
                   command=self._reload).pack(side="right", padx=(0, S(10)))
        self.bind("<Control-s>", lambda _e: self._save())

    def _build_grid(self):
        S = self.S
        card = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="both", expand=True, padx=S(24), pady=(0, S(20)))
        card.rowconfigure(1, weight=1)
        card.columnconfigure(1, weight=1)

        left_w = sum(self.left_widths)
        make = lambda **kw: tk.Canvas(card, bg="white", highlightthickness=0, **kw)
        self.corner = make(width=left_w, height=self.head_height)
        self.top = make(height=self.head_height)
        self.side = make(width=left_w)
        self.body = make()
        self.corner.grid(row=0, column=0, sticky="nsew")
        self.top.grid(row=0, column=1, sticky="nsew")
        self.side.grid(row=1, column=0, sticky="nsew")
        self.body.grid(row=1, column=1, sticky="nsew")

        yview = ttk.Scrollbar(card, orient="vertical", command=self._yview)
        xview = ttk.Scrollbar(card, orient="horizontal", command=self._xview)
        yview.grid(row=1, column=2, sticky="ns")
        xview.grid(row=2, column=1, sticky="ew")
        self.body.configure(yscrollcommand=yview.set, xscrollcommand=xview.set)

        self.body.bind("<Button-1>", self._on_body_click)
        self.body.bind("<Double-Button-1>", self._on_body_double_click)
        self.side.bind("<Button-1>", self._on_side_click)
        self.corner.bind("<Button-1>", self._on_corner_click)
        for canvas in (self.body, self.side):
            canvas.bind("<MouseWheel>", self._on_wheel)
            canvas.bind("<Shift-MouseWheel>", self._on_shift_wheel)
        self.top.bind("<Shift-MouseWheel>", self._on_shift_wheel)

    def _yview(self, *args):
        self.body.yview(*args)
        self.side.yview(*args)

    def _xview(self, *args):
        self.body.xview(*args)
        self.top.xview(*args)

    def _on_wheel(self, event):
        self._yview("scroll", -1 if event.delta > 0 else 1, "units")

    def _on_shift_wheel(self, event):
        self._xview("scroll", -1 if event.delta > 0 else 1, "units")

    # ---- data ----------------------------------------------------------- #
    def _source_stamp(self):
        return tuple(p.stat().st_mtime if p.exists() else None
                     for p in (MASTER_FILE, SCHEDULE_FILE))

    def _load_sources(self, initial=False, auto=False):
        """Read Resource Master, Manager Master and the schedule; keep what we had on failure."""
        try:
            self._source_mtimes = self._source_stamp()
            self.resources, self.managers = read_masters()
            self.schedule = read_schedule()
            self.work_options = work_options(self.schedule)
            self.schedule_keys = {task_key(e) for e in self.schedule}
        except (OSError, KeyError, ValueError) as exc:
            if auto:
                return
            messagebox.showerror(APP_NAME, f"Could not read the lists:\n{exc}")
            if initial and not self.resources:
                self.master_info_var.set("Master lists not loaded")
            return
        self.master_info_var.set(
            f"Lists updated {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}  •  "
            f"{MASTER_FILE.name}  •  {SCHEDULE_FILE.name}")

    def _watch_sources(self):
        """Check every few seconds whether the master or the schedule was saved."""
        try:
            if self._source_stamp() != self._source_mtimes and not self.dropdown_open():
                self._load_sources(auto=True)
                self._redraw()
        except OSError:
            pass
        self.after(3000, self._watch_sources)

    def _load_file(self):
        try:
            self.months = load_allocations()
            self._file_mtime = (ALLOCATION_FILE.stat().st_mtime
                                if ALLOCATION_FILE.exists() else None)
        except PermissionError:
            # Excel holds the file locked; saving would overwrite what we couldn't read.
            messagebox.showerror(
                APP_NAME, f"{ALLOCATION_FILE.name} is open in Excel. Close it and click Reload.")
            self.months = None

    def _reload(self):
        if self.dirty and not messagebox.askyesno(
                APP_NAME, "Reloading discards your unsaved changes. Continue?"):
            return
        self._load_sources()
        self._load_file()
        self.dirty = False
        self._redraw()

    def _month(self, key=None, create=False):
        key = key or self.month
        if self.months is None:
            return MonthData()
        if create:
            return self.months.setdefault(key, MonthData())
        return self.months.get(key) or MonthData()

    def _manager(self, resource, key=None):
        """A resource keeps the latest manager chosen in this or an earlier month."""
        key = key or self.month
        for k in sorted((k for k in self.months or {} if k <= key), reverse=True):
            if resource in self.months[k].managers:
                return self.months[k].managers[resource]
        return ""

    def _resources_for(self, key):
        """Master resources, plus anyone saved in this month who has since left the master."""
        names = set(self.resources) | set(self._month(key).gender) \
            | {r for r, _d in self._month(key).cells}
        return sorted(names, key=sort_key, reverse=self.sort_descending)

    def wwa_rows(self, dates, detailed=False):
        """Rows of the WWA layout for these dates: (resource, M/F, manager, {date: text})."""
        keys = _unique((d.year, d.month) for d in dates)
        names = sorted({n for k in keys for n in self._resources_for(k)}, key=sort_key,
                       reverse=self.sort_descending)
        rows = []
        for name in names:
            work = {}
            for d in dates:
                data = self._month((d.year, d.month))
                tasks = data.tasks.get((name, d))
                text = data.cells.get((name, d))
                if detailed and tasks:
                    text = "\n\n".join(task_detail(t) for t in tasks)
                if text:
                    work[d] = text
            gender = self.resources.get(name) or next(
                (self._month(k).gender[name] for k in keys if name in self._month(k).gender), "")
            rows.append((name, gender, self._manager(name, keys[0]), work))
        return rows

    # ---- drawing -------------------------------------------------------- #
    def _redraw(self):
        self.month_var.set(f"{calendar.month_name[self.month[1]]} {self.month[0]}")
        self.dates = month_dates(*self.month)
        self.row_names = self._resources_for(self.month)
        self._draw_corner()
        self._draw_top()
        self._draw_side()
        self._draw_body()
        self._update_summary()
        if self.weekly is not None and self.weekly.winfo_exists():
            self.weekly.render()

    def _cell(self, canvas, x, y, w, h, fill, text="", fg=TEXT, font=None, anchor="center",
              pad=0):
        canvas.create_rectangle(x, y, x + w, y + h, fill=fill, outline=BORDER)
        if text:
            tx = x + w / 2 if anchor == "center" else x + pad
            canvas.create_text(tx, y + h / 2, text=text, fill=fg, anchor=anchor,
                               justify="center" if anchor == "center" else "left",
                               width=w - self.S(10), font=font or (FONT, 9))

    def _draw_corner(self):
        c = self.corner
        c.delete("all")
        x = 0
        arrow = " ▼" if self.sort_descending else " ▲"
        for title, w in zip(LEFT_HEADERS, self.left_widths):
            text = title + arrow if title == "Resource" else title
            self._cell(c, x, 0, w, self.head_height, NAVY, text, "white", (FONT, 10, "bold"))
            x += w

    def _draw_top(self):
        c = self.top
        c.delete("all")
        today = date.today()
        for i, d in enumerate(self.dates):
            x = i * self.date_width
            fill = ACCENT if d == today else NAVY
            c.create_rectangle(x, 0, x + self.date_width, self.head_height, fill=fill,
                               outline=NAVY_2)
            c.create_text(x + self.date_width / 2, self.head_height * 0.3, text=d.strftime("%A"),
                          fill=ON_NAVY_MUTED if d != today else "white", font=(FONT, 9))
            c.create_text(x + self.date_width / 2, self.head_height * 0.68,
                          text=d.strftime("%d-%b-%y").lstrip("0"), fill="white",
                          font=(FONT, 10, "bold"))
        c.configure(scrollregion=(0, 0, len(self.dates) * self.date_width, self.head_height),
                    xscrollincrement=self.date_width)

    def _draw_side(self):
        c = self.side
        c.delete("all")
        data_h = len(self.row_names) * self.row_height
        for r, name in enumerate(self.row_names):
            y = r * self.row_height
            gender = self.resources.get(name) or self._month().gender.get(name, "")
            manager = self._manager(name)
            left_master = name not in self.resources
            values = ((str(r + 1), "center"), (name, "w"), (gender, "center"),
                      (manager or "Select", "center"))
            x = 0
            for col, ((text, anchor), w) in enumerate(zip(values, self.left_widths)):
                fg = TEXT
                if col == 3 and not manager:
                    fg = MUTED
                if col == 1 and left_master:
                    fg = WARN
                self._cell(c, x, y, w, self.row_height, "white", text, fg,
                           (FONT, 10, "bold") if col == 1 else (FONT, 10),
                           anchor=anchor, pad=self.S(10))
                x += w
        c.configure(scrollregion=(0, 0, sum(self.left_widths), data_h),
                    yscrollincrement=self.row_height)

    def _draw_body(self):
        c = self.body
        c.delete("all")
        today = date.today()
        data = self._month()
        cells = data.cells
        options = set(self.work_options)
        for r, name in enumerate(self.row_names):
            y = r * self.row_height
            for i, d in enumerate(self.dates):
                x = i * self.date_width
                work = cells.get((name, d))
                fill = "white"
                if is_off_day(d):
                    fill = OFF_DAY_BG
                elif d == today:
                    fill = TODAY_BG
                fg = TEXT
                if work:
                    fill = ACCENT_SOFT
                    tasks = data.tasks.get((name, d))
                    # Something was removed from the schedule since it was allocated
                    if (any(t not in self.schedule_keys for t in tasks) if tasks
                            else work not in options):
                        fill, fg = AMBER_SOFT, AMBER_DARK
                self._cell(c, x, y, self.date_width, self.row_height, fill,
                           self._display(work), fg, (FONT, 9))
        c.configure(scrollregion=(0, 0, len(self.dates) * self.date_width,
                                  len(self.row_names) * self.row_height),
                    xscrollincrement=self.date_width, yscrollincrement=self.row_height)

    @staticmethod
    def _display(work):
        """Client on one line, assignment on the next; further pairs are only counted."""
        lines = (work or "").split("\n")
        text = lines[0].replace(SEPARATOR, "\n", 1)
        return f"{text}\n+{len(lines) - 1} more" if len(lines) > 1 else text

    def _update_summary(self):
        allocated = len(self._month().cells)
        self.stat_vars["resources"].set(str(len(self.row_names)))
        self.stat_vars["work"].set(str(len(self.work_options)))
        self.stat_vars["allocated"].set(str(allocated))
        self.save_btn.state(["!disabled"] if self.dirty else ["disabled"])
        if self.months is None:
            self.status_var.set(f"{ALLOCATION_FILE.name} could not be read")
        elif self.dirty:
            self.status_var.set("Unsaved changes")
        elif not self.work_options:
            self.status_var.set("No work in Work Schedule.xlsx yet")
        else:
            self.status_var.set("All changes saved")

    # ---- interaction ---------------------------------------------------- #
    def dropdown_open(self):
        return self.dropdown is not None and self.dropdown.winfo_exists()

    def _open_dropdown(self, canvas, x, y, width, options, current, on_pick):
        """Open a list under the canvas position x, y (canvas coordinates)."""
        if self.months is None:
            return
        root_x = canvas.winfo_rootx() + int(x - canvas.canvasx(0))
        root_y = canvas.winfo_rooty() + int(y - canvas.canvasy(0))
        self.dropdown = CellDropdown(self, root_x, root_y, width, options, current, on_pick,
                                     self.S)

    def _cell_at(self, event):
        """(row, column) of a click on the grid, or None outside the cells."""
        x, y = self.body.canvasx(event.x), self.body.canvasy(event.y)
        col, row = int(x // self.date_width), int(y // self.row_height)
        if 0 <= row < len(self.row_names) and 0 <= col < len(self.dates):
            return row, col
        return None

    def _cancel_click(self):
        if self._click_job:
            self.after_cancel(self._click_job)
            self._click_job = None

    def _on_body_click(self, event):
        """Wait briefly, so that a double click doesn't show the work first."""
        cell = self._cell_at(event)
        self._cancel_click()
        if cell:
            self._click_job = self.after(DOUBLE_CLICK_MS, lambda: self._show_cell(*cell))

    def _on_body_double_click(self, event):
        self._cancel_click()
        cell = self._cell_at(event)
        if cell:
            self._open_allocate(*cell)

    def _show_cell(self, row, col):
        self._click_job = None
        name, d = self.row_names[row], self.dates[col]
        if (name, d) in self._month().cells:
            self._show_day(name, d)
        else:
            self.status_var.set("No work allocated on this date. Double-click to allocate.")

    def _short_names(self):
        return {full: short for short, full in self.managers}

    def _day_rows(self, resource, day):
        """(rows in the day details format, count of tasks no longer in the schedule)."""
        data = self._month((day.year, day.month))
        cell = (resource, day)
        short_names = self._short_names()
        tasks = data.tasks.get(cell)
        if tasks:
            lines = [e for e in self.schedule if task_key(e) in tasks]
        else:   # allocated before tasks were recorded: every line of that client and assignment
            lines = [e for e in self.schedule if work_label(e) == data.cells.get(cell)]
        rows = []
        for e in lines:
            manager = _clean(e.get("Manager"))
            rows.append(tuple(fmt_cell(v) if v is not None else "" for v in (
                e.get("Client Short Name"),
                short_names.get(manager) or e.get("Manager Short Name"),
                e.get("Assignment"), e.get("Work Type / Area"), e.get("Frequency"),
                e.get("Financial Year"), e.get("Period From"), e.get("Period To"),
                e.get("Target Date"), e.get("Due Date"), e.get("Client Full Name"),
                manager)))
        return rows, sum(t not in self.schedule_keys for t in tasks or ())

    def _show_day(self, resource, day, from_tab=False):
        work = self._month((day.year, day.month)).cells.get((resource, day))
        rows, removed = self._day_rows(resource, day) if work else ([], 0)
        note = None
        if not work:
            note = "No work allocated on this date."
        elif removed or not rows:
            note = f"Some allocated work is no longer in {SCHEDULE_FILE.name}."
        if self.details is not None and self.details.winfo_exists():
            self.details.destroy()
        self.details = DayDetails(
            self, resource, day, rows, note, self.S,
            on_change=self._pick_day if from_tab else None,
            on_clear=(lambda: self._open_clear_day(resource, day, from_tab)) if work else None)
        if from_tab:
            self._track_tab(self.details, DAY_TAB)

    def _open_allocate(self, row, col):
        self._click_job = None
        if self.months is None:
            return
        name, d = self.row_names[row], self.dates[col]
        has_work = (name, d) in self._month().cells
        AllocateWindow(self, name, d, self.schedule, self._month().tasks.get((name, d), []),
                       self._short_names(), lambda tasks: self._set_tasks(name, d, tasks),
                       self.S,
                       on_clear=(lambda: self._open_clear_day(name, d)) if has_work else None)

    def _on_side_click(self, event):
        x, y = self.side.canvasx(event.x), self.side.canvasy(event.y)
        row = int(y // self.row_height)
        manager_x = sum(self.left_widths[:3])
        if x < manager_x or not 0 <= row < len(self.row_names):
            return
        name = self.row_names[row]
        options = [(CLEAR_LABEL, "")] + [(f"{short}{SEPARATOR}{full}", short)
                                         for short, full in self.managers]
        self._open_dropdown(self.side, manager_x, (row + 1) * self.row_height, self.S(260),
                            options, self._manager(name),
                            lambda value: self._set_manager(name, value))

    def _on_corner_click(self, event):
        resource_x = self.left_widths[0]
        if resource_x <= event.x < resource_x + self.left_widths[1]:
            self.sort_descending = not self.sort_descending
            self._redraw()

    def _set_tasks(self, resource, d, tasks):
        data = self._month((d.year, d.month), create=True)
        if tasks:
            data.tasks[resource, d] = list(tasks)
            data.cells[resource, d] = cell_text(tasks)
        else:
            data.tasks.pop((resource, d), None)
            data.cells.pop((resource, d), None)
        self._mark_dirty()

    def _open_clear(self):
        if self.months is None:
            return
        data = self._month()
        items = []
        for name in self.row_names:
            for d in self.dates:
                work = data.cells.get((name, d))
                if work:
                    tasks = data.tasks.get((name, d)) or [None]
                    items += [(name, d, task, work) for task in tasks]
        if items:
            ClearWindow(self, self.month_var.get(), items, self._clear_items, self.S)
        else:
            messagebox.showinfo("Clear Allocations",
                                f"No work is allocated in {self.month_var.get()}.")

    def _clear_items(self, items):
        for resource, day, task, _work in items:
            tasks = self._month((day.year, day.month)).tasks.get((resource, day), [])
            self._set_tasks(resource, day, [t for t in tasks if t != task])

    def _open_clear_day(self, resource, day, from_tab=False):
        """Tick the tasks of one Resource and day to clear."""
        data = self._month((day.year, day.month))
        work = data.cells.get((resource, day))
        if not work:
            return
        items = [(resource, day, task, work)
                 for task in data.tasks.get((resource, day)) or [None]]

        def clear(chosen):
            self._clear_items(chosen)
            # Refresh the day's window; with nothing left, only the Day tab keeps it open
            if self.details is not None and self.details.winfo_exists():
                if from_tab or (resource, day) in data.cells:
                    self._show_day(resource, day, from_tab)
                else:
                    self.details.destroy()

        ClearWindow(self, f"{resource} - {day.strftime(DAY_TITLE_FMT)}", items, clear, self.S)

    def _set_manager(self, resource, value):
        self._month(create=True).managers[resource] = value
        self._mark_dirty()

    def _mark_dirty(self):
        self.dirty = True
        self._redraw()

    def _go_to(self, key):
        self.month = key
        self._redraw()
        self._xview("moveto", 0)
        self._yview("moveto", 0)

    def _go_to_today(self):
        today = date.today()
        self._go_to((today.year, today.month))
        self._scroll_to_today()

    def _scroll_to_today(self):
        today = date.today()
        if self.month == (today.year, today.month) and self.dates:
            self._xview("moveto", (today.day - 1) / len(self.dates))

    # ---- saving --------------------------------------------------------- #
    def _save(self):
        if not self.dirty or self.months is None:
            return True
        current = ALLOCATION_FILE.stat().st_mtime if ALLOCATION_FILE.exists() else None
        if current != self._file_mtime and not messagebox.askyesno(
                APP_NAME, f"{ALLOCATION_FILE.name} was changed outside this app.\n\n"
                          "Saving replaces those changes. Continue?"):
            return False

        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for key in sorted(self.months):
            if self.months[key].cells or self.months[key].managers:
                write_wwa(wb.create_sheet(sheet_name(key)), month_title(key), month_dates(*key),
                          self.wwa_rows(month_dates(*key)))
        if not wb.sheetnames:
            wb.create_sheet("Allocation")
        write_tasks(wb.create_sheet(TASKS_SHEET), self.months)
        temp = ALLOCATION_FILE.with_name(ALLOCATION_FILE.stem + ".saving.xlsx")
        try:
            wb.save(temp)
            os.replace(temp, ALLOCATION_FILE)
        except PermissionError:
            temp.unlink(missing_ok=True)
            messagebox.showerror(
                APP_NAME, f"{ALLOCATION_FILE.name} is open in Excel.\n\n"
                          "Close it and click Save again. Nothing has been lost.")
            return False
        self._file_mtime = ALLOCATION_FILE.stat().st_mtime
        self.dirty = False
        self._update_summary()
        return True

    def _export(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="Export to Excel", defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialdir=str(BASE_DIR),
            initialfile=f"Work Allocation Export {sheet_name(self.month)}.xlsx")
        if not path:
            return
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name(self.month)
        dates = month_dates(*self.month)
        write_wwa(ws, month_title(self.month), dates, self.wwa_rows(dates))
        try:
            wb.save(path)
        except PermissionError:
            messagebox.showerror(APP_NAME, f"{Path(path).name} is open. Close it and try again.")
            return
        self.status_var.set(f"Exported to {Path(path).name}")

    def save_as(self, parent, stem, kind, writer):
        """Ask where to save an Excel or PDF export, then write it."""
        label = "PDF file" if kind == "pdf" else "Excel workbook"
        path = filedialog.asksaveasfilename(
            parent=parent, title=f"Export to {'PDF' if kind == 'pdf' else 'Excel'}",
            defaultextension=f".{kind}", filetypes=[(label, f"*.{kind}")],
            initialdir=str(BASE_DIR), initialfile=f"{stem}.{kind}")
        if not path:
            return
        try:
            writer(path)
        except PermissionError:
            messagebox.showerror(APP_NAME, f"{Path(path).name} is open. Close it and try again.",
                                 parent=parent)
        except ImportError:
            messagebox.showerror(APP_NAME, "PDF export needs the reportlab package:\n\n"
                                           "pip install reportlab", parent=parent)
        else:
            messagebox.showinfo(APP_NAME, f"Exported to {Path(path).name}", parent=parent)

    def _on_close(self):
        if self.dirty:
            answer = messagebox.askyesnocancel(APP_NAME, "Save your changes before closing?")
            if answer is None or (answer and not self._save()):
                return
        self.destroy()


if __name__ == "__main__":
    if sys.platform == "win32":
        try:  # crisp text on high-DPI screens
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    AllocatorApp().mainloop()
