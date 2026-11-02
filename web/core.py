"""
Data rules shared by the two web apps. No Streamlit or Tkinter here.

Reads and writes the same Excel files, in the same layouts, as the desktop apps, so a firm can
use either version on the same data.
"""

import calendar
import io
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

BASE_DIR = Path(__file__).resolve().parent.parent
MASTER_FILE = BASE_DIR / "Work Schedule Master.xlsx"
SCHEDULE_FILE = BASE_DIR / "Work Schedule.xlsx"
ALLOCATION_FILE = BASE_DIR / "Work Allocation.xlsx"
SCHEDULE_SHEET = "Scheduler"
RESOURCE_SHEET = "Resource Master"
MANAGER_SHEET = "Manager Master"
TASKS_SHEET = "Allocated Tasks"

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
    "Client Group": 24, "Assignment": 28, "Work Type / Area": 48, "Frequency": 18,
    "Financial Year": 14, "Period From": 14, "Period To": 14, "Manager": 22,
    "Manager Short Name": 16, "Due Date": 14, "Target Date": 14, "Work Status": 14,
}
DATE_FMT = "%d-%b-%Y"            # 01-Apr-2026
OLD_DATE_FMTS = ("%d-%m-%Y",)    # accepted when reading older files
DAY_TITLE_FMT = "%b %d, %Y [%A]"   # Nov 03, 2026 [Tuesday]
NA = "NA"
SEPARATOR = " - "
COMPLETE = "complete"
DUPLICATE_KEYS = ("Client Full Name", "Assignment", "Work Type / Area", "Frequency",
                  "Period From", "Period To")

# Frequencies that are split across a Financial Year (April to March): months per line item.
FY_FREQUENCIES = {
    "monthly": 1,
    "quarterly": 3,
    "half yearly": 6, "halfyearly": 6, "six monthly": 6, "semi annual": 6,
    "annual": 12, "yearly": 12,
}
NA_FREQUENCIES = {"na", "n/a", "n.a.", "not applicable"}

# These fields identify a schedule line in the hidden sheet of 'Work Allocation.xlsx'.
TASK_FIELDS = ("Client Short Name", "Assignment", "Work Type / Area", "Frequency",
               "Financial Year", "Period From", "Period To")

# Layout of the 'WWA Format' sheet
TITLE_ROW, DAY_ROW, HEADER_ROW, FIRST_DATA_ROW = 5, 6, 7, 8
FIRST_DATE_COL = 5                 # column E
LEFT_HEADERS = ("Sr\nNo", "Resource", "M/F", "Manager")
LEFT_WIDTHS = (4.6, 29.6, 6.1, 9.2)
WORKDAY_WIDTH, OFF_DAY_WIDTH = 19.1, 13.7
DETAIL_WIDTH = 36                  # day columns in a detailed weekly export
XL_NAVY, XL_OFF_DAY = "002060", "F6C6AD"
XL_FONT = "Book Antiqua"

# Layout of the 'Resource Day Details Format' sheet
DAY_COLUMNS = ("Client Short Name", "Manager", "Assignment", "Work Type / Area", "Frequency",
               "FY", "Period From", "Period To", "Target Date", "Due Date",
               "Client Full Name", "Manager Name")
DAY_XL_WIDTHS = (20.9, 10.8, 18.9, 25.2, 17.6, 14.0, 13.8, 17.2, 16.2, 17.0, 29.8, 36.8)
DAY_PDF_WIDTHS = (56, 42, 62, 100, 50, 46, 52, 52, 52, 52, 100, 70)
DAY_ANCHORS = "lclllccccc" + "ll"


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def clean(value):
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


def unique(seq):
    """De-duplicate while keeping the order used in the master sheet."""
    seen, out = set(), []
    for item in seq:
        if item and item not in seen:
            seen.add(item)
            out.append(item)
    return out


def sort_key(name):
    return name.casefold()


def mtime(path):
    return path.stat().st_mtime if path.exists() else None


def frequency_kind(freq):
    """("fy", months per item) for FY frequencies, ("na", None), or ("manual", None)."""
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
    """Split FY start_year (1 Apr start_year to 31 Mar next year) into equal periods."""
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


def month_title(key):
    return f"Weekly Work Allocation - {calendar.month_name[key[1]]} {key[0]}"


def week_of(day):
    """Monday to Sunday of the week containing day."""
    monday = day - timedelta(days=day.weekday())
    return [monday + timedelta(days=i) for i in range(7)]


def week_title(dates):
    return (f"Weekly Work Allocation - {dates[0].strftime(DATE_FMT)} "
            f"to {dates[-1].strftime(DATE_FMT)}")


def day_title(day):
    return f"Schedule for the day - {day.strftime(DAY_TITLE_FMT)}"


def replace_file(path, write):
    """Write to a temporary file, then swap it in, so a failed save leaves the file intact."""
    temp = path.with_name(path.stem + ".saving.xlsx")
    try:
        write(temp)
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def workbook_bytes(write):
    buffer = io.BytesIO()
    write(buffer)
    return buffer.getvalue()


# --------------------------------------------------------------------------- #
# Master data
# --------------------------------------------------------------------------- #
def _open_shared(path):
    """Open a workbook read-only, even while it is open (locked) in Excel."""
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        if sys.platform != "win32":
            raise
        # Excel holds the file locked; Windows' CopyFile can still copy it.
        import ctypes
        import tempfile
        tmp = Path(tempfile.gettempdir()) / f"work_scheduler_web_{path.stem}.xlsx"
        if not ctypes.windll.kernel32.CopyFileW(str(path), str(tmp), False):
            raise
        return openpyxl.load_workbook(tmp, read_only=True, data_only=True)


def _read_table(ws, required_headers):
    """Find the header row anywhere on the sheet and return the rows as dicts."""
    rows = list(ws.iter_rows(values_only=True))
    for idx, row in enumerate(rows):
        cleaned = [clean(v) for v in row]
        if all(h in cleaned for h in required_headers):
            col = {h: cleaned.index(h) for h in required_headers}
            out = []
            for data_row in rows[idx + 1:]:
                record = {h: clean(data_row[c]) if c < len(data_row) else None
                          for h, c in col.items()}
                if any(record.values()):
                    out.append(record)
            return out
    raise ValueError(
        f"Sheet '{ws.title}' has no header row with: {', '.join(required_headers)}")


def _read_managers(ws):
    """Return (manager names, {name: short name}); the short name column is optional."""
    for name_col in ("Manager Name", "Manager"):
        for short_col in ("Mgr Short Name", "Manager Short Name", None):
            headers = [name_col] + ([short_col] if short_col else [])
            try:
                rows = _read_table(ws, headers)
            except ValueError:
                continue
            return (unique(r[name_col] for r in rows),
                    {r[name_col]: r.get(short_col) for r in rows if r[name_col]})
    return [], {}


class Masters:
    """Everything read from 'Work Schedule Master.xlsx'."""

    def __init__(self, path=None):
        wb = _open_shared(path or MASTER_FILE)
        try:
            client_headers = ["Client Full Name", "Client Short Name"]
            try:
                clients = _read_table(wb["Client List"], [*client_headers, "Client Group"])
            except ValueError:   # older workbook without the 'Client Group' column
                clients = _read_table(wb["Client List"], client_headers)
            assignments = _read_table(wb["Assignment Master"], ["Assignment"])
            work_types = _read_table(wb["Work Type Master"], ["Assignment", "Work Type / Area"])
            frequency_sheet = ("Frequency Master" if "Frequency Master" in wb.sheetnames
                               else "Work Type Master")
            frequencies = _read_table(wb[frequency_sheet],
                                      ["Assignment", "Work Type / Area", "Frequency"])
            self.managers, self.manager_short = [], {}
            sheet = next((n for n in (MANAGER_SHEET, "Manager List") if n in wb.sheetnames), None)
            if sheet:
                self.managers, self.manager_short = _read_managers(wb[sheet])
            rows = (_read_table(wb[RESOURCE_SHEET], ["Resource Full Name", "M/F"])
                    if RESOURCE_SHEET in wb.sheetnames else [])
        finally:
            wb.close()

        self.clients, self.client_groups = {}, {}
        for r in clients:
            if r["Client Full Name"]:
                self.clients.setdefault(r["Client Full Name"], r["Client Short Name"] or "")
                self.client_groups.setdefault(r["Client Full Name"], r.get("Client Group"))
        self.assignments = unique(r["Assignment"] for r in assignments)
        self.work_type_rows = [r for r in work_types if r["Assignment"] and r["Work Type / Area"]]
        self.frequency_rows = [r for r in frequencies
                               if r["Assignment"] and r["Work Type / Area"] and r["Frequency"]]
        self.resources = {r["Resource Full Name"]: r["M/F"] or "" for r in rows
                          if r["Resource Full Name"]}

    def work_types_for(self, assignment):
        return unique(r["Work Type / Area"] for r in self.work_type_rows
                      if r["Assignment"] == assignment)

    def frequencies_for(self, assignment, work_type):
        return unique(r["Frequency"] for r in self.frequency_rows
                      if r["Assignment"] == assignment and r["Work Type / Area"] == work_type)

    def short_names(self):
        """[(manager short name, full name)]"""
        return [(self.manager_short.get(n) or n, n) for n in self.managers]


# --------------------------------------------------------------------------- #
# Schedule
# --------------------------------------------------------------------------- #
def load_schedule(path=None):
    path = path or SCHEDULE_FILE
    if not path.exists():
        return []
    wb = _open_shared(path)
    try:
        ws = wb[SCHEDULE_SHEET] if SCHEDULE_SHEET in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        return []
    headers = [clean(h) for h in rows[0]]
    entries = []
    for row in rows[1:]:
        if any(v is not None for v in row):
            e = {h: v for h, v in zip(headers, row) if h}
            for c in DATE_COLUMNS:
                v = e.get(c)
                e[c] = NA if isinstance(v, str) and v.strip().upper() == NA else to_date(v)
            status = clean(e.get("Work Status"))
            e["Work Status"] = next((s for s in WORK_STATUSES
                                     if status and s.lower() == status.lower()), None)
            entries.append(e)
    return entries


def write_schedule(target, entries, columns=SCHEDULE_COLUMNS):
    """Write entries to an Excel file or stream, with the given columns in the given order."""
    columns = list(columns)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SCHEDULE_SHEET
    ws.append(columns)
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
        status_col = get_column_letter(columns.index("Work Status") + 1)
        dv = DataValidation(type="list", formula1=f'"{",".join(WORK_STATUSES)}"',
                            allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{status_col}2:{status_col}{max(len(entries) + 1, 2) + 500}")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(target)


def save_schedule(entries, path=None):
    for i, e in enumerate(entries, start=1):
        e["Sr No"] = i
    replace_file(path or SCHEDULE_FILE, lambda temp: write_schedule(temp, entries))


def sync_from_masters(entries, masters):
    """Bring rows' Client Group and manager short name in line with the masters."""
    changed = False
    for e in entries:
        client = e.get("Client Full Name")
        if client in masters.client_groups and e.get("Client Group") != masters.client_groups[client]:
            e["Client Group"] = masters.client_groups[client]
            changed = True
        name = e.get("Manager")
        if name in masters.manager_short and e.get("Manager Short Name") != masters.manager_short[name]:
            e["Manager Short Name"] = masters.manager_short[name]
            changed = True
    return changed


def build_line_items(masters, client, assignment, work_type, frequency, fy=None, start=None,
                     end=None, added=None):
    """Turn one form entry into its line items."""
    base = {
        "Date Added": (added or datetime.now()).strftime("%d-%m-%Y %H:%M"),
        "Client Full Name": client,
        "Client Short Name": masters.clients.get(client, ""),
        "Client Group": masters.client_groups.get(client),
        "Assignment": assignment,
        "Work Type / Area": work_type,
        "Frequency": frequency,
        "Financial Year": "",
    }
    mode, months = frequency_kind(frequency)
    if mode == "fy":
        return [dict(base, **{"Financial Year": fy, "Period From": a, "Period To": b})
                for a, b in fy_periods(fy_start_year(fy), months)]
    if mode == "na":
        return [dict(base, **{"Period From": NA, "Period To": NA})]
    return [dict(base, **{"Period From": start, "Period To": end})]


def is_duplicate(entry, rows):
    return any(all(e.get(k) == entry[k] for k in DUPLICATE_KEYS) for e in rows)


def split_duplicates(items, *existing):
    """(new items, duplicates), checked against each list and against earlier items."""
    new, dups = [], []
    for item in items:
        (dups if any(is_duplicate(item, rows) for rows in (*existing, new)) else new).append(item)
    return new, dups


def describe(e):
    return (f"{e.get('Client Short Name')} | {e.get('Assignment')} | "
            f"{e.get('Work Type / Area')} | {e.get('Frequency')} | "
            f"{fmt_cell(e.get('Period From'))} to {fmt_cell(e.get('Period To'))}")


def search(entries, query):
    """Rows holding every word of the query, as [(index in entries, row)]."""
    words = query.lower().split()
    found = []
    for i, e in enumerate(entries):
        text = " ".join(str(fmt_cell(e.get(c))) for c in SCHEDULE_COLUMNS).lower()
        if all(w in text for w in words):
            found.append((i, e))
    return found


def is_overdue(entry, today=None):
    due = entry.get("Due Date")
    return (isinstance(due, date) and due < (today or date.today())
            and entry.get("Work Status") != "Complete")


# --------------------------------------------------------------------------- #
# Allocations
# --------------------------------------------------------------------------- #
class MonthData:
    """One month of allocations, keyed by Resource Full Name."""

    def __init__(self):
        self.cells = {}      # (resource, date) -> text shown in the WWA cell
        self.tasks = {}      # (resource, date) -> [task_key, ...]
        self.managers = {}   # resource -> manager short name ("" = none)
        self.gender = {}     # resource -> M/F as saved


def work_label(entry):
    client, assignment = clean(entry.get("Client Short Name")), clean(entry.get("Assignment"))
    return f"{client}{SEPARATOR}{assignment}" if client and assignment else None


def task_key(entry):
    return tuple(clean(fmt_cell(entry.get(f))) or "" for f in TASK_FIELDS)


def task_label(key):
    return f"{key[0]}{SEPARATOR}{key[1]}"


def cell_text(tasks):
    return "\n".join(unique(task_label(t) for t in tasks))


def task_period(task):
    return task[5] if task[5] == task[6] else f"{task[5]} to {task[6]}"


def task_detail(task):
    return "\n".join(filter(None, (task_label(task), task[2], task_period(task))))


def financial_year(entry):
    return clean(entry.get("Financial Year")) or NA


def is_complete(entry):
    return (clean(entry.get("Work Status")) or "").casefold() == COMPLETE


def load_allocations(path=None):
    """Read every month sheet of 'Work Allocation.xlsx' into {(year, month): MonthData}."""
    path = path or ALLOCATION_FILE
    if not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
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
                resource = clean(row[1])
                if not resource:
                    continue
                data.gender[resource] = clean(row[2]) or ""
                data.managers[resource] = clean(row[3]) or ""
                for i, d in dates.items():
                    work = clean(row[i]) if i < len(row) else None
                    if work:
                        data.cells[resource, d] = work
            months[first.year, first.month] = data
        if TASKS_SHEET in wb.sheetnames:
            for row in wb[TASKS_SHEET].iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue
                resource, day = clean(row[0]), date.fromisoformat(str(row[1]))
                data = months.get((day.year, day.month))
                if data and (resource, day) in data.cells:
                    key = tuple(clean(v) or "" for v in row[2:2 + len(TASK_FIELDS)])
                    data.tasks.setdefault((resource, day), []).append(key)
        for data in months.values():
            for cell, tasks in data.tasks.items():
                data.cells[cell] = cell_text(tasks)
    finally:
        wb.close()
    return months


class Book:
    """All saved months plus the lists they are shown against."""

    def __init__(self, months, resources, descending=False):
        self.months, self.resources, self.descending = months, resources, descending

    def month(self, key, create=False):
        if create:
            return self.months.setdefault(key, MonthData())
        return self.months.get(key) or MonthData()

    def manager(self, resource, key):
        """A resource keeps the latest manager chosen in this or an earlier month."""
        for k in sorted((k for k in self.months if k <= key), reverse=True):
            if resource in self.months[k].managers:
                return self.months[k].managers[resource]
        return ""

    def resources_for(self, key):
        """Master resources, plus anyone saved in this month who has since left the master."""
        data = self.month(key)
        names = set(self.resources) | set(data.gender) | {r for r, _d in data.cells}
        return sorted(names, key=sort_key, reverse=self.descending)

    def set_tasks(self, resource, day, tasks):
        data = self.month((day.year, day.month), create=True)
        if tasks:
            data.tasks[resource, day] = list(tasks)
            data.cells[resource, day] = cell_text(tasks)
        else:
            data.tasks.pop((resource, day), None)
            data.cells.pop((resource, day), None)

    def set_manager(self, resource, key, value):
        self.month(key, create=True).managers[resource] = value

    def wwa_rows(self, dates, detailed=False):
        """Rows of the WWA layout for these dates: (resource, M/F, manager, {date: text})."""
        keys = unique((d.year, d.month) for d in dates)
        names = sorted({n for k in keys for n in self.resources_for(k)}, key=sort_key,
                       reverse=self.descending)
        rows = []
        for name in names:
            work = {}
            for d in dates:
                data = self.month((d.year, d.month))
                tasks = data.tasks.get((name, d))
                text = data.cells.get((name, d))
                if detailed and tasks:
                    text = "\n\n".join(task_detail(t) for t in tasks)
                if text:
                    work[d] = text
            gender = self.resources.get(name) or next(
                (self.month(k).gender[name] for k in keys if name in self.month(k).gender), "")
            rows.append((name, gender, self.manager(name, keys[0]), work))
        return rows

    def day_rows(self, resource, day, schedule, short_names):
        """(rows in the day details format, count of tasks no longer in the schedule)."""
        data = self.month((day.year, day.month))
        cell = (resource, day)
        tasks = data.tasks.get(cell)
        if tasks:
            lines = [e for e in schedule if task_key(e) in tasks]
        else:   # allocated before tasks were recorded: every line of that client and assignment
            lines = [e for e in schedule if work_label(e) == data.cells.get(cell)]
        rows = []
        for e in lines:
            manager = clean(e.get("Manager"))
            rows.append(tuple(fmt_cell(v) if v is not None else "" for v in (
                e.get("Client Short Name"),
                short_names.get(manager) or e.get("Manager Short Name"),
                e.get("Assignment"), e.get("Work Type / Area"), e.get("Frequency"),
                e.get("Financial Year"), e.get("Period From"), e.get("Period To"),
                e.get("Target Date"), e.get("Due Date"), e.get("Client Full Name"),
                manager)))
        keys = {task_key(e) for e in schedule}
        return rows, sum(t not in keys for t in tasks or ())

    def save(self, path=None):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for key in sorted(self.months):
            if self.months[key].cells or self.months[key].managers:
                write_wwa(wb.create_sheet(sheet_name(key)), month_title(key), month_dates(*key),
                          self.wwa_rows(month_dates(*key)))
        if not wb.sheetnames:
            wb.create_sheet("Allocation")
        write_tasks(wb.create_sheet(TASKS_SHEET), self.months)
        replace_file(path or ALLOCATION_FILE, wb.save)


def write_tasks(ws, months):
    """Fill the hidden sheet that links each cell to its schedule lines."""
    ws.append(["Resource", "Date", *TASK_FIELDS])
    for key in sorted(months):
        for (resource, day), tasks in sorted(months[key].tasks.items(), key=lambda i: i[0]):
            for task in tasks:
                ws.append([resource, day.isoformat(), *task])
    ws.sheet_state = "hidden"


def write_wwa(ws, title, dates, rows, day_width=None):
    """Fill a sheet in the WWA layout. rows = [(resource, M/F, manager, {date: work})]."""
    side = Side(style="thin")
    box = Border(left=side, right=side, top=side, bottom=side)
    centre = Alignment(horizontal="center", vertical="center", wrap_text=True)
    navy = PatternFill("solid", fgColor=XL_NAVY)
    off_day = PatternFill("solid", fgColor=XL_OFF_DAY)
    font = Font(name=XL_FONT, size=11)

    ws.cell(TITLE_ROW, 1, title).font = Font(name=XL_FONT, size=11, bold=True)
    for col, (heading, width) in enumerate(zip(LEFT_HEADERS, LEFT_WIDTHS), start=1):
        cell = ws.cell(HEADER_ROW, col, heading)
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
        width = day_width or WORKDAY_WIDTH
        lines = max((sum(1 + len(line) // int(width) for line in text.split("\n"))
                     for text in work.values()), default=1)
        if lines > 1:
            ws.row_dimensions[row].height = 15.5 * lines
    ws.freeze_panes = ws.cell(FIRST_DATA_ROW, FIRST_DATE_COL)


def _landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)


def month_xlsx(book, key):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name(key)
    dates = month_dates(*key)
    write_wwa(ws, month_title(key), dates, book.wwa_rows(dates))
    return workbook_bytes(wb.save)


def week_xlsx(dates, rows, detailed):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Schedule"
    write_wwa(ws, week_title(dates), dates, rows, DETAIL_WIDTH if detailed else None)
    _landscape(ws)
    return workbook_bytes(wb.save)


def day_xlsx(resource, day, rows):
    """A day's tasks in the 'Resource Day Details Format' layout."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resource Day Details"
    side = Side(style="thin")
    box = Border(left=side, right=side, top=side, bottom=side)
    bold = Font(name=XL_FONT, size=11, bold=True)
    ws["A1"], ws["A2"], ws["B2"] = "Resource Day Details", resource, day_title(day)
    ws["A1"].font = ws["A2"].font = ws["B2"].font = bold
    for col, (title, width, anchor) in enumerate(zip(DAY_COLUMNS, DAY_XL_WIDTHS, DAY_ANCHORS),
                                                 start=1):
        cell = ws.cell(3, col, title)
        cell.font = Font(name=XL_FONT, size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=XL_NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = box
        ws.column_dimensions[get_column_letter(col)].width = width
        for r, row in enumerate(rows, start=4):
            cell = ws.cell(r, col, row[col - 1] or None)
            cell.font, cell.border = Font(name=XL_FONT, size=11), box
            cell.alignment = Alignment(horizontal="left" if anchor == "l" else "center",
                                       vertical="center", wrap_text=True)
    _landscape(ws)
    return workbook_bytes(wb.save)


def _pdf(title, subtitle, table, widths, off_columns=(), font_size=8):
    """A landscape A4 table as PDF bytes. table[0] is the heading row."""
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
    buffer = io.BytesIO()
    SimpleDocTemplate(buffer, pagesize=landscape(A4), title=title, leftMargin=margin,
                      rightMargin=margin, topMargin=margin, bottomMargin=margin).build(
        [para(title, heading), para(subtitle, heading), Spacer(1, 8), grid])
    return buffer.getvalue()


def day_pdf(resource, day, rows):
    return _pdf(resource, day_title(day), [list(DAY_COLUMNS), *rows], DAY_PDF_WIDTHS, font_size=7)


def week_pdf(dates, rows):
    headings = [h.replace("\n", " ") for h in LEFT_HEADERS]
    table = [headings + [d.strftime("%A\n%d-%b-%y") for d in dates]]
    for sr, (resource, gender, manager, work) in enumerate(rows, start=1):
        table.append([sr, resource, gender, manager, *(work.get(d) for d in dates)])
    off = [len(headings) + i for i, d in enumerate(dates) if is_off_day(d)]
    return _pdf("Weekly Work Allocation",
                f"{dates[0].strftime(DATE_FMT)} to {dates[-1].strftime(DATE_FMT)}",
                table, (24, 92, 30, 42, *[88] * len(dates)), off, font_size=7)
