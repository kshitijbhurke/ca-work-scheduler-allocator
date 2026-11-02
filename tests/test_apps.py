"""
End-to-end checks of both apps, run against a temporary copy of the sample workbooks.

    python -m unittest discover -s tests -v

The real windows are built and driven through their own handlers, with pop-ups answered
automatically, so nothing in the project folder is changed.
"""

import importlib.util
import shutil
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
FY_2026 = "FY 2026-27"
CLIENT = "Cedar Foods Private Limited"
ASSIGNMENT = "Accounting"
WORK_TYPE = "Filing of TDS Return in Form 26Q / 24Q"


def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


scheduler = load("Work_Scheduler")
allocator = load("Work_Allocator")


class AppTest(unittest.TestCase):
    """Points both apps at a temporary folder and answers pop-ups with Yes / OK."""

    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.dir = Path(temp.name)
        for name in ("Work Schedule Master.xlsx", "Work Schedule.xlsx"):
            shutil.copy(ROOT / name, self.dir / name)

        for module in (scheduler, allocator):
            for attr, target in (("BASE_DIR", self.dir),
                                 ("MASTER_FILE", self.dir / "Work Schedule Master.xlsx"),
                                 ("SCHEDULE_FILE", self.dir / "Work Schedule.xlsx")):
                self._patch(module, attr, target)
            for attr in ("showinfo", "showwarning", "showerror"):
                self._patch(module.messagebox, attr, mock.Mock())
            for attr in ("askyesno", "askyesnocancel"):
                self._patch(module.messagebox, attr, mock.Mock(return_value=True))
        self._patch(allocator, "ALLOCATION_FILE", self.dir / "Work Allocation.xlsx")
        self._patch(scheduler.os, "startfile", mock.Mock(), create=True)

    def _patch(self, target, attr, value, **kwargs):
        patcher = mock.patch.object(target, attr, value, **kwargs)
        patcher.start()
        self.addCleanup(patcher.stop)

    def open(self, app_class):
        app = app_class()
        self.addCleanup(app.destroy)
        app.update()
        return app

    def fill_form(self, app):
        app.client_var.set(CLIENT)
        app._on_client()
        app.assignment_var.set(ASSIGNMENT)
        app._on_assignment()
        app.work_type_var.set(WORK_TYPE)
        app._on_work_type()
        app.frequency_var.set("Quarterly")
        app._on_frequency()
        app.fy_var.set(FY_2026)
        app._sync()


class PeriodRules(unittest.TestCase):
    def test_line_items_per_frequency(self):
        for months, count in ((1, 12), (3, 4), (6, 2), (12, 1)):
            periods = scheduler.fy_periods(2026, months)
            self.assertEqual(len(periods), count)
            self.assertEqual(periods[0][0], date(2026, 4, 1))
            self.assertEqual(periods[-1][1], date(2027, 3, 31))

    def test_frequency_kinds(self):
        self.assertEqual(scheduler.frequency_kind("Half-Yearly"), ("fy", 6))
        self.assertEqual(scheduler.frequency_kind("NA"), ("na", None))
        self.assertEqual(scheduler.frequency_kind("Annual Dec End"), ("manual", None))

    def test_financial_year_turns_in_april(self):
        self.assertEqual(scheduler.current_fy_start(date(2027, 3, 31)), 2026)
        self.assertEqual(scheduler.current_fy_start(date(2027, 4, 1)), 2027)

    def test_off_days(self):
        self.assertTrue(allocator.is_off_day(date(2026, 11, 1)))    # Sunday
        self.assertTrue(allocator.is_off_day(date(2026, 11, 7)))    # 1st Saturday
        self.assertFalse(allocator.is_off_day(date(2026, 11, 14)))  # 2nd Saturday
        self.assertTrue(allocator.is_off_day(date(2026, 11, 21)))   # 3rd Saturday
        self.assertFalse(allocator.is_off_day(date(2026, 11, 2)))   # Monday


class SchedulerFlow(AppTest):
    def test_sample_data_loads(self):
        app = self.open(scheduler.SchedulerApp)
        self.assertEqual(len(app.entries), 29)
        self.assertGreater(len(app.master_data.clients), 0)

    def test_add_submit_and_reopen(self):
        app = self.open(scheduler.SchedulerApp)
        self.fill_form(app)
        app._add_entry()
        self.assertEqual(len(app.pending), 4)
        self.assertEqual(len(scheduler.load_schedule()), 29)    # nothing saved before Submit
        app._submit_pending()
        self.assertEqual(app.pending, [])
        self.assertEqual(len(scheduler.load_schedule()), 33)

    def test_duplicates_are_skipped(self):
        app = self.open(scheduler.SchedulerApp)
        self.fill_form(app)
        app._add_entry()
        app._submit_pending()
        self.fill_form(app)
        app._add_entry()
        self.assertEqual(app.pending, [])
        scheduler.messagebox.showwarning.assert_called()

    def test_bulk_update_and_delete(self):
        app = self.open(scheduler.SchedulerApp)
        rows = app.tree.get_children()[3:5]
        app.tree.selection_set(rows)
        app.update()
        app.status_edit_var.set("WIP")
        app._apply_updates()
        saved = scheduler.load_schedule()
        self.assertEqual([e["Work Status"] for e in saved[3:5]], ["WIP", "WIP"])
        app.tree.selection_set(app.tree.get_children()[0])
        app.update()
        app._delete_selected()
        self.assertEqual(len(scheduler.load_schedule()), 28)

    def test_export_chosen_fields(self):
        app = self.open(scheduler.SchedulerApp)
        path = self.dir / "export.xlsx"
        with mock.patch.object(scheduler.filedialog, "asksaveasfilename",
                               return_value=str(path)):
            app._export_excel(["Client Short Name", "Work Status"])
        ws = openpyxl.load_workbook(path).active
        self.assertEqual([c.value for c in ws[1]], ["Client Short Name", "Work Status"])
        self.assertEqual(ws.max_row, 30)


class AllocatorFlow(AppTest):
    def allocate(self, app, day):
        """Allocate the first pending task of the schedule to the first Resource."""
        task = next(e for e in app.schedule if not allocator.is_complete(e))
        name = sorted(app.resources, key=allocator.sort_key)[0]
        app._set_tasks(name, day, [allocator.task_key(task)])
        return name, allocator.task_key(task)

    def test_sample_data_loads(self):
        app = self.open(allocator.AllocatorApp)
        self.assertEqual(len(app.schedule), 29)
        self.assertGreater(len(app.resources), 0)
        self.assertEqual(len(app.dates), len(allocator.month_dates(*app.month)))

    def test_allocate_save_and_reload(self):
        app = self.open(allocator.AllocatorApp)
        day = date(*app.month, 10)
        name, key = self.allocate(app, day)
        self.assertTrue(app.dirty)
        self.assertFalse(allocator.ALLOCATION_FILE.exists())    # nothing saved before Save
        self.assertTrue(app._save())
        months = allocator.load_allocations()
        self.assertEqual(months[app.month].tasks[name, day], [key])
        hidden = openpyxl.load_workbook(allocator.ALLOCATION_FILE)[allocator.TASKS_SHEET]
        self.assertEqual(hidden.sheet_state, "hidden")

    def test_clear_allocation(self):
        app = self.open(allocator.AllocatorApp)
        day = date(*app.month, 10)
        name, key = self.allocate(app, day)
        app._clear_items([(name, day, key, "")])
        self.assertNotIn((name, day), app._month().cells)

    def test_complete_work_is_not_offered(self):
        app = self.open(allocator.AllocatorApp)
        window = allocator.AllocateWindow(app, "Someone", date.today(), app.schedule, [], {},
                                          lambda _tasks: None, app.S)
        complete = [e for e in app.schedule if allocator.is_complete(e)]
        self.assertTrue(complete)
        first = complete[0]
        for combo, value in zip(window.combos, (first["Client Short Name"],
                                                allocator._clean(first["Assignment"]),
                                                allocator.financial_year(first))):
            combo.set(value)
        window._render()
        self.assertNotIn(allocator.task_key(first), window.shown)
        window.destroy()

    def test_day_and_week_views_and_exports(self):
        app = self.open(allocator.AllocatorApp)
        day = date(*app.month, 10)
        name, _key = self.allocate(app, day)
        rows, removed = app._day_rows(name, day)
        self.assertEqual((len(rows), removed), (1, 0))
        dates = allocator.week_of(day)
        week = app.wwa_rows(dates, detailed=True)

        allocator.write_day(self.dir / "day.xlsx", name, day, rows)
        allocator.day_pdf(str(self.dir / "day.pdf"), name, day, rows)
        allocator.write_week(self.dir / "week.xlsx", dates, week, True)
        allocator.week_pdf(str(self.dir / "week.pdf"), dates, week)
        for file in ("day.xlsx", "day.pdf", "week.xlsx", "week.pdf"):
            self.assertGreater((self.dir / file).stat().st_size, 0, file)
        for file in ("day.pdf", "week.pdf"):
            self.assertEqual((self.dir / file).read_bytes()[:4], b"%PDF")

        weekly = allocator.WeeklySchedule(app, app.S)
        weekly._toggle_detail()
        app.update()
        weekly.destroy()


if __name__ == "__main__":
    unittest.main()
