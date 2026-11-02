"""
Checks of the two web apps and of their compatibility with the desktop file formats.

    python -m unittest discover -s tests -v

Runs against a temporary copy of the sample workbooks. The apps run headless through
Streamlit's own test harness.
"""

import shutil
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import openpyxl
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parent.parent
WEB = ROOT / "web"
sys.path.insert(0, str(WEB))

import core  # noqa: E402

CLIENT = "Cedar Foods Private Limited"
ASSIGNMENT = "Accounting"
WORK_TYPE = "Filing of TDS Return in Form 26Q / 24Q"


class WebTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.dir = Path(temp.name)
        for name in ("Work Schedule Master.xlsx", "Work Schedule.xlsx"):
            shutil.copy(ROOT / name, self.dir / name)
        for attr, name in (("MASTER_FILE", "Work Schedule Master.xlsx"),
                           ("SCHEDULE_FILE", "Work Schedule.xlsx"),
                           ("ALLOCATION_FILE", "Work Allocation.xlsx")):
            patcher = mock.patch.object(core, attr, self.dir / name)
            patcher.start()
            self.addCleanup(patcher.stop)

    def run_app(self, name):
        app = AppTest.from_file(str(WEB / name), default_timeout=30).run()
        self.assertFalse(app.exception, app.exception)
        return app

    def select(self, app, label, value):
        next(s for s in app.selectbox if s.label == label).select(value).run()
        self.assertFalse(app.exception, app.exception)

    def button(self, app, label):
        next(b for b in app.button if b.label == label).click().run()
        self.assertFalse(app.exception, app.exception)


class SchedulerWeb(WebTest):
    def fill(self, app):
        self.select(app, "1. Client", CLIENT)
        self.select(app, "2. Assignment", ASSIGNMENT)
        self.select(app, "3. Work Type / Area", WORK_TYPE)
        self.select(app, "4. Frequency", "Quarterly")
        self.select(app, "5. Financial Year (creates 4 line items)", "FY 2026-27")

    def test_opens_with_sample_data(self):
        app = self.run_app("Work_Scheduler_Web.py")
        self.assertEqual(app.metric[0].value, "29")

    def test_add_then_submit(self):
        app = self.run_app("Work_Scheduler_Web.py")
        self.fill(app)
        self.button(app, "＋ Add to Pending")
        self.assertEqual(len(app.session_state["pending"]), 4)
        self.assertEqual(len(core.load_schedule()), 29)     # nothing saved before Submit
        self.button(app, "✓ Submit to Scheduler")
        self.assertEqual(len(core.load_schedule()), 33)
        self.assertEqual(app.metric[0].value, "33")

    def test_duplicate_is_refused(self):
        app = self.run_app("Work_Scheduler_Web.py")
        self.fill(app)
        self.button(app, "＋ Add to Pending")
        self.button(app, "✓ Submit to Scheduler")
        self.fill(app)
        self.button(app, "＋ Add to Pending")
        self.assertEqual(app.session_state["pending"], [])
        self.assertTrue(app.warning)

    def test_saved_file_reads_in_the_desktop_format(self):
        entries = core.load_schedule()
        entries[0]["Due Date"] = date(2026, 12, 31)
        core.save_schedule(entries)
        ws = openpyxl.load_workbook(core.SCHEDULE_FILE)[core.SCHEDULE_SHEET]
        self.assertEqual([c.value for c in ws[1]], core.SCHEDULE_COLUMNS)
        self.assertEqual(ws.max_row, 30)
        self.assertEqual(core.load_schedule()[0]["Due Date"], date(2026, 12, 31))


class AllocatorWeb(WebTest):
    def test_opens_with_all_tabs(self):
        app = self.run_app("Work_Allocator_Web.py")
        self.assertEqual([t.label for t in app.tabs],
                         ["Work Allocation", "Resource's Schedule (Day)", "Weekly Schedule"])

    def test_allocate_save_and_reload(self):
        app = self.run_app("Work_Allocator_Web.py")
        schedule = core.load_schedule()
        task = core.task_key(next(e for e in schedule if not core.is_complete(e)))
        book = app.session_state["book"]
        key = app.session_state["month"]
        resource, day = sorted(book.resources)[0], date(*key, 10)
        book.set_tasks(resource, day, [task])
        app.session_state["dirty"] = True
        app.run()
        self.assertFalse(app.exception, app.exception)
        self.button(app, "✓ Save")
        self.assertEqual(core.load_allocations()[key].tasks[resource, day], [task])

    def test_outside_change_is_not_overwritten_silently(self):
        app = self.run_app("Work_Allocator_Web.py")
        core.Book({}, {}).save()            # another session saves first
        app.session_state["dirty"] = True
        app.run()
        self.button(app, "✓ Save")
        self.assertTrue(app.session_state["save_conflict"])
        self.assertTrue(any(b.label == "Save anyway" for b in app.button))

    def test_exports(self):
        masters = core.Masters()
        schedule = core.load_schedule()
        task = core.task_key(next(e for e in schedule if not core.is_complete(e)))
        book = core.Book({}, masters.resources)
        resource, day = sorted(masters.resources)[0], date(2026, 11, 10)
        book.set_tasks(resource, day, [task])
        rows, removed = book.day_rows(resource, day, schedule, {})
        self.assertEqual((len(rows), removed), (1, 0))
        dates = core.week_of(day)
        week = book.wwa_rows(dates, True)
        for data in (core.day_xlsx(resource, day, rows), core.week_xlsx(dates, week, True),
                     core.month_xlsx(book, (2026, 11))):
            self.assertEqual(data[:2], b"PK")
        for data in (core.day_pdf(resource, day, rows), core.week_pdf(dates, week)):
            self.assertEqual(data[:4], b"%PDF")


class SharedWithDesktop(WebTest):
    """A file saved by the web app opens in the desktop app, and the reverse."""

    def setUp(self):
        super().setUp()
        import test_apps
        self.allocator = test_apps.allocator

    def test_web_allocations_load_in_desktop(self):
        schedule = core.load_schedule()
        task = core.task_key(next(e for e in schedule if not core.is_complete(e)))
        book = core.Book({}, core.Masters().resources)
        resource, day = sorted(book.resources)[0], date(2026, 11, 10)
        book.set_tasks(resource, day, [task])
        book.set_manager(resource, (2026, 11), "AM")
        book.save()
        with mock.patch.object(self.allocator, "ALLOCATION_FILE", core.ALLOCATION_FILE):
            months = self.allocator.load_allocations()
        self.assertEqual(months[2026, 11].tasks[resource, day], [task])
        self.assertEqual(months[2026, 11].managers[resource], "AM")

    def test_task_keys_match_desktop(self):
        for e in core.load_schedule():
            self.assertEqual(core.task_key(e), self.allocator.task_key(e))


if __name__ == "__main__":
    unittest.main()
