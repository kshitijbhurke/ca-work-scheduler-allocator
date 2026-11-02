# Work Scheduler and Work Allocator (Desktop)

[![Tests](https://github.com/kshitijbhurke/ca-work-scheduler-allocator/actions/workflows/tests.yml/badge.svg)](https://github.com/kshitijbhurke/ca-work-scheduler-allocator/actions/workflows/tests.yml)

Two standalone Windows desktop apps for Chartered Accountant firms. They use Excel files in this folder and don't need a database, a server or any other folder.

- **Work Scheduler** builds the list of client work.
- **Work Allocator** allots that work to Resources, date by date, and produces daily and weekly schedules.

The clients, managers and resources supplied here are fictional samples. Replace them with your own before real use.

For what the apps offer and where they're heading, see the [Product Overview](Product%20Overview.md).

## Try it in two minutes

These are Windows desktop apps, so they run on your PC, not on GitHub.

1. Install [Python 3](https://www.python.org/downloads/) for Windows. On the first installer screen, tick **Add python.exe to PATH**.
2. On this page, click **Code**, then **Download ZIP**, and extract the ZIP to a normal folder.
3. Double-click `Run Work Scheduler.bat`. On first use it installs the two packages it needs, which requires the internet once.
4. Double-click `Run Work Allocator.bat`. The 29 sample rows are ready to allocate.

Automated tests run both apps on Windows each time the code changes. To run them yourself:

```bash
python -m unittest discover -s tests -v
```

The code is published for evaluation only. See [LICENSE](LICENSE).

## Files

| File | Purpose |
|---|---|
| `Work_Scheduler.py` | The Work Scheduler app |
| `Work_Allocator.py` | The Work Allocator app |
| `Run Work Scheduler.bat` | Double-click to start the Work Scheduler |
| `Run Work Allocator.bat` | Double-click to start the Work Allocator |
| `Work Schedule Master.xlsx` | Master lists. The apps only read this file |
| `Work Schedule.xlsx` | The saved schedule, with 29 sample rows. The Work Scheduler writes this file and the Work Allocator reads it |
| `Work Allocation.xlsx` | The saved allocations. Created at your first Save in the Work Allocator |
| `requirements.txt` | Python packages needed |
| `Product Overview.md` | What the apps offer, and the path to larger firms |
| `tests/test_apps.py` | Automated checks of both apps, run on a temporary copy of the sample files |

Keep the Excel files in the same folder as the apps. Neither `.py` file needs the other, but the Work Allocator needs a `Work Schedule.xlsx` to allocate from.

## Setup

Install Python 3 for Windows, including Tkinter, which the standard installer includes. The `.bat` files install the required packages on first use. To install them yourself, run this in the folder:

```bash
pip install -r requirements.txt
```

`reportlab` is used only for PDF export.

## Master workbook

| Sheet | Columns used | Used for |
|---|---|---|
| `Client List` | Client Full Name, Client Short Name, Client Group | Client lists |
| `Assignment Master` | Assignment | Assignment lists. Every assignment used in the two sheets below must be listed here |
| `Work Type Master` | Assignment, Work Type / Area | Work types for each assignment |
| `Frequency Master` | Assignment, Work Type / Area, Frequency | Frequencies for each work type |
| `Manager Master` | Manager Name, Mgr Short Name | Manager lists |
| `Resource Master` | Resource Full Name, M/F | Resources in the Work Allocator, listed A to Z |
| `WWA Format`, `Resource Day Details Format` | None | Reference only. They show the layouts the Work Allocator produces |

Keep sheet names and column headings exactly as they are. The apps pick up saved changes to the master within a few seconds, even while it's open in Excel.

To start with your own data, replace the rows in the master, then close the apps and delete `Work Schedule.xlsx` and `Work Allocation.xlsx`. The apps create new ones.

## Work Scheduler

1. Under **New Schedule Entry**, choose Client, Assignment, Work Type / Area and Frequency.
2. Enter the period. Each Frequency creates its line items as shown below.
3. Click **＋ Add to Pending**, review the list, then click **✓ Submit to Scheduler**. Nothing is saved until you submit.
4. To update scheduled work, tick rows, fill in Manager, Due Date, Target Date or Work Status, then click **Apply to Selected**.

| Frequency | You enter | Line items |
|---|---|---|
| Monthly | Financial Year | 12 |
| Quarterly | Financial Year | 4 |
| Half Yearly | Financial Year | 2 |
| Annual | Financial Year | 1 |
| NA | Nothing | 1, with the period shown as NA |
| Others, e.g. Annual Dec End, As per Engagement | From and To dates | 1 |

Financial years run from April to March. Duplicates are skipped, with a pop-up listing them. You can also search, delete rows, and use **Export to Excel** with your choice of fields.

Row colours: light amber is WIP, light green is Complete, and red text means the Due Date has passed and the work isn't Complete.

## Work Allocator

### Work Allocation tab

The main screen shows Resources against every date of the month. Use **‹** and **›** to change month.

- **Single click on a cell:** shows the work allotted that day, in the `Resource Day Details Format` layout.
- **Double click on a cell:** opens the allocation window. Choose Client Short Name, Assignment and Financial Year, tick pending tasks, then click **✓ Allocate**. Complete tasks aren't listed.
- **Manager column:** click to choose the Resource's manager. The choice carries into later months.
- **Resource heading:** click to switch between A to Z and Z to A.
- **Clear Tasks:** appears in the single-click and double-click windows, and clears chosen tasks for that day.
- **Clear Allocations:** lists the whole month with tick boxes and a search, to clear many together.
- **✓ Save:** nothing is saved until you click this. **⟳ Reload** discards unsaved changes.

Cell colours: blue is allocated, amber means the work is no longer in `Work Schedule.xlsx`, and light orange is a Sunday or a 1st or 3rd Saturday. Those days can still be allocated.

### Resource's Schedule (Day) tab

Choose a Resource and a day, then click **Show**. Use **Export to Excel** or **Export to PDF**.

### Weekly Schedule tab

Shows all Resources from Monday to Sunday.

- **This Week**, **Next Week** and **Custom Week** choose the week.
- **Show Detailed** adds Work Type / Area and the period covered. **Show Concise** switches back.
- **Export to Excel** and **Export to PDF** save the week as it's displayed, including changes you haven't saved yet.

## Good to know

- **Close the Excel file before saving from an app.** While `Work Schedule.xlsx` or `Work Allocation.xlsx` is open in Excel, the app can't save. Nothing is lost, and you can save again after closing it.
- **Editing `Work Schedule.xlsx` directly:** close and reopen the Work Scheduler first, because saving from an open app overwrites the file with what it has in memory.
- **Editing `Work Allocation.xlsx` directly:** the Work Allocator warns you before saving over the change. Click **⟳ Reload** to bring it in.
- **Renaming:** if you rename a `.py` file, change its `.bat` file to match.
- **Backups:** the two saved Excel files are the only copies of your data. Keep them backed up.
