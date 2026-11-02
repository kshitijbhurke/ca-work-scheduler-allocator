# Work Scheduler & Work Allocator: Product Overview

## Summary

Work Scheduler and Work Allocator give a Chartered Accountant firm one connected view of what work is due and who is doing it each day. The Work Scheduler turns the firm's client engagements into a dated list of work for the financial year. The Work Allocator assigns that work to articles and staff, date by date, and produces daily and weekly schedules.

The two Windows desktop apps are ready for small and mid-size firms today. They keep all data in Excel files on the firm's own computers, with no server or subscription. Larger firms will need a planned upgrade that adds logins and access controls.

## The problem

Most small and mid-size CA firms run recurring compliance and audit work across many clients with a small team. The work calendar and the staff calendar usually live in separate spreadsheets, messages and memories.

- **Recurring work is rebuilt by hand.** A monthly engagement means twelve separate items to set up each year.
- **Allocation is hard to see.** Partners cannot easily tell who is free, who is overloaded, or which work has no one on it.
- **Status lives with individuals.** Finding what is pending or overdue means asking around.
- **Enterprise practice software can be too much.** It may bring costs, setup and training that a smaller firm does not need.

## What the apps do

The firm maintains its masters once, the Work Scheduler creates the work, and the Work Allocator puts that work on people's calendars.

### Work Scheduler

- **Guided entry:** choose client, assignment, work type and frequency in steps, each list filtered by the one before it.
- **Automatic line items:** one entry for a financial year (April to March) creates 12 monthly, 4 quarterly, 2 half-yearly or 1 annual item. Other periods use From and To dates.
- **Review before saving:** entries wait in a pending list until submitted, and duplicates are skipped with a list of what was left out.
- **Bulk updates:** tick rows to set manager, due date, target date or status together.
- **Status at a glance:** amber for work in progress, green for complete, red text for overdue, plus live totals.
- **Search and export:** multi-word search, and an Excel export with a choice of fields.

### Work Allocator

- **Monthly grid:** every Resource against every date, with Sundays and the 1st and 3rd Saturdays shaded.
- **Fast allocation:** double-click a cell, choose client, assignment and financial year, and tick pending tasks. Completed work is not offered.
- **Day view:** a single click shows that person's tasks for the day, with client, period, target date and due date.
- **Weekly schedule:** all Resources from Monday to Sunday, in a concise or detailed view.
- **Exports:** day and week schedules in Excel and PDF, ready to circulate, and the month in Excel.
- **Change control:** nothing is saved until Save is clicked, the app warns if the saved file changed outside it, and work later removed from the schedule is flagged in amber.

## Benefits by role

| Role | What they gain |
| --- | --- |
| Partner | One view of who is working on what, and which work is overdue |
| Manager | Quicker allocation, and bulk updates to dates and status across clients |
| Article or staff member | A clear schedule for the day or week, with target and due dates |
| Office administrator | Client, assignment and staff lists kept in one familiar Excel workbook |

## How it is built, and data privacy

Client data stays on the firm's own computers. After the first-time package setup, the apps need no internet connection, and they do not send data anywhere.

- **Plain Excel storage:** the firm can open, back up and print its data without the apps.
- **Configurable masters:** clients, assignments, work types, frequencies, managers and staff are edited in one workbook, and changes appear within seconds.
- **Low setup:** Python on a Windows PC and two standard packages, set up automatically on first run.
- **Tested:** automated tests run both apps on Windows, covering scheduling, allocation, saving and exports.

## Growth path to larger firms

The current apps suit firms where a few trusted people maintain the schedule. Access controls are planned and not yet built, so the steps below are the proposed route, not current features.

| Stage | Status | What it adds |
| --- | --- | --- |
| Small and mid-size firms | Available today | Desktop apps on Windows, data in Excel files, access set by folder sharing, one person editing at a time |
| Growing teams | Planned | Shared database, individual logins, several people working at once, Excel import and export kept |
| Large firms | Planned | Roles for partner, manager and staff, access by branch or team, audit trail of changes, partner-level reporting |

The master workbook already has a Manager Login ID column ready for individual logins. The firm's masters and working layouts carry forward, so a firm can start small and keep what it has set up.

## Suggested five-minute demonstration

1. Open the Work Scheduler and point out the sample clients, status colours and totals.
2. Add a Quarterly entry for one client and financial year, and show the 4 line items it creates.
3. Add the same entry again to show duplicates being stopped, then submit the first set.
4. Tick several rows and apply a manager, due date and status together.
5. Open the Work Allocator, double-click a date for one Resource, tick tasks and allocate.
6. Single-click that cell to show the day view, then click Save.
7. Open the Weekly Schedule, switch to the detailed view, and export it to PDF.

## Trying it, current limits and next step

The apps can be tried with the fictional sample data included. The steps are in the [README](README.md), and the code is published for evaluation only.

These are the current limits to raise openly:

- **Windows only.** The apps need Python on a Windows PC.
- **One editor at a time.** Two people saving the same file can overwrite each other, so one person should edit each file at a time.
- **No logins yet.** Anyone who can open the folder can see and change all of its data.
- **Backups are the firm's job.** The Excel files are the only copy of the data.

The suggested next step is a short demonstration, followed by a trial using one team's real client list.
