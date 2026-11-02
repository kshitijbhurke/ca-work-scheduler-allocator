"""
Work Allocator (web)
--------------------
Streamlit version of the desktop Work Allocator, on the same Excel files.

    streamlit run web/Work_Allocator_Web.py

Changes are held for this browser session until Save writes them to 'Work Allocation.xlsx'.
"""

import calendar
from datetime import date, timedelta
from html import escape

import pandas as pd
import streamlit as st

import core
import theme

LEFT_COLUMNS = ("#", "Resource", "M/F", "Manager")
NO_MANAGER = "— None —"
TAB_NAMES = ("Work Allocation", "Resource's Schedule (Day)", "Weekly Schedule")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@st.cache_resource(show_spinner=False)
def read_masters(path, stamp):
    return core.Masters(path)


def flash(kind, text):
    st.session_state["flash"] = (kind, text)


def show_flash():
    kind, text = st.session_state.pop("flash", (None, None))
    if kind:
        getattr(st, kind)(text)


def load_book(resources):
    state = st.session_state
    state["book"] = core.Book(core.load_allocations(), resources, state.get("descending", False))
    state["file_stamp"] = core.mtime(core.ALLOCATION_FILE)
    state["dirty"] = False


def go_to(key):
    st.session_state["month"] = key


def column_name(d):
    return f"{d:%d %a}"


def cell_display(work):
    """First client and assignment; further pairs are only counted."""
    lines = (work or "").split("\n")
    return f"{lines[0]}  +{len(lines) - 1}" if len(lines) > 1 else lines[0]


def grid(book, key, schedule_keys, work_options):
    """The month as a styled table. Returns (frame, styler, dates)."""
    dates = core.month_dates(*key)
    today = date.today()
    data = book.month(key)
    rows, styles = [], []
    for sr, (name, gender, manager, work) in enumerate(book.wwa_rows(dates), start=1):
        row = {"#": sr, "Resource": name, "M/F": gender, "Manager": manager}
        css = dict.fromkeys(LEFT_COLUMNS, "")
        if name not in book.resources:
            css["Resource"] = f"color: {theme.AMBER_TEXT};"
        for d in dates:
            column = column_name(d)
            row[column] = cell_display(work.get(d))
            fill = ""
            if work.get(d):
                tasks = data.tasks.get((name, d))
                removed = (any(t not in schedule_keys for t in tasks) if tasks
                           else work[d] not in work_options)
                fill = (f"background-color: {theme.AMBER_BG}; color: {theme.AMBER_TEXT};"
                        if removed else f"background-color: {theme.ALLOCATED_BG};")
            elif core.is_off_day(d):
                fill = f"background-color: {theme.OFF_DAY_BG};"
            elif d == today:
                fill = f"background-color: {theme.TODAY_BG};"
            css[column] = fill
        rows.append(row)
        styles.append(css)
    columns = [*LEFT_COLUMNS, *(column_name(d) for d in dates)]
    frame = pd.DataFrame(rows, columns=columns)
    style_frame = pd.DataFrame(styles, columns=columns, index=frame.index).fillna("")
    return frame, frame.style.apply(lambda _f: style_frame, axis=None), dates


def apply_cell_click(event, frame, dates):
    """Move the Resource and date pickers to a newly clicked cell."""
    state = st.session_state
    cells = event.selection.cells if event else []
    if not cells:
        state.pop("last_cell", None)
        return
    row, column = cells[0]
    if state.get("last_cell") == (state["month"], row, column):
        return
    state["last_cell"] = (state["month"], row, column)
    day = next((d for d in dates if column_name(d) == column), None)
    if day and 0 <= row < len(frame):
        state["pick_resource"] = frame.iloc[row]["Resource"]
        state["pick_day"] = day


def tick_list(rows, ticks, key):
    """Rows of a list with a tick-box column; returns the positions of the ticked rows."""
    frame = pd.DataFrame(rows)
    frame.insert(0, "Select", ticks)
    edited = st.data_editor(frame, key=key, hide_index=True, width="stretch",
                            disabled=[c for c in frame.columns if c != "Select"],
                            column_config={"Select": st.column_config.CheckboxColumn(
                                "Select", width="small")})
    return [i for i, value in enumerate(edited["Select"]) if value]


def set_tasks(resource, day, tasks):
    st.session_state["book"].set_tasks(resource, day, tasks)
    st.session_state["dirty"] = True


def allocate_panel(resource, day, schedule, short_names):
    state = st.session_state
    book = state["book"]
    chosen = list(book.month((day.year, day.month)).tasks.get((resource, day), []))
    stem = f"{resource}_{day.isoformat()}"
    st.markdown("**Allocate work**")

    def matching(*filters):
        rows = schedule
        for get, value in filters:
            rows = [e for e in rows if get(e) == value]
        return rows

    get_client = lambda e: core.clean(e.get("Client Short Name"))
    get_assignment = lambda e: core.clean(e.get("Assignment"))
    options = lambda get, rows: sorted(core.unique(get(e) for e in rows), key=core.sort_key)
    first = chosen[0] if chosen else None

    def pick(column, label, values, default, key):
        index = values.index(default) if default in values else None
        return column.selectbox(label, values, index=index, key=key, placeholder="Choose")

    a, b, c = st.columns(3)
    client = pick(a, "Client Short Name", options(get_client, schedule),
                  first and first[0], f"client_{stem}")
    by_client = matching((get_client, client))
    assignment = pick(b, "Assignment", options(get_assignment, by_client),
                      first and first[1], f"assignment_{stem}_{client}")
    by_assignment = matching((get_client, client), (get_assignment, assignment))
    fy = pick(c, "Financial Year", options(core.financial_year, by_assignment),
              first and (first[4] or core.NA), f"fy_{stem}_{client}_{assignment}")
    if not (client and assignment and fy):
        st.caption("Choose all three to list the pending tasks.")
        return

    rows = [e for e in matching((get_client, client), (get_assignment, assignment),
                                (core.financial_year, fy)) if not core.is_complete(e)]
    if not rows:
        st.info("There are no tasks pending under this selection.")
        return
    keys = [core.task_key(e) for e in rows]
    listing = [{field: core.fmt_cell(e.get(field)) if e.get(field) is not None else ""
                for field in ("Work Type / Area", "Frequency", "Period From", "Period To",
                              "Target Date", "Due Date", "Work Status")}
               | {"Manager": short_names.get(core.clean(e.get("Manager")))
                  or e.get("Manager Short Name") or ""}
               for e in rows]
    version = state.get("dirty_count", 0)
    ticked = tick_list(listing, [k in chosen for k in keys],
                       f"pick_{stem}_{client}_{assignment}_{fy}_{version}")
    picked = [k for k in chosen if k not in keys] + [keys[i] for i in ticked]
    st.caption(f"{len(picked)} task{'' if len(picked) == 1 else 's'} selected for this day")
    if st.button("✓ Allocate", type="primary", key=f"allocate_{stem}"):
        set_tasks(resource, day, picked)
        state["dirty_count"] = version + 1
        flash("success", f"Allocated. Click Save to keep this change.")
        st.rerun()


def clear_items(items):
    state = st.session_state
    book = state["book"]
    for resource, day, task in items:
        tasks = book.month((day.year, day.month)).tasks.get((resource, day), [])
        set_tasks(resource, day, [t for t in tasks if t != task])
    state["dirty_count"] = state.get("dirty_count", 0) + 1


def item_row(resource, day, task, work):
    client, assignment, work_type, period = "", work, "", ""
    if task:
        client, assignment, work_type = task[0], task[1], task[2]
        period = core.task_period(task)
    return {"Resource": resource, "Date": day.strftime(core.DATE_FMT), "Day": f"{day:%A}",
            "Client Short Name": client, "Assignment": assignment,
            "Work Type / Area": work_type, "Period Covered": period}


def clear_panel(items, key):
    """Tick allocated work and clear it. items = [(resource, date, task or None, work)]."""
    state = st.session_state
    rows = [item_row(*item) for item in items]
    ticked = tick_list(rows, [False] * len(rows), f"{key}_{state.get('dirty_count', 0)}")
    if st.button(f"Clear Selected ({len(ticked)})", key=f"btn_{key}", disabled=not ticked):
        clear_items([items[i][:3] for i in ticked])
        flash("success", f"Cleared {len(ticked)}. Click Save to keep this change.")
        st.rerun()


def day_table(rows):
    frame = pd.DataFrame(rows, columns=list(core.DAY_COLUMNS))
    st.dataframe(frame, hide_index=True, width="stretch")


def downloads(stem, excel, pdf, key):
    left, right, _rest = st.columns([1, 1, 3])
    left.download_button("Export to Excel", excel, f"{stem}.xlsx", XLSX, key=f"xlsx_{key}",
                         width="stretch")
    right.download_button("Export to PDF", pdf, f"{stem}.pdf", "application/pdf",
                          key=f"pdf_{key}", width="stretch")


def day_view(resource, day, schedule, short_names, key):
    book = st.session_state["book"]
    work = book.month((day.year, day.month)).cells.get((resource, day))
    if not work:
        st.caption("No work allocated on this date.")
        return
    rows, removed = book.day_rows(resource, day, schedule, short_names)
    if removed or not rows:
        st.warning(f"Some allocated work is no longer in {core.SCHEDULE_FILE.name}.")
    day_table(rows)
    downloads(f"Day Schedule {resource} {day.isoformat()}", core.day_xlsx(resource, day, rows),
              core.day_pdf(resource, day, rows), key)


def allocation_tab(masters, schedule):
    state = st.session_state
    book = state["book"]
    key = state["month"]
    schedule_keys = {core.task_key(e) for e in schedule}
    work_options = set(core.unique(core.work_label(e) for e in schedule))
    short_names = {full: short for short, full in masters.short_names()}

    nav = st.columns([0.4, 1.3, 0.4, 0.9, 0.9, 0.9, 0.9, 0.8], vertical_alignment="center")
    nav[0].button("‹", on_click=go_to, args=(core.shift_month(key, -1),), key="month_back")
    nav[1].markdown(f"### {calendar.month_name[key[1]]} {key[0]}")
    nav[2].button("›", on_click=go_to, args=(core.shift_month(key, 1),), key="month_next")
    today = date.today()
    nav[3].button("This Month", on_click=go_to, args=((today.year, today.month),))
    book.descending = nav[4].toggle("Z to A", key="descending")
    if nav[5].button("⟳ Reload", help="Discards unsaved changes"):
        load_book(masters.resources)
        flash("info", "Reloaded from the saved file.")
        st.rerun()
    nav[6].download_button("Export to Excel", core.month_xlsx(book, key),
                           f"Work Allocation Export {core.sheet_name(key)}.xlsx", XLSX,
                           key="month_export")
    if nav[7].button("✓ Save", type="primary", disabled=not state["dirty"]):
        save(masters)

    if state.get("save_conflict"):
        st.warning(f"{core.ALLOCATION_FILE.name} was changed outside this session. "
                   "Saving replaces those changes.")
        if st.button("Save anyway"):
            save(masters, force=True)

    if state["dirty"]:
        st.warning("Unsaved changes. Click Save before closing this page.")
    elif not schedule:
        st.info(f"No work in {core.SCHEDULE_FILE.name} yet.")
    theme.legend(("Allocated", theme.ALLOCATED_BG, theme.TEXT),
                 ("No longer in schedule", theme.AMBER_BG, theme.AMBER_TEXT),
                 ("Off day", theme.OFF_DAY_BG, theme.TEXT), ("Today", theme.TODAY_BG, theme.TEXT))

    frame, styled, dates = grid(book, key, schedule_keys, work_options)
    config = {"#": st.column_config.Column(width="small", pinned=True),
              "Resource": st.column_config.Column(width="medium", pinned=True),
              "M/F": st.column_config.Column(width="small"),
              "Manager": st.column_config.Column(width="small")}
    config |= {column_name(d): st.column_config.Column(width="medium") for d in dates}
    event = st.dataframe(styled, hide_index=True, width="stretch", column_config=config,
                         on_select="rerun", selection_mode="single-cell",
                         key=f"grid_{key}_{state.get('dirty_count', 0)}")
    apply_cell_click(event, frame, dates)
    st.caption("Click a date cell to see and allocate that day's work.")

    names = book.resources_for(key)
    if state.get("pick_resource") not in names:
        state["pick_resource"] = names[0] if names else None
    if state.get("pick_day") not in dates:
        state["pick_day"] = today if today in dates else dates[0]
    if not names:
        st.info("No Resources in the master workbook.")
        return

    with st.container(border=True):
        a, b, c = st.columns(3)
        resource = a.selectbox("Resource", names, key="pick_resource")
        day = b.selectbox("Date", dates, key="pick_day",
                          format_func=lambda d: d.strftime(core.DAY_TITLE_FMT))
        managers = [NO_MANAGER, *(f"{short}{core.SEPARATOR}{full}"
                                  for short, full in masters.short_names())]
        current = book.manager(resource, key)
        index = next((i for i, m in enumerate(managers)
                      if current and m.startswith(current + core.SEPARATOR)), 0)
        manager = c.selectbox("Manager for this Resource", managers, index=index,
                              key=f"manager_{resource}_{key}_{state.get('dirty_count', 0)}")
        short = "" if manager == NO_MANAGER else manager.split(core.SEPARATOR)[0]
        if short != current:
            book.set_manager(resource, key, short)
            state["dirty"] = True
            state["dirty_count"] = state.get("dirty_count", 0) + 1
            st.rerun()

        st.markdown(f"**{resource}**: {core.day_title(day)}")
        day_view(resource, day, schedule, short_names, "cell")
        allocate_panel(resource, day, schedule, short_names)
        tasks = book.month((day.year, day.month)).tasks.get((resource, day))
        work = book.month((day.year, day.month)).cells.get((resource, day))
        if work:
            with st.expander("Clear Tasks for this day"):
                clear_panel([(resource, day, t, work) for t in tasks or [None]],
                            f"clear_day_{resource}_{day}")

    data = book.month(key)
    items = []
    for name in names:
        for d in dates:
            work = data.cells.get((name, d))
            if work:
                items += [(name, d, t, work) for t in data.tasks.get((name, d)) or [None]]
    with st.expander("Clear Allocations for this month"):
        query = st.text_input("Search", key="clear_query").casefold().split()
        shown = [i for i in items
                 if all(w in " ".join(item_row(*i).values()).casefold() for w in query)]
        if shown:
            clear_panel(shown, f"clear_month_{key}_{'_'.join(query)}")
        else:
            st.caption(f"No work is allocated in {calendar.month_name[key[1]]} {key[0]}.")


def save(masters, force=False):
    state = st.session_state
    book = state["book"]
    if not force and core.mtime(core.ALLOCATION_FILE) != state["file_stamp"]:
        state["save_conflict"] = True
        st.rerun()
    try:
        book.save()
    except PermissionError:
        flash("error", f"{core.ALLOCATION_FILE.name} is open in Excel. Close it and click Save "
                       "again. Nothing has been lost.")
        st.rerun()
    state["file_stamp"] = core.mtime(core.ALLOCATION_FILE)
    state["dirty"] = False
    state.pop("save_conflict", None)
    flash("success", "All changes saved.")
    st.rerun()


def day_tab(masters, schedule):
    book = st.session_state["book"]
    short_names = {full: short for short, full in masters.short_names()}
    names = sorted(masters.resources, key=core.sort_key)
    a, b = st.columns(2)
    resource = a.selectbox("Resource", names, index=None, key="day_resource",
                           placeholder="Choose a Resource")
    day = b.date_input("Day", value=date.today(), key="day_date", format="DD/MM/YYYY")
    if resource and day:
        st.markdown(f"**{resource}**: {core.day_title(day)}")
        day_view(resource, day, schedule, short_names, "tab")


def week_table(dates, rows):
    """All Resources from Monday to Sunday, keeping line breaks inside the cells."""
    today = date.today()
    head = "".join(f"<th>{escape(h.replace(chr(10), ' '))}</th>" for h in core.LEFT_HEADERS)
    for d in dates:
        fill = "background:#2563EB;" if d == today else ""
        head += f'<th style="{fill}">{d:%A}<br>{d:%d-%b-%y}</th>'
    body = ""
    for sr, (resource, gender, manager, work) in enumerate(rows, start=1):
        body += (f"<tr><td>{sr}</td><td class='name'>{escape(resource)}</td>"
                 f"<td>{escape(gender)}</td><td>{escape(manager)}</td>")
        for d in dates:
            fill = (theme.ALLOCATED_BG if work.get(d)
                    else theme.OFF_DAY_BG if core.is_off_day(d) else "#FFFFFF")
            body += (f'<td style="background:{fill}">'
                     f"{escape(work.get(d, '')).replace(chr(10), '<br>')}</td>")
        body += "</tr>"
    st.html(f"""
    <style>
    .week {{ overflow-x: auto; }}
    .week table {{ border-collapse: collapse; width: 100%; font-size: 0.8rem; color: {theme.TEXT}; }}
    .week th {{ background: {theme.NAVY}; color: white; padding: 6px; border: 1px solid #D7DEE8; }}
    .week td {{ padding: 6px; border: 1px solid #D7DEE8; text-align: center; min-width: 110px;
                background: white; }}
    .week td:nth-child(-n+4) {{ min-width: 0; }}
    .week td.name {{ font-weight: 600; text-align: left; min-width: 140px; }}
    </style>
    <div class="week"><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>
    """)


def week_tab():
    book = st.session_state["book"]
    today = date.today()
    a, b, c = st.columns([2, 1, 1], vertical_alignment="bottom")
    mode = a.radio("Week", ("This Week", "Next Week", "Custom Week"), horizontal=True,
                   key="week_mode")
    pick = today
    if mode == "Next Week":
        pick = today + timedelta(days=7)
    elif mode == "Custom Week":
        pick = b.date_input("Any day in the week", value=today, key="week_day",
                            format="DD/MM/YYYY") or today
    detailed = c.toggle("Show Detailed", key="week_detailed")
    dates = core.week_of(pick)
    rows = book.wwa_rows(dates, detailed)
    st.markdown(f"**{dates[0].strftime(core.DATE_FMT)} to {dates[-1].strftime(core.DATE_FMT)}**")
    week_table(dates, rows)
    stem = f"Weekly Schedule {dates[0].isoformat()}{' Detailed' if detailed else ''}"
    downloads(stem, core.week_xlsx(dates, rows, detailed), core.week_pdf(dates, rows), "week")


def main():
    st.set_page_config(page_title="Work Allocator", layout="wide")
    theme.header("Work Allocator", "Allocate scheduled work to Resources, date by date")
    state = st.session_state
    try:
        masters = read_masters(core.MASTER_FILE, core.mtime(core.MASTER_FILE))
        schedule = core.load_schedule()
    except FileNotFoundError:
        st.error(f"{core.MASTER_FILE.name} was not found next to the apps.")
        st.stop()
    except (KeyError, ValueError) as exc:
        st.error(f"Could not read the lists: {exc}")
        st.stop()

    today = date.today()
    state.setdefault("month", (today.year, today.month))
    if "book" not in state:
        load_book(masters.resources)
    elif not state["dirty"] and core.mtime(core.ALLOCATION_FILE) != state["file_stamp"]:
        load_book(masters.resources)
    state["book"].resources = masters.resources

    book, key = state["book"], state["month"]
    a, b, c = st.columns(3)
    a.metric("Resources", len(book.resources_for(key)))
    b.metric("Work Options", len(core.unique(core.work_label(e) for e in schedule)))
    c.metric("Allocated This Month", len(book.month(key).cells))
    show_flash()

    tabs = st.tabs(TAB_NAMES)
    with tabs[0]:
        allocation_tab(masters, schedule)
    with tabs[1]:
        day_tab(masters, schedule)
    with tabs[2]:
        week_tab()


main()
