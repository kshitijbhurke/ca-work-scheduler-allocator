"""
Work Scheduler (web)
--------------------
Streamlit version of the desktop Work Scheduler, on the same Excel files.

    streamlit run web/Work_Scheduler_Web.py
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

import core
import theme

NOT_STARTED = "Not started (blank)"
NO_MANAGER = "— None —"
NO_CHANGE = "(no change)"
TABLE_COLUMNS = (   # (schedule column, heading)
    ("Sr No", "#"), ("Client Short Name", "Client"), ("Client Group", "Client Group"),
    ("Assignment", "Assignment"), ("Work Type / Area", "Work Type / Area"),
    ("Work Status", "Status"), ("Frequency", "Frequency"), ("Financial Year", "FY"),
    ("Period From", "Period From"), ("Period To", "Period To"), ("Due Date", "Due Date"),
    ("Target Date", "Target Date"), ("Manager", "Manager"),
    ("Manager Short Name", "Mgr Short Name"), ("Client Full Name", "Client Full Name"),
    ("Date Added", "Date Added"),
)
FORM_KEYS = ("assignment", "work_type", "frequency", "fy", "period_from", "period_to")


@st.cache_resource(show_spinner=False)
def read_masters(path, stamp):
    return core.Masters(path)


def flash(kind, text):
    st.session_state["flash"] = (kind, text)


def show_flash():
    kind, text = st.session_state.pop("flash", (None, None))
    if kind:
        getattr(st, kind)(text)


def save(entries, done):
    """Save the schedule. A locked file leaves it as it was."""
    try:
        core.save_schedule(entries)
    except PermissionError:
        flash("error", f"Can't save: {core.SCHEDULE_FILE.name} is open in Excel. "
                       "Close it and try again.")
        return False
    flash("success", done)
    return True


def plural(n):
    return f"{n} entr{'y' if n == 1 else 'ies'}"


def clear_from(*keys):
    for key in keys:
        st.session_state[key] = None


def add_to_pending(masters, entries):
    state = st.session_state
    items = core.build_line_items(masters, state.client, state.assignment, state.work_type,
                                  state.frequency, state.get("fy"), state.get("period_from"),
                                  state.get("period_to"))
    pending = state.setdefault("pending", [])
    new, dups = core.split_duplicates(items, pending, entries)
    pending.extend(new)
    if dups:
        flash("warning", f"{plural(len(dups))} already scheduled and not added: "
                         + "; ".join(core.describe(e) for e in dups))
    elif new:
        flash("success", f"Added {plural(len(new))} to Pending. Click Submit to save.")
    # Keep the client, FY and dates so several work types can be added quickly.
    clear_from("assignment", "work_type", "frequency")


def submit_pending():
    state = st.session_state
    entries = core.load_schedule()      # the latest file, in case someone else has saved
    to_add, skipped = core.split_duplicates(state.get("pending", []), entries)
    if to_add and not save(entries + to_add, f"{plural(len(to_add))} submitted to the Scheduler."):
        return
    state["pending"] = []
    if skipped:
        flash("warning", f"{plural(len(skipped))} already in the Scheduler and skipped; "
                         f"{plural(len(to_add))} submitted.")


def build_form(masters, entries):
    state = st.session_state
    st.subheader("New Schedule Entry")
    st.caption("Complete each step to unlock the next.")

    def sanitise(key, options):
        if state.get(key) is not None and state[key] not in options:
            state[key] = None

    sanitise("client", list(masters.clients))
    client = st.selectbox("1. Client", list(masters.clients), index=None, key="client",
                          placeholder="Choose a client",
                          on_change=clear_from, args=("assignment", "work_type", "frequency"))
    if not client:
        return
    st.caption(f"Short name **{masters.clients[client]}**  ·  "
               f"Group **{masters.client_groups.get(client) or 'None'}**")

    sanitise("assignment", masters.assignments)
    assignment = st.selectbox("2. Assignment", masters.assignments, index=None, key="assignment",
                              placeholder="Choose an assignment",
                              on_change=clear_from, args=("work_type", "frequency"))
    if not assignment:
        return
    work_types = masters.work_types_for(assignment)
    if not work_types:
        st.warning("No Work Type / Area is mapped to this Assignment in 'Work Type Master'.")
        return

    sanitise("work_type", work_types)
    work_type = st.selectbox("3. Work Type / Area", work_types, index=None, key="work_type",
                             placeholder="Choose a work type",
                             on_change=clear_from, args=("frequency",))
    if not work_type:
        return
    frequencies = masters.frequencies_for(assignment, work_type)

    sanitise("frequency", frequencies)
    frequency = st.selectbox("4. Frequency", frequencies, index=None, key="frequency",
                             placeholder="Choose a frequency")
    if not frequency:
        return

    mode, months = core.frequency_kind(frequency)
    ready = True
    if mode == "fy":
        start = core.current_fy_start()
        years = [core.fy_label(y) for y in range(start - 2, start + 5)]
        count = 12 // months
        fy = st.selectbox(f"5. Financial Year (creates {count} line "
                          f"item{'s' if count > 1 else ''})", years, index=None, key="fy",
                          placeholder="Choose a financial year")
        ready = bool(fy)
        if fy:
            periods = core.fy_periods(core.fy_start_year(fy), months)
            st.caption(f"{periods[0][0]:%d-%b-%Y} to {periods[-1][1]:%d-%b-%Y}")
    elif mode == "na":
        st.caption("5. Period From and To will be recorded as NA.")
    else:
        left, right = st.columns(2)
        first = left.date_input("5. From", value=None, key="period_from", format="DD/MM/YYYY")
        last = right.date_input("To", value=None, key="period_to", format="DD/MM/YYYY",
                                min_value=first)
        ready = bool(first and last and last >= first)
        if first and last and last >= first:
            days = (last - first).days + 1
            st.caption(f"{first:%d %b %Y} to {last:%d %b %Y}  ·  "
                       f"{days} day{'s' if days != 1 else ''}")

    st.button("＋ Add to Pending", type="primary", disabled=not ready, width="stretch",
              on_click=add_to_pending, args=(masters, entries))


def row_frame(rows):
    """Rows as display text under the on-screen headings."""
    return pd.DataFrame([{heading: core.fmt_cell(e.get(col)) if e.get(col) is not None else ""
                          for col, heading in TABLE_COLUMNS} for e in rows],
                        columns=[h for _c, h in TABLE_COLUMNS])


def row_styles(df, rows):
    today = date.today()
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    for pos, e in enumerate(rows):
        css = ""
        if e.get("Work Status") == "WIP":
            css = f"background-color: {theme.WIP_BG};"
        elif e.get("Work Status") == "Complete":
            css = f"background-color: {theme.COMPLETE_BG};"
        if core.is_overdue(e, today):
            css += f" color: {theme.OVERDUE_TEXT};"
        styles.iloc[pos, :] = css
    return styles


def tick_table(rows, key, ticked=False, height="content"):
    """Rows with a tick-box column; returns the positions of the ticked rows."""
    df = row_frame(rows)
    df.insert(0, "Select", ticked)
    styled = df.style.apply(lambda frame: row_styles(frame, rows), axis=None)
    edited = st.data_editor(
        styled, key=key, hide_index=True, width="stretch", height=height,
        disabled=[c for c in df.columns if c != "Select"],
        column_config={"Select": st.column_config.CheckboxColumn("Select", width="small"),
                       "#": st.column_config.Column(width="small")})
    return [pos for pos, value in enumerate(edited["Select"]) if value]


def build_pending():
    state = st.session_state
    pending = state.setdefault("pending", [])
    with st.container(border=True):
        st.subheader(f"Pending Submission: {len(pending)}")
        st.caption("Not saved until you click Submit.")
        if not pending:
            st.caption("No pending entries. Fill the form and click Add to Pending.")
            return
        ticked = tick_table(pending, f"pending_{len(pending)}")
        left, right, _rest = st.columns([1, 1, 3])
        if left.button("Remove Selected", disabled=not ticked, width="stretch"):
            state["pending"] = [e for i, e in enumerate(pending) if i not in ticked]
            st.rerun()
        right.button("✓ Submit to Scheduler", type="primary", width="stretch",
                     on_click=submit_pending)


@st.dialog("Delete from Scheduler")
def confirm_delete(keys, listing, extra):
    st.write(f"Permanently delete {plural(len(keys))}?")
    st.text("\n".join(listing) + (f"\n... and {extra} more" if extra else ""))
    if st.button("Delete", type="primary"):
        entries = core.load_schedule()
        kept = [e for e in entries if row_key(e) not in keys]
        save(kept, f"Deleted {plural(len(entries) - len(kept))}.")
        st.rerun()


def row_key(e):
    return tuple(core.fmt_cell(e.get(k)) for k in core.DUPLICATE_KEYS)


def build_update_bar(masters, rows):
    """Manager / Due Date / Target Date / Work Status editor for the ticked rows."""
    with st.container(border=True):
        if rows:
            st.caption(f"Updating {plural(len(rows))}: only the fields you fill in are changed.")
        else:
            st.caption("Tick rows below to set Manager, dates and Work Status, or to delete them.")
        # A new key after each update empties the fields, so they can't reach other rows.
        round_no = st.session_state.get("edit_round", 0)
        a, b, c, d = st.columns(4)
        manager = a.selectbox("Manager", [NO_CHANGE, NO_MANAGER, *masters.managers],
                              key=f"edit_manager_{round_no}")
        due = b.date_input("Due Date", value=None, key=f"edit_due_{round_no}",
                           format="DD/MM/YYYY")
        target = c.date_input("Target Date", value=None, key=f"edit_target_{round_no}",
                              format="DD/MM/YYYY")
        status = d.selectbox("Work Status", [NO_CHANGE, NOT_STARTED, *core.WORK_STATUSES],
                             key=f"edit_status_{round_no}")

        clear_dates = st.checkbox("Clear Due Date and Target Date", key=f"edit_clear_{round_no}")
        late = bool(due and target and target > due)
        allow = True
        if late:
            allow = st.checkbox(f"Target Date {target:%d-%b-%Y} is after Due Date "
                                f"{due:%d-%b-%Y}. Apply anyway", key=f"edit_late_{round_no}")
        label = f"Apply to Selected ({len(rows)})" if rows else "Apply to Selected"
        if not st.button(label, type="primary", disabled=not rows or not allow):
            return

        changes = {}
        if manager != NO_CHANGE:
            changes["Manager"] = None if manager == NO_MANAGER else manager
            changes["Manager Short Name"] = masters.manager_short.get(manager)
        if clear_dates:
            changes["Due Date"] = changes["Target Date"] = None
        if due:
            changes["Due Date"] = due
        if target:
            changes["Target Date"] = target
        if status != NO_CHANGE:
            changes["Work Status"] = None if status == NOT_STARTED else status
        if not changes:
            st.warning("Fill in at least one field to apply to the selected rows.")
            return
        keys = {row_key(e) for e in rows}
        entries = core.load_schedule()
        for e in entries:
            if row_key(e) in keys:
                e.update(changes)
        if save(entries, f"Updated {plural(len(rows))}: " + ", ".join(changes) + "."):
            st.session_state["edit_round"] = round_no + 1
        st.rerun()


def build_scheduler(masters, entries):
    state = st.session_state
    st.subheader("Scheduler")
    query = st.text_input("Search", key="query", placeholder="e.g. ATPL FY 2026-27 Monthly",
                          label_visibility="collapsed")
    found = core.search(entries, query)
    rows = [e for _i, e in found]
    for i, (_index, e) in enumerate(found, start=1):
        e["Sr No"] = e.get("Sr No") or i

    if not entries:
        st.info("Nothing submitted yet. Add entries to Pending, then click Submit to Scheduler.")
        return
    theme.legend(("WIP", theme.WIP_BG, theme.TEXT), ("Complete", theme.COMPLETE_BG, theme.TEXT),
                 ("Overdue", "#FFFFFF", theme.OVERDUE_TEXT))

    select_all = st.checkbox("Select all shown", key="select_all")
    # The key changes with the file, so ticks never carry over to rows that have moved.
    stamp = core.mtime(core.SCHEDULE_FILE)
    table = st.container()
    with table:
        ticked = tick_table(rows, f"table_{stamp}_{query}_{select_all}", select_all, 460)
    chosen = [rows[i] for i in ticked]
    st.caption(f"Showing {len(rows)} of {plural(len(entries))}  ·  {len(chosen)} selected")

    build_update_bar(masters, chosen)

    left, right = st.columns(2)
    with left:
        if st.button(f"Delete Selected ({len(chosen)})" if chosen else "Delete Selected",
                     disabled=not chosen):
            confirm_delete({row_key(e) for e in chosen},
                           [core.describe(e) for e in chosen[:8]], max(len(chosen) - 8, 0))
    with right.expander("Export to Excel"):
        fields = st.multiselect("Fields", core.SCHEDULE_COLUMNS, default=core.SCHEDULE_COLUMNS,
                                key="export_fields")
        export = [dict(e, **{"Sr No": i}) for i, e in enumerate(rows, start=1)]
        st.download_button(
            f"Download {plural(len(export))}", disabled=not fields,
            data=core.workbook_bytes(lambda out: core.write_schedule(out, export, fields))
            if fields else b"",
            file_name=f"Work Schedule Export {datetime.now():%d-%m-%Y %H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main():
    st.set_page_config(page_title="Work Scheduler", layout="wide")
    theme.header("Work Scheduler",
                 "Plan client assignments, work areas, frequencies and coverage periods")
    try:
        masters = read_masters(core.MASTER_FILE, core.mtime(core.MASTER_FILE))
    except FileNotFoundError:
        st.error(f"{core.MASTER_FILE.name} was not found next to the apps.")
        st.stop()
    except (KeyError, ValueError) as exc:
        st.error(f"Could not load the master workbook: {exc}")
        st.stop()
    try:
        entries = core.load_schedule()
    except PermissionError:
        st.error(f"Can't read {core.SCHEDULE_FILE.name}. Close it in Excel and refresh this page.")
        st.stop()
    core.sync_from_masters(entries, masters)

    a, b, c, d = st.columns(4)
    a.metric("Scheduled Entries", len(entries))
    b.metric("Clients Covered", len({e.get("Client Full Name") for e in entries}))
    c.metric("Work in Progress", sum(e.get("Work Status") == "WIP" for e in entries))
    d.metric("Completed", sum(e.get("Work Status") == "Complete" for e in entries))
    show_flash()

    form, main_area = st.columns([1, 3], gap="large")
    with form, st.container(border=True):
        build_form(masters, entries)
    with main_area:
        build_pending()
        build_scheduler(masters, entries)


main()
