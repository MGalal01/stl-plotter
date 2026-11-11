"""STL Plotter - turn two-column measurement text files into a formatted Excel workbook.

Start it with "Run STL Plotter.bat" (or:  python -m streamlit run app.py).
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import stl_core as core

APP_DIR = Path(__file__).parent
SETTINGS_FILE = APP_DIR / "settings.json"
DEFAULTS = {
    "file_fields": ["Noise type", "Mic position", "Mic number", "Notes"],
    "experiment_fields": ["Experiment title", "Date", "Operator", "Sample / specimen", "Notes"],
    "chart_title": "STL vs Frequency",
    "x_title": "Frequency (Hz)",
    "y_title": "STL (dB)",
    "chart_border_pt": 2.25,
    "plot_border_pt": 1.0,
}
PLOTLY_SYMBOLS = ["circle", "square", "triangle-up", "diamond", "x", "star", "cross"]
SCOPE_FILE = "Each file"
SCOPE_EXPERIMENT = "Whole experiment"
B_COLUMN = "Another column"
B_NUMBER = "A number"
RESERVED_FILE_FIELDS = {"file", "column name", "order"}


# Settings are only written to disk when started from "Run STL Plotter.bat" (which sets this).
# On a shared/hosted server every visitor would otherwise overwrite everyone else's settings.
PERSIST_SETTINGS = os.environ.get("STL_PLOTTER_SAVE_SETTINGS") == "1"


def load_settings() -> dict:
    if not PERSIST_SETTINGS:
        return dict(DEFAULTS)
    try:
        saved = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        return {**DEFAULTS, **{k: v for k, v in saved.items() if k in DEFAULTS}}
    except (OSError, ValueError):
        return dict(DEFAULTS)


def save_settings(settings: dict) -> None:
    if not PERSIST_SETTINGS:
        return
    try:
        SETTINGS_FILE.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    except OSError:
        pass  # read-only folder: the app still works, it just won't remember


def parse_field_list(text: str, reserved: set[str]) -> list[str]:
    out: list[str] = []
    for item in re.split(r"[\n,]+", text or ""):
        item = item.strip()
        if item and item.lower() not in reserved and item.lower() not in {o.lower() for o in out}:
            out.append(item)
    return out


# ------------------------------------------------------------------ button callbacks
def add_field() -> None:
    name = st.session_state.get("new_field_name", "").strip()
    if not name:
        st.session_state.field_message = ("warning", "Type a name for the new field first.")
        return
    per_file = st.session_state.get("new_field_scope") == SCOPE_FILE
    key = "file_fields_text" if per_file else "exp_fields_text"
    existing = parse_field_list(st.session_state[key], set())
    if name.lower() in {e.lower() for e in existing} or (per_file and name.lower() in RESERVED_FILE_FIELDS):
        st.session_state.field_message = ("warning", f"“{name}” already exists.")
        return
    st.session_state[key] = "\n".join(existing + [name])
    st.session_state.new_field_name = ""
    where = "the table above" if per_file else "Experiment details"
    st.session_state.field_message = ("success", f"Added “{name}” to {where}.")


def add_calculation() -> None:
    a = st.session_state.get("calc_a")
    op = st.session_state.get("calc_op")
    if st.session_state.get("calc_b_kind") == B_NUMBER:
        b = float(st.session_state.get("calc_b_number", 0.0))
    else:
        b = st.session_state.get("calc_b_column")
    if a is None or b is None:
        return
    st.session_state.calc_counter = st.session_state.get("calc_counter", 0) + 1
    st.session_state.calcs.append(
        {
            "id": st.session_state.calc_counter,
            "name": st.session_state.get("calc_name", "").strip(),
            "a": a,
            "op": op,
            "b": b,
        }
    )
    st.session_state.calc_name = ""


def remove_calculation(calc_id: int) -> None:
    st.session_state.calcs = [c for c in st.session_state.calcs if c["id"] != calc_id]


# ------------------------------------------------------------------ setup
st.set_page_config(page_title="STL Plotter", page_icon="📈", layout="wide")
if "settings" not in st.session_state:
    st.session_state.settings = load_settings()
    st.session_state.file_fields_text = "\n".join(st.session_state.settings["file_fields"])
    st.session_state.exp_fields_text = "\n".join(st.session_state.settings["experiment_fields"])
    st.session_state.calcs = []
settings = st.session_state.settings

# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.header("Settings")
    if PERSIST_SETTINGS:
        st.caption("These are remembered for next time.")
    else:
        st.caption("These apply to this session only.")

    st.subheader("Fields")
    st.text_area(
        "Fields per file (one per line)",
        key="file_fields_text",
        height=130,
        help="Each becomes a column you fill in for every uploaded file. Delete a line to remove a field.",
    )
    st.text_area(
        "Fields for the whole experiment (one per line)",
        key="exp_fields_text",
        height=130,
        help="Filled in once and written to the Info sheet. Delete a line to remove a field.",
    )

    st.subheader("Chart")
    chart_title = st.text_input("Chart title", settings["chart_title"])
    x_title = st.text_input("X axis title", settings["x_title"])
    y_title = st.text_input("Y axis title", settings["y_title"])
    log_x = st.checkbox("Logarithmic frequency axis", value=True)
    smooth = st.checkbox("Smooth lines", value=True)
    markers = st.checkbox("Markers", value=True)
    per_file_charts = st.checkbox("Also add one chart per column", value=False)
    chart_border_pt = st.number_input(
        "Chart area border (pt, 0 = none)", 0.0, 6.0, float(settings["chart_border_pt"]), 0.25
    )
    plot_border_pt = st.number_input(
        "Graph area border (pt, 0 = none)", 0.0, 6.0, float(settings["plot_border_pt"]), 0.25
    )

    st.subheader("Data")
    zeros_as_blank = st.checkbox(
        "Treat 0 as “no data”",
        value=True,
        help="Readings of exactly 0 are left blank in Excel and the curve is joined across them.",
    )

file_fields = parse_field_list(st.session_state.file_fields_text, reserved=RESERVED_FILE_FIELDS)
exp_fields = parse_field_list(st.session_state.exp_fields_text, reserved=set())
new_settings = {
    "file_fields": file_fields,
    "experiment_fields": exp_fields,
    "chart_title": chart_title,
    "x_title": x_title,
    "y_title": y_title,
    "chart_border_pt": chart_border_pt,
    "plot_border_pt": plot_border_pt,
}
if new_settings != settings:
    st.session_state.settings = new_settings
    save_settings(new_settings)

# ------------------------------------------------------------------ 1. upload
st.title("STL Plotter")
st.caption("Text files in → formatted Excel workbook with a chart out.")

st.subheader("1 · Upload measurement files")
uploads = st.file_uploader(
    "Two-column text files (frequency, value)",
    type=["txt", "csv", "dat", "tsv", "asc"],
    accept_multiple_files=True,
)
if not uploads:
    st.info("Drop one or more text files above to begin.")
    st.stop()

measurements: dict[str, core.Measurement] = {}
for up in uploads:
    label, n = up.name, 2
    while label in measurements:  # same file name uploaded twice
        label = f"{up.name} [{n}]"
        n += 1
    m = core.parse_measurement(up.name, up.getvalue())
    if not m.freq:
        st.error(f"**{up.name}**: no rows with two numbers were found, so it was left out.")
        continue
    measurements[label] = m
if not measurements:
    st.stop()

# ------------------------------------------------------------------ 2. name columns
st.subheader("2 · Name the columns and describe each file")
st.caption(
    "Names and fields are pre-filled from the file names where possible. "
    "Double-click a cell to edit. “Order” sets the column and legend order."
)

labels = list(measurements)
signature = (tuple(labels), tuple(file_fields))
if st.session_state.get("signature") != signature:
    saved = st.session_state.get("saved_rows", {})
    rows = []
    for i, label in enumerate(labels, start=1):
        old = saved.get(label, {})
        guesses = core.guess_fields(label)
        row = {
            "File": label,
            "Column name": old.get("Column name") or core.default_name(measurements[label].filename),
            "Order": old.get("Order") or i,
        }
        for f in file_fields:
            previous = old.get(f)
            row[f] = previous if isinstance(previous, str) and previous else guesses.get(f, "")
        rows.append(row)
    st.session_state.base_table = pd.DataFrame(rows)
    st.session_state.signature = signature
    st.session_state.editor_version = st.session_state.get("editor_version", 0) + 1

column_config = {
    "File": st.column_config.TextColumn("File", disabled=True),
    "Column name": st.column_config.TextColumn("Column name", required=True, width="medium"),
    "Order": st.column_config.NumberColumn("Order", min_value=1, step=1, width="small"),
}
column_config.update({f: st.column_config.TextColumn(f) for f in file_fields})
edited = st.data_editor(
    st.session_state.base_table,
    key=f"editor_{st.session_state.editor_version}",
    column_config=column_config,
    hide_index=True,
    num_rows="fixed",
    use_container_width=True,
)
st.session_state.saved_rows = {r["File"]: r for r in edited.to_dict("records")}

with st.container(border=True):
    st.markdown("**Add another field** (any other condition you want to record)")
    f1, f2, f3 = st.columns([3, 2, 1.4], vertical_alignment="bottom")
    f1.text_input("Field name", key="new_field_name", placeholder="e.g. Gain (dB), Distance (m), Temperature")
    f2.selectbox("Applies to", [SCOPE_FILE, SCOPE_EXPERIMENT], key="new_field_scope")
    f3.button("➕ Add field", on_click=add_field, use_container_width=True)
    message = st.session_state.pop("field_message", None)
    if message:
        getattr(st, message[0])(message[1])
    st.caption("To remove or reorder fields, edit the lists in the sidebar.")

edited = edited.sort_values("Order", kind="stable", na_position="last")
ordered_labels = edited["File"].tolist()
ordered = [measurements[l] for l in ordered_labels]
names = core.unique_names(
    edited["Column name"].tolist(), [core.default_name(m.filename) for m in ordered]
)
if [str(n).strip() for n in edited["Column name"].tolist()] != names:
    st.warning("Some column names were empty or repeated, so they were adjusted: " + ", ".join(names))

df = core.combine(ordered, names, zeros_as_blank=zeros_as_blank)

# ------------------------------------------------------------------ 3. calculated columns
# Columns are referenced by a stable key, so renaming a column does not break a calculation.
key_to_name: dict[str, str] = {f"file:{l}": n for l, n in zip(ordered_labels, names)}
used_names = {n.lower() for n in names}
calc_defs: list[tuple[str, str]] = []   # (column name, definition text)
calc_rows: list[tuple[int, str, str]] = []  # (id, column name, definition text)
broken: list[dict] = []
for c in st.session_state.calcs:
    a_name = key_to_name.get(c["a"])
    b_value = c["b"] if isinstance(c["b"], float) else key_to_name.get(c["b"])
    if a_name is None or b_value is None:
        broken.append(c)
        continue
    b_text = core.format_number(b_value) if isinstance(b_value, float) else b_value
    definition = f"{a_name} {c['op']} {b_text}"
    name = core.make_unique(c["name"] or definition, used_names)
    df[name] = core.calculate(df, a_name, c["op"], b_value)
    key_to_name[f"calc:{c['id']}"] = name
    calc_defs.append((name, definition))
    calc_rows.append((c["id"], name, definition))

st.subheader("3 · Calculated columns (optional)")
st.caption(
    "Add, subtract, multiply or divide a column by another column or by a number. "
    "A calculated column can be used in further calculations."
)
for calc_id, name, definition in calc_rows:
    r1, r2 = st.columns([6, 1], vertical_alignment="center")
    r1.markdown(f"**{name}**  =  {definition}")
    r2.button("Remove", key=f"remove_calc_{calc_id}", on_click=remove_calculation, args=(calc_id,),
              use_container_width=True)
for c in broken:
    r1, r2 = st.columns([6, 1], vertical_alignment="center")
    r1.warning(f"“{c['name'] or 'A calculation'}” uses a column that is no longer available, so it was skipped.")
    r2.button("Remove", key=f"remove_calc_{c['id']}", on_click=remove_calculation, args=(c["id"],),
              use_container_width=True)

with st.container(border=True):
    keys = list(key_to_name)
    c1, c2, c3, c4 = st.columns([3, 1.2, 1.8, 3], vertical_alignment="bottom")
    c1.selectbox("Column", keys, format_func=lambda k: key_to_name.get(k, k), key="calc_a")
    c2.selectbox("Operation", core.OPERATIONS, key="calc_op")
    b_kind = c3.selectbox("With", [B_COLUMN, B_NUMBER], key="calc_b_kind")
    if b_kind == B_NUMBER:
        c4.number_input("Number", value=1.0, format="%g", key="calc_b_number")
    else:
        c4.selectbox("Second column", keys, index=min(1, len(keys) - 1),
                     format_func=lambda k: key_to_name.get(k, k), key="calc_b_column")
    n1, n2 = st.columns([5, 1.6], vertical_alignment="bottom")
    n1.text_input("Name of the new column", key="calc_name", placeholder="Leave empty to name it after the formula")
    n2.button("➕ Add calculation", on_click=add_calculation, use_container_width=True)

# ------------------------------------------------------------------ 4. experiment details
st.subheader("4 · Experiment details")
experiment: dict[str, str] = {}
if exp_fields:
    cols = st.columns(min(3, len(exp_fields)))
    for i, f in enumerate(exp_fields):
        with cols[i % len(cols)]:
            if f.strip().lower() == "date":
                experiment[f] = st.date_input(f, value=dt.date.today(), format="YYYY-MM-DD").isoformat()
            elif f.strip().lower() in ("notes", "comments", "remarks"):
                experiment[f] = st.text_area(f, height=68, key=f"exp_{f}")
            else:
                experiment[f] = st.text_input(f, key=f"exp_{f}")
else:
    st.caption("No experiment fields are defined. Add one above or in the sidebar.")

# ------------------------------------------------------------------ 5. preview
st.subheader("5 · Preview")
all_columns = list(df.columns)
plot_columns = st.multiselect(
    "Columns to show on the chart",
    all_columns,
    default=all_columns,
    help="All columns are always written to the Data sheet. This only chooses which ones are drawn.",
)

notes = []
for m, name in zip(ordered, names):
    bits = [f"{len(m.freq)} points"]
    if zeros_as_blank and m.zero_count:
        bits.append(f"{m.zero_count} zero readings left blank")
    if m.skipped_lines:
        bits.append(f"{m.skipped_lines} non-numeric lines skipped")
    if m.duplicate_rows:
        bits.append(f"{m.duplicate_rows} repeated frequencies dropped")
    notes.append(f"**{name}**: " + ", ".join(bits))
st.caption("  \n".join(notes))
if len({tuple(m.freq) for m in ordered}) > 1:
    st.info(
        "The files do not all share the same frequency points. They were merged into one "
        "frequency column, with blanks where a file has no reading."
    )
if len(all_columns) > len(core.SERIES_COLORS):
    st.info(
        f"More than {len(core.SERIES_COLORS)} columns: colours repeat after the eighth, "
        "so later curves are told apart by marker shape."
    )

if plot_columns:
    fig = go.Figure()
    n_colors = len(core.SERIES_COLORS)
    for i, name in enumerate(all_columns):  # colour follows the column, as in the Excel chart
        if name not in plot_columns:
            continue
        color = "#" + core.SERIES_COLORS[i % n_colors]
        fig.add_trace(
            go.Scatter(
                x=df.index,
                y=df[name],
                name=name,
                mode="lines+markers" if markers else "lines",
                connectgaps=True,
                line=dict(color=color, width=2, shape="spline" if smooth else "linear"),
                marker=dict(color=color, size=6, symbol=PLOTLY_SYMBOLS[(i + i // n_colors) % len(PLOTLY_SYMBOLS)]),
                hovertemplate="%{x:g} Hz<br>%{y:.2f}<extra>" + name + "</extra>",
            )
        )
    fig.update_layout(
        title=chart_title or None,
        template="plotly_white",
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(color="#0b0b0b"),
        height=480,
        margin=dict(l=60, r=20, t=60 if chart_title else 20, b=60),
        legend=dict(orientation="h", yanchor="top", y=-0.2, x=0),
        showlegend=True,
    )
    if plot_border_pt > 0:  # same frame as the Excel graph area
        axis_style = dict(showline=True, mirror=True, linecolor="#000000",
                          linewidth=max(1, round(plot_border_pt * 1.33)), gridcolor="#e1e0d9", zeroline=False)
    else:
        axis_style = dict(showline=True, linecolor="#c3c2b7", gridcolor="#e1e0d9", zeroline=False)
    fig.update_xaxes(title_text=x_title, type="log" if log_x else "linear", **axis_style)
    fig.update_yaxes(title_text=y_title, **axis_style)
    st.plotly_chart(fig, use_container_width=True, theme=None)
else:
    st.warning("No columns are selected for the chart, so the Excel file will contain the data without a chart.")

with st.expander("Show the data table"):
    st.dataframe(df, use_container_width=True)

# ------------------------------------------------------------------ 6. download
st.subheader("6 · Download")
infos = []
for label, m, name in zip(ordered_labels, ordered, names):
    row = st.session_state.saved_rows[label]
    fields = {f: "" if pd.isna(row.get(f)) else str(row.get(f, "")) for f in file_fields}
    infos.append(
        core.FileInfo(
            column_name=name,
            filename=m.filename,
            fields=fields,
            points=len(m.freq),
            blanked_zeros=m.zero_count if zeros_as_blank else 0,
        )
    )
opts = core.ChartOptions(
    title=chart_title,
    x_title=x_title,
    y_title=y_title,
    log_x=log_x,
    smooth=smooth,
    markers=markers,
    per_file_charts=per_file_charts,
    chart_border_pt=chart_border_pt,
    plot_border_pt=plot_border_pt,
)
workbook = core.build_workbook(df, infos, experiment, opts, plot_columns=plot_columns, calculations=calc_defs)

title_value = next((v for k, v in experiment.items() if "title" in k.lower() and v.strip()), "")
suggested = re.sub(r'[\\/:*?"<>|]+', " ", title_value).strip() or f"STL results {dt.date.today().isoformat()}"
d1, d2 = st.columns([3, 1], vertical_alignment="bottom")
with d1:
    out_name = st.text_input("Excel file name", value=suggested)
with d2:
    safe = re.sub(r'[\\/:*?"<>|]+', " ", out_name).strip() or suggested
    st.download_button(
        "Download Excel file",
        data=workbook,
        file_name=safe if safe.lower().endswith(".xlsx") else safe + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
