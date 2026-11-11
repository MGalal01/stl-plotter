"""Parsing of two-column measurement text files and export to a formatted Excel workbook.

Kept free of any Streamlit imports so it can be tested or reused from a plain script.
"""
from __future__ import annotations

import io
import math
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Categorical series colours (fixed order) and marker shapes. The shape changes with
# every series as well, so curves stay distinguishable in black-and-white print.
SERIES_COLORS = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300", "6250D6", "E34948"]
MARKER_SHAPES = ["circle", "square", "triangle", "diamond", "x", "star", "plus"]

_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_CALC_FILL = PatternFill("solid", fgColor="375623")   # header of calculated columns
_LABEL_FILL = PatternFill("solid", fgColor="DDEBF7")
_SIDE = Side(style="thin", color="595959")
_BORDER = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


# --------------------------------------------------------------------------- parsing

@dataclass
class Measurement:
    filename: str
    freq: list[float]
    value: list[float]
    skipped_lines: int = 0      # lines that were not two numbers (headers, comments)
    duplicate_rows: int = 0     # repeated frequencies (first one kept)

    @property
    def zero_count(self) -> int:
        return sum(1 for v in self.value if v == 0)


def _decode(raw: bytes) -> str:
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16", errors="replace")
    for enc in ("utf-8-sig", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")


def _split_numbers(line: str) -> list[str]:
    parts = [p for p in re.split(r"[\t; ]+", line) if p]
    if len(parts) >= 2:
        # separated by tab/space/semicolon, so a comma can only be a decimal comma
        return [p.replace(",", ".") for p in parts]
    return [p for p in line.split(",") if p.strip()]


def parse_measurement(filename: str, raw: bytes) -> Measurement:
    """Read 'frequency  value' rows. Lines that are not two numbers are skipped."""
    freq: list[float] = []
    value: list[float] = []
    seen: set[float] = set()
    skipped = duplicates = 0
    for line in _decode(raw).splitlines():
        line = line.strip()
        if not line:
            continue
        parts = _split_numbers(line)
        try:
            f, v = float(parts[0]), float(parts[1])
        except (ValueError, IndexError):
            skipped += 1
            continue
        if not (math.isfinite(f) and math.isfinite(v)):
            skipped += 1
            continue
        f = float(f"{f:.9g}")  # tidy float noise so identical frequencies line up across files
        if f in seen:
            duplicates += 1
            continue
        seen.add(f)
        freq.append(f)
        value.append(v)
    return Measurement(filename, freq, value, skipped, duplicates)


def default_name(filename: str) -> str:
    """Readable starting name for a column, derived from the file name."""
    stem = Path(filename).stem
    return re.sub(r"[_\s]+", " ", stem).strip() or stem


def guess_fields(filename: str) -> dict[str, str]:
    """Best-effort guesses for common fields from the file name. All are editable."""
    stem = Path(filename).stem
    out: dict[str, str] = {}
    m = re.search(r"(white|pink|brown|red|blue)[\s_]*noise", stem, re.I)
    if m:
        out["Noise type"] = f"{m.group(1).capitalize()} noise"
    m = re.search(r"(\d{1,2})[\s_]*(?:o'?)?clock", stem, re.I)
    if m:
        out["Mic position"] = f"{m.group(1)} o'clock"
    m = re.search(r"mic[\s_]*(\d+)", stem, re.I)
    if m:
        out["Mic number"] = m.group(1)
    return out


def unique_names(names: list[str], fallbacks: list[str]) -> list[str]:
    """Blank names fall back to the file name; repeats get ' (2)', ' (3)' appended."""
    out: list[str] = []
    used: set[str] = set()
    for name, fb in zip(names, fallbacks):
        base = str(name).strip() if name is not None and str(name).strip() not in ("", "nan", "None") else fb
        candidate, n = base, 2
        while candidate.lower() in used:
            candidate = f"{base} ({n})"
            n += 1
        used.add(candidate.lower())
        out.append(candidate)
    return out


def combine(measurements: list[Measurement], names: list[str], zeros_as_blank: bool = True) -> pd.DataFrame:
    """One row per frequency, one column per file. Missing points are NaN."""
    columns = {}
    for m, name in zip(measurements, names):
        s = pd.Series(m.value, index=m.freq, dtype="float64")
        if zeros_as_blank:
            s = s.where(s != 0)
        columns[name] = s
    df = pd.concat(columns, axis=1).sort_index()
    df.index.name = "Frequency"
    return df


# --------------------------------------------------------------------------- maths

OPERATIONS = ["−", "+", "×", "÷"]


def make_unique(name: str, used: set[str]) -> str:
    """Return name, or 'name (2)' etc. if it is already taken. Adds the result to used."""
    base = name.strip() or "Calculated"
    candidate, n = base, 2
    while candidate.lower() in used:
        candidate = f"{base} ({n})"
        n += 1
    used.add(candidate.lower())
    return candidate


def calculate(df: pd.DataFrame, a: str, op: str, b: str | float) -> pd.Series:
    """Column a (op) column b, or column a (op) a number. Blanks stay blank."""
    left = df[a]
    right = df[b] if isinstance(b, str) else float(b)
    if op == "+":
        out = left + right
    elif op == "−":
        out = left - right
    elif op == "×":
        out = left * right
    elif op == "÷":
        out = left / right
    else:
        raise ValueError(f"Unknown operation: {op}")
    return out.replace([math.inf, -math.inf], float("nan"))


def format_number(x: float) -> str:
    return f"{x:g}"


# --------------------------------------------------------------------------- excel

@dataclass
class ChartOptions:
    title: str = "Sound Transmission Loss"
    x_title: str = "Frequency (Hz)"
    y_title: str = "STL (dB)"
    log_x: bool = True
    smooth: bool = True
    markers: bool = True
    marker_size: int = 5
    per_file_charts: bool = False
    chart_border_pt: float = 2.25   # outer chart area border, 0 = none
    plot_border_pt: float = 1.0     # border around the plot (graph) area, 0 = none


@dataclass
class FileInfo:
    column_name: str
    filename: str
    fields: dict[str, str] = field(default_factory=dict)
    points: int = 0
    blanked_zeros: int = 0


def _put(ws, row: int, col: int, value, *, font=None, fill=None, align=_CENTER, fmt=None):
    cell = ws.cell(row=row, column=col)
    if isinstance(value, float) and math.isnan(value):
        value = None
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"  # keep user text as text, never as a formula
    cell.border = _BORDER
    cell.alignment = align
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    return cell


def _text_width(values, minimum=10, maximum=60) -> float:
    longest = max((len(str(v)) for v in values if v is not None), default=0)
    return max(minimum, min(maximum, longest + 3))


def _line(color: str, width_pt: float) -> LineProperties:
    return LineProperties(solidFill=color, w=int(width_pt * 12700))


def _build_chart(ws_data, n_rows: int, series_cols: list[int], opts: ChartOptions,
                 title: str, x_bounds: tuple[float, float] | None, show_legend: bool = True) -> ScatterChart:
    chart = ScatterChart()
    chart.title = title or None
    chart.style = 2
    chart.scatterStyle = "smoothMarker" if opts.smooth else "lineMarker"
    chart.height, chart.width = 12.5, 24
    chart.display_blanks = "span"  # join the curve across blank (no-data) cells
    chart.varyColors = False       # otherwise a single series gets one legend entry per point
    chart.roundedCorners = False
    chart.graphical_properties = GraphicalProperties(
        solidFill="FFFFFF",
        ln=_line("000000", opts.chart_border_pt) if opts.chart_border_pt > 0 else LineProperties(noFill=True),
    )
    if opts.plot_border_pt > 0:
        chart.plot_area.graphicalProperties = GraphicalProperties(noFill=True, ln=_line("000000", opts.plot_border_pt))

    chart.x_axis.title = opts.x_title
    chart.y_axis.title = opts.y_title
    # without overlay=False Excel draws titles and legend on top of the plot area
    for titled in (chart, chart.x_axis, chart.y_axis):
        if titled.title is not None:
            titled.title.overlay = False
    chart.x_axis.delete = chart.y_axis.delete = False
    chart.x_axis.number_format = "0"
    chart.y_axis.number_format = "General"
    chart.x_axis.majorTickMark = chart.y_axis.majorTickMark = "out"
    chart.x_axis.crosses = "min"   # x axis sits at the bottom, even with negative values
    chart.y_axis.crosses = "min"   # y axis sits at the left edge
    grid = GraphicalProperties(ln=_line("D9D9D9", 0.75))
    chart.x_axis.majorGridlines = ChartLines(spPr=grid)
    chart.y_axis.majorGridlines = ChartLines(spPr=grid)
    if x_bounds:
        lo, hi = x_bounds
        if opts.log_x and lo > 0:
            chart.x_axis.scaling.logBase = 10
            chart.x_axis.minorGridlines = ChartLines(spPr=GraphicalProperties(ln=_line("EDEDED", 0.5)))
            chart.x_axis.scaling.min, chart.x_axis.scaling.max = lo, hi
        elif not opts.log_x:
            chart.x_axis.scaling.max = hi

    x_ref = Reference(ws_data, min_col=1, min_row=2, max_row=n_rows + 1)
    for col in series_cols:
        i = col - 2  # colour/shape follow the column, so a per-file chart matches the combined one
        color = SERIES_COLORS[i % len(SERIES_COLORS)]
        shape = MARKER_SHAPES[(i + i // len(SERIES_COLORS)) % len(MARKER_SHAPES)]
        y_ref = Reference(ws_data, min_col=col, min_row=1, max_row=n_rows + 1)
        s = Series(y_ref, x_ref, title_from_data=True)
        s.smooth = opts.smooth
        s.graphicalProperties.line = _line(color, 1.5)
        if opts.markers:
            s.marker.symbol = shape
            s.marker.size = opts.marker_size
            s.marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=_line(color, 0.75))
        else:
            s.marker.symbol = "none"
        chart.series.append(s)

    if show_legend:
        chart.legend.position = "b"
        chart.legend.overlay = False
    else:
        chart.legend = None
    return chart


def build_workbook(df: pd.DataFrame, files: list[FileInfo], experiment: dict[str, str],
                   opts: ChartOptions | None = None, plot_columns: list[str] | None = None,
                   calculations: list[tuple[str, str]] | None = None) -> bytes:
    """Create the .xlsx (Data sheet with table + chart, optional Charts sheet, Info sheet).

    plot_columns: which columns to draw (default: all).
    calculations: (column name, definition text) for columns that were calculated, for the Info sheet.
    """
    opts = opts or ChartOptions()
    calculations = calculations or []
    calculated = {name for name, _ in calculations}
    wb = Workbook()

    # ---- Data sheet
    ws = wb.active
    ws.title = "Data"
    headers = [opts.x_title or "Frequency"] + [str(c) for c in df.columns]
    for c, h in enumerate(headers, start=1):
        _put(ws, 1, c, h, font=_HEADER_FONT, fill=_CALC_FILL if h in calculated and c > 1 else _HEADER_FILL)
    for r, (freq, row) in enumerate(df.iterrows(), start=2):
        _put(ws, r, 1, float(freq), fmt="0.0##", fill=_LABEL_FILL)
        for c, v in enumerate(row.tolist(), start=2):
            _put(ws, r, c, v, fmt="0.00")
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = _text_width([headers[0]], minimum=14, maximum=24)
    for c, h in enumerate(headers[1:], start=2):
        ws.column_dimensions[get_column_letter(c)].width = _text_width([h], minimum=14, maximum=32)

    n_rows, n_series = len(df), len(df.columns)
    freqs = [float(f) for f in df.index]
    x_bounds = (min(freqs), max(freqs)) if freqs else None
    series_cols = [i + 2 for i, c in enumerate(df.columns) if plot_columns is None or c in plot_columns]
    if n_rows and series_cols:
        chart = _build_chart(ws, n_rows, series_cols, opts, opts.title, x_bounds, show_legend=True)
        ws.add_chart(chart, f"{get_column_letter(n_series + 3)}2")

        # ---- optional one chart per file
        if opts.per_file_charts and len(series_cols) > 1:
            wc = wb.create_sheet("Charts")
            for i, col in enumerate(series_cols):
                single = _build_chart(ws, n_rows, [col], opts, str(df.columns[col - 2]), x_bounds, show_legend=False)
                single.height, single.width = 9, 16
                wc.add_chart(single, f"{'B' if i % 2 == 0 else 'L'}{2 + (i // 2) * 20}")

    # ---- Info sheet
    wi = wb.create_sheet("Info")
    _put(wi, 1, 1, "Experiment information", font=_HEADER_FONT, fill=_HEADER_FILL)
    _put(wi, 1, 2, None, fill=_HEADER_FILL)
    wi.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    r = 2
    for key, val in experiment.items():
        _put(wi, r, 1, key, font=Font(bold=True), fill=_LABEL_FILL, align=_LEFT)
        _put(wi, r, 2, val, align=_LEFT)
        r += 1

    r += 1
    field_names: list[str] = []
    for f in files:
        for k in f.fields:
            if k not in field_names:
                field_names.append(k)
    cols = ["Column name", "Source file"] + field_names + ["Data points", "Zero readings left blank"]
    for c, h in enumerate(cols, start=1):
        _put(wi, r, c, h, font=_HEADER_FONT, fill=_HEADER_FILL)
    wi.row_dimensions[r].height = 32
    table_values: list[list] = [cols]
    for f in files:
        r += 1
        row = [f.column_name, f.filename] + [f.fields.get(k, "") for k in field_names] + [f.points, f.blanked_zeros]
        table_values.append(row)
        for c, v in enumerate(row, start=1):
            _put(wi, r, c, v, align=_LEFT if c <= 2 else _CENTER)
    if calculations:
        r += 2
        for c, h in enumerate(["Calculated column", "Definition"], start=1):
            _put(wi, r, c, h, font=_HEADER_FONT, fill=_CALC_FILL)
        for name, definition in calculations:
            r += 1
            _put(wi, r, 1, name, align=_LEFT)
            _put(wi, r, 2, definition, align=_LEFT)
            table_values.append([name, definition] + [None] * (len(cols) - 2))
    for c in range(1, len(cols) + 1):
        col_values = [row[c - 1] for row in table_values]
        if c == 1:
            col_values += list(experiment.keys())
        elif c == 2:
            col_values += [str(v)[:60] for v in experiment.values()]
        wi.column_dimensions[get_column_letter(c)].width = _text_width(col_values, minimum=14, maximum=55)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
