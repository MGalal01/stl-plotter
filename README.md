# STL Plotter

Turns two-column measurement text files (frequency, STL) into a formatted Excel workbook with a
scatter chart (smooth lines + markers), an experiment-information sheet, and optional calculated
columns.

## Use it

1. Upload one or more text files (tab-, space-, comma- or semicolon-separated: `frequency  value`).
2. Rename the columns and fill in the per-file fields (mic position, noise type, ...).
   Use **Add field** for any other condition you want to record.
3. Optionally add calculated columns (column − column, column × 2.5, ...).
4. Fill in the experiment details, check the preview, and download the Excel file.

Readings of exactly 0 are treated as "no data" by default (left blank, the curve joins across
them). Untick that in the sidebar if zeros are real values.

## Run it on your own computer (Windows)

1. Install Python 3.10 or newer from python.org (tick "Add Python to PATH").
2. Double-click `Install requirements.bat` (once).
3. Double-click `Run STL Plotter.bat`. The app opens in your browser.

Or from a terminal: `pip install -r requirements.txt` then `python -m streamlit run app.py`.

## Files

- `app.py` – the web interface
- `stl_core.py` – file parsing, calculations and the Excel builder
