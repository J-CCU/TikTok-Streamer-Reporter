# TikTok Streamer Reporter

A desktop tool that turns raw TikTok Shop LIVE performance exports (CSV or XLSX) into an Excel report of attributed GMV and streaming hours per streamer, by day, week and month.

The workflow is review-first. The app imports and deduplicates every live room, prefills the streamer only when the room title gives clear evidence, and lets you assign or correct the rest in a table. You can't generate the report until every stream has a streamer. It is meant for TikTok Shop sellers who need per-host sales and hours totals from their LIVE exports.

## Quick start

1. Install Python 3.8+ with Tkinter, then `python3 -m pip install XlsxWriter` (Windows: `py -m pip install XlsxWriter`).
2. Run `python3 tiktok_streamer_reporter.py` (Windows: `py tiktok_streamer_reporter.py`).
3. Click **Add Exports...**, choose your TikTok LIVE `.csv` / `.xlsx` files, and fill in any empty **Streamer** cells.
4. Click **Generate Excel Report** and choose where to save `TikTok_Streamer_Report.xlsx`.

On an Apple Silicon Mac you can use the prebuilt app instead (see [Prebuilt app](#prebuilt-app-macos-apple-silicon)).

## Features

- Imports one or more TikTok LIVE exports in `.csv` or `.xlsx` format. You can mix both formats in one run.
- Finds the header row automatically within the first 100 rows of each file.
- Deduplicates by Room ID. When a room appears in more than one export, the newest export wins. If the exports are equally recent, the row with more filled-in fields wins.
- Prefills streamer names from title evidence only, in this order:
  - a built-in alias (the `KNOWN_ALIASES` list in the source, which maps nicknames to display names). A title is matched only when exactly one known name appears in it.
  - an explicit host marker followed by the name at the end of the title: `W/ <name>` (also `W.` or `W-`), `WITH <name>`, `HOSTED BY <name>`, `HOST: <name>` (also `HOST-`). The text after the marker is ignored if it contains a weekday or `?`, has more than three words, or starts with a word such as `sale`, `live` or `breaks`.
  - names learned from those markers and reused in other titles
  - the narrow phrase `<Name> on a <weekday>` at the start of the title
- Never guesses from schedules, time slots, or neighboring streams. Days with no streams are allowed and don't cause warnings.
- Spreadsheet-style review table with filters (All rows, Needs streamer, Assigned), a per-file source filter and a text search.
- Editing tools: double-click to edit one cell, bulk-assign selected rows, clear selected rows, copy the streamer from the previous or next stream in the same file, and Rename / Merge a name across all rows.
- Excel output has one calendar-style sheet per month with daily and weekly GMV and hours. GMV cells are color-banded. Durations use `[hh]:mm` format, so totals over 24 hours display correctly.
- Command-line mode for batch exports. It uses automatic title detection only.

## Requirements

- Python 3.8 or newer, with Tkinter (included with the python.org installers)
- [XlsxWriter](https://pypi.org/project/XlsxWriter/), needed to write the report

Reading input needs only the Python standard library. Pandas and openpyxl aren't required.

```bash
# Windows
py -m pip install XlsxWriter

# macOS
python3 -m pip install XlsxWriter
```

## How to run

### Prebuilt app (macOS, Apple Silicon)

1. Unzip `TikTok Streamer Reporter.app.zip`.
2. Open `TikTok Streamer Reporter.app`. If macOS blocks the unsigned app, right-click it and choose **Open**.

`TikTok Streamer Reporter` (no extension) is the same program as a standalone arm64 Mac executable. You can run it from Terminal.

The prebuilt binaries are for Apple Silicon Macs only. On Windows, Intel Macs or Linux, run from source (see the next section).

### From source

```bash
# Windows
py tiktok_streamer_reporter.py

# macOS
python3 tiktok_streamer_reporter.py
```

If you don't give any file arguments, the GUI opens.

### GUI workflow

1. Click **Add Exports...** and pick one or more TikTok `.csv` / `.xlsx` files. The app imports them right away. Use **Remove Selected** or **Clear** to change the file list. Adding or removing a file re-imports the whole list; assignments you made in this session are kept, because they are stored by Room ID.
2. Read the summary line: number of files, unique rows, duplicates removed, rows assigned and rows still needing a streamer.
3. Click **Show Needs Streamer**, or set **Show** to *Needs streamer*, to see the rows that still need a name.
4. Assign streamers:
   - Double-click a **Streamer** cell, or select a row and press Return. Type or pick a name, then press Return or Tab to save it, or Esc to cancel.
   - To assign several rows, select them, type or pick a name in **Assign selected to:**, then click **Apply to selected**.
   - **Use previous** / **Use next** copies the streamer from the neighboring stream in the same source file. Select exactly one row first.
   - **Rename / Merge Name...** replaces every assignment of one name with another. Use an existing name to merge two names into one.
   - **Clear selected** removes the assignments from the selected rows.
5. When every row is assigned, click **Generate Excel Report** and choose where to save it. The default file name is `TikTok_Streamer_Report.xlsx`.

Press Cmd+F (macOS) or Ctrl+F to jump to the search box.

### Command line

```bash
# Windows
py tiktok_streamer_reporter.py export1.csv export2.xlsx -o TikTok_Streamer_Report.xlsx

# macOS
python3 tiktok_streamer_reporter.py export1.csv export2.xlsx -o TikTok_Streamer_Report.xlsx
```

- `-o / --output` sets the output path. Without it, the report is saved as `TikTok_Streamer_Report.xlsx` in the current folder.
- The CLI uses automatic title detection only. If any stream is still unassigned, the CLI doesn't write a report. It lists up to 20 unassigned streams instead. To finish those rows, use the GUI.
- Exit codes:

  | Code | Meaning |
  |---|---|
  | `0` | Success |
  | `2` | Missing file or import error |
  | `3` | Unassigned streams |
  | `4` | Excel write error |

## Inputs

Raw TikTok Shop LIVE export files, `.csv` or `.xlsx`.

- **Required headers:** `Room ID`, `Room Title`, `Start Time`. Matching ignores case and extra spaces.
- **Optional headers:**
  - `End Time`
  - `Duration`
  - `Attributed GMV`. If this column is missing, `GMV` is used.
- **CSV:** UTF-8 (with or without a BOM) or cp1252. The delimiter can be a comma, tab or semicolon and is detected automatically.
- **XLSX:** the app reads each sheet in order and uses the first one that contains the required headers.
- **Date/time formats:**
  - `YYYY-MM-DD HH:MM[:SS]`
  - `YYYY/MM/DD HH:MM:SS`
  - `MM/DD/YYYY HH:MM[:SS]`
  - `MM/DD/YYYY hh:mm[:ss] AM/PM`
  - ISO 8601
  - Excel serial dates
- **Duration formats:** `1h 23m 4s`, `HH:MM`, `HH:MM:SS`, or an Excel day fraction. If the duration is missing or can't be read, the app uses End Time minus Start Time.
- **GMV:** `$` signs and commas are removed. A value in parentheses is read as a negative number. A blank or unreadable value counts as 0.
- **Skipped rows:** rows without a readable Start Time are skipped silently. If a row has no Room ID, the app creates one from the file name and row number.
- **Freshness:** if a file name contains a 14-digit `YYYYMMDDHHMMSS` timestamp starting with `20`, that timestamp sets the export's freshness for deduplication. Otherwise the file's modified time is used.

## Output

The output is a single `.xlsx` workbook, saved to the path you choose (GUI) or the path given with `-o` (CLI).

| Sheet | Contents |
|---|---|
| One per month (for example `January`, or `January 2026` when the data spans more than one year) | A calendar grid with weeks running Monday to Sunday. Each streamer in the report has two columns: GMV and `Hrs`. Each week has a `Week N` row with weekly totals, then one row per day. The first and last weeks are full Monday-to-Sunday weeks, so they show days from the neighboring month. Those days are greyed out with blank GMV and hours, and they aren't counted in that sheet's weekly totals, so each month sheet only totals its own month. |
| `Combined Data` | One row per deduplicated stream, with an autofilter. |
| `Run Info` | Row count after deduplication, duplicates removed, the duplicate rule, the list of streamers, the hours format and the attribution policy. |

**Month sheets.** In the daily rows, GMV cells are color-filled by amount:

| Daily GMV | Fill color |
|---|---|
| under $1,000 | red |
| $1,000 to $1,999.99 | yellow |
| $2,000 to $3,999.99 | light green |
| $4,000 to $9,999.99 | green |
| $10,000 and above | cyan |

Streamer columns follow the fixed order in `PREFERRED_ORDER` in the source. Any other names follow in alphabetical order. Only streamers with at least one assigned stream get columns.

**Combined Data columns.**

| Column | Contents |
|---|---|
| Source File | Name of the export the row came from |
| Room ID | TikTok room ID, or a generated ID if the export had none |
| Room Title | Title as exported |
| Start Time | Stream start date and time |
| End Time | Stream end date and time, if exported |
| Date | Date of the stream start |
| Duration Minutes | Stream length in whole minutes |
| Duration (HH:MM) | Stream length in `[hh]:mm` format |
| Attributed GMV | GMV credited to the stream |
| Streamer | Assigned streamer name |
| Assignment Method | How the streamer was assigned: `auto: title alias`, `auto: explicit host marker`, `auto: learned explicit alias`, `auto: direct host phrase`, `manual: table cell edit`, `manual: bulk table assignment`, `manual: copied previous row` / `manual: copied next row`, or `session: room assignment` (a manual choice re-applied after a re-import). A rename adds `+ manual rename/merge`. |
| Extracted Candidate | The text from the title that triggered an automatic assignment |
| `Raw: <header>` | Every original column from the TikTok export, so no source data is hidden |

## Settings, config and cache

None. The app doesn't write any config, cache or state files. Manual assignments, renames and new names last only until you close the app. Each launch starts fresh with the built-in alias list.

To change the built-in aliases or the streamer column order, edit the `KNOWN_ALIASES` and `PREFERRED_ORDER` lists near the top of `tiktok_streamer_reporter.py`. To change the GMV color bands, edit `GMV_BANDS`.

## Building the app with PyInstaller

The prebuilt app was made with PyInstaller as a windowed macOS `.app` bundle. No spec file is included; this command produces an equivalent build:

```bash
# macOS
python3 -m pip install pyinstaller XlsxWriter
python3 -m PyInstaller --windowed --name "TikTok Streamer Reporter" tiktok_streamer_reporter.py

# Windows
py -m pip install pyinstaller XlsxWriter
py -m PyInstaller --windowed --name "TikTok Streamer Reporter" tiktok_streamer_reporter.py
```

The output goes to `dist/`. On macOS that's `dist/TikTok Streamer Reporter.app`. On Windows it's `dist\TikTok Streamer Reporter\TikTok Streamer Reporter.exe`.

XlsxWriter must be installed in the same Python environment you build from. If it isn't, the built app can't write reports.

PyInstaller doesn't cross-compile. Build the Windows version on Windows and the Mac version on a Mac. A build made on Apple Silicon runs only on arm64 Macs.

## Troubleshooting

| Problem | Fix |
|---|---|
| "couldn't find TikTok headers" | The file has no row with `Room ID`, `Room Title` and `Start Time` in its first 100 rows. Re-export the LIVE performance report from TikTok Shop and don't edit the headers. |
| "unsupported input type" | Only `.csv` and `.xlsx` are supported. Re-save `.xls` files as `.xlsx`. |
| "not a valid .xlsx file" | The file is damaged or isn't really an XLSX file. Re-download it, or export it as CSV. |
| "Could not import one or more files" | Every listed file must import, or none are loaded. The message names the failing file; remove it with **Remove Selected** and add a fixed copy. |
| "No valid TikTok live rows were found" | None of the rows have a Start Time the app can read. Check that the date format is one of those listed under Inputs. |
| **Generate Excel Report** is greyed out | Some rows still need a streamer. Click **Show Needs Streamer**. |
| "Missing dependency 'xlsxwriter'" (CLI) or "This app needs XlsxWriter" (GUI) | Run `py -m pip install XlsxWriter` (Windows) or `python3 -m pip install XlsxWriter` (macOS). If you see this in a built app, rebuild it with XlsxWriter installed. |
| Wrong automatic name | Edit the cell, or use **Rename / Merge Name...** to fix every row with that name. Clearing a cell also overrides the automatic match for the current session. To fix it permanently, edit `KNOWN_ALIASES` in the source. |
| Two names for one person | Select a row with the name to replace, click **Rename / Merge Name...**, and enter the name you want to keep. |
| GMV is lower than expected | The newest export wins for duplicate Room IDs. Check `Source File` in `Combined Data` to see which file each row came from. |
| macOS says the app is damaged or from an unidentified developer | The app isn't notarized. Right-click it and choose **Open**, or allow it in System Settings > Privacy & Security. |
| Prebuilt binary won't start on Windows or an Intel Mac | The included binaries are arm64 macOS only. Run from source or build the app on that platform. |

## File guide

| File | What it is |
|---|---|
| `tiktok_streamer_reporter.py` | Full source: GUI, CLI, import, deduplication, streamer detection and Excel writer |
| `TikTok Streamer Reporter` | Prebuilt standalone macOS arm64 executable |
| `TikTok Streamer Reporter.app.zip` | Prebuilt macOS app bundle, zipped |
| `TikTok Streamer Zip.zip` | Older bundle containing `MAC TikTok Streamer Reporter.app` and a file named `WINDOWS TikTok Streamer Reporter`. The "WINDOWS" file is also a macOS arm64 executable, not a Windows `.exe`. Windows users should run from source. |
| `README.md` | This document |
