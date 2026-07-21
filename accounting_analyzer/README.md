# AccountingAnalyzer v2

Offline desktop application for accountants: load a spreadsheet, explore the
data, run financial analyses (summary, category breakdown, bank
reconciliation, anomaly detection), chart it, and export a PDF report.

## What's new in v2

- **Any spreadsheet format**, not just `.xlsx`:

  | Format | Extension | Reader |
  |---|---|---|
  | Excel workbook | `.xlsx` `.xltx` | openpyxl |
  | Excel macro-enabled | `.xlsm` `.xltm` | openpyxl + static macro analysis |
  | Excel legacy | `.xls` | xlrd |
  | Excel 2003 XML (SpreadsheetML) | `.xml` | built-in parser |
  | CSV / TSV / text | `.csv` `.tsv` `.txt` | built-in (delimiter sniffing, `$1,234.56` and `(500.00)` coercion) |
  | OpenDocument | `.ods` | built-in parser (no extra dependency) |

- **Macro analysis (safe)** — macro-enabled workbooks are inspected
  *statically* with oletools: module names, auto-exec triggers, suspicious
  keywords. **Macros are never executed.** See the new 🧬 Macros page and the
  "Macro Analysis" section of the PDF report.
- **Session autosave** — your open file, page, and every selection are saved
  automatically (every 60 s, after each analysis, and on close) to
  `~/.accounting_analyzer/session.json`. On the next launch the app offers to
  restore where you left off.
- **Correctness fixes** over v1:
  - Bank reconciliation now matches each entry at most once — duplicate
    amounts (two rents of $500) no longer double-match a single bank line —
    and runs in O(n log n).
  - Column pickers show `A — Amount` style labels and select by position, so
    duplicate column names can't silently target the wrong column.
  - All PDF export checkboxes actually work (Category Breakdown and Charts
    sections were previously ignored), and the anomaly section honors the
    threshold chosen in the UI.
  - Mouse-wheel scrolling works on Windows, macOS and Linux, and no longer
    hijacks the wheel on every page.
  - Files whose formulas have no cached value now produce a visible warning
    instead of silently loading empty cells.
  - No more `pip install` at startup — missing dependencies are reported with
    instructions instead (the app is safe to run offline).

## Install & run

```bat
launch.bat          :: Windows: installs dependencies, then starts the app
```

or manually:

```bash
pip install openpyxl xlrd matplotlib reportlab oletools
python compta_analyzer_en.py
```

`openpyxl`, `matplotlib` and `reportlab` are required; `xlrd` (legacy `.xls`)
and `oletools` (macro analysis) are optional — without them the matching
features degrade gracefully.

## Validate against your own files (no GUI needed)

Point the built-in self-test at any folder; it batch-loads every supported
spreadsheet and prints a PASS/FAIL line per file:

```bat
python compta_analyzer_en.py --self-test "C:\Users\user\OneDrive\Desktop\AI-Finance-Course"
```

## Development

```bash
python make_fixtures.py     # regenerate fixtures/ (same dataset in all 6 formats)
python test_analyzer.py     # 18 headless engine tests
python compta_analyzer_en.py --smoke-test fixtures/ledger.xlsm   # scripted GUI run
```

The engine (`ExcelAnalyzer`), session store (`SessionManager`) and PDF
builder (`build_pdf_report`) are UI-free and importable without tkinter.
