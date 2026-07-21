# Code Analysis — `compta_analyzer_en.py` (v1)

Analysis of the originally uploaded `compta_analyzer_en.zip` (1,529-line
Python file + `launch.bat`). Every issue below was addressed in the v2
rewrite that lives alongside this document; the v1 references remain here as
the rationale for those changes.

## What it is

**AccountingAnalyzer**, a self-contained Tkinter desktop app plus a Windows
launcher (`launch.bat` — pip-installs `openpyxl`, `matplotlib`, `reportlab`,
then runs the script). The app loads an Excel workbook and offers eight
pages: raw data grid, financial summary with computed ratios, category
breakdown with pie/bar charts, bank reconciliation between two sheets,
z-score anomaly detection, a general chart builder (6 chart types), and PDF
report export. The naming (`compta`, a leftover French `'charge'` keyword)
indicates it is an English translation of a French original. **Nothing
malicious** — a legitimate offline accounting tool; the only
security-adjacent flag is that it runs `pip install` at import time.

The structure is clean for a single file: `ExcelAnalyzer` (pure analysis
engine, no UI dependencies), `PDFReportGenerator` (a tidy reportlab
wrapper), and `AccountingApp` (all the Tkinter UI). The analysis math
(mean/variance/z-scores computed by hand, no numpy/pandas) is correct.

## Correctness bugs, most serious first

1. **Reconciliation double-matches duplicate amounts**
   (`bank_reconciliation`). The inner loop finds a match in sheet 2 and
   breaks, but never removes the matched entry from the search list — only
   from the "unmatched" report list. Three payments of $500.00 in sheet 1
   against one in sheet 2 are all reported as "matched" to the same row,
   inflating the matched count and match rate. Duplicate amounts are the
   norm in accounting data, so this hits the tool's core use case. Also
   O(n²). *(v2: sorted two-pointer matching, each entry consumed once.)*

2. **`.xls` files crash despite being advertised.** The home page and the
   file dialog both claim `.xlsx, .xls` support, but openpyxl reads only
   `.xlsx`/`.xlsm` — opening a real legacy `.xls` fails into the error
   dialog. *(v2: real `.xls` support via xlrd.)*

3. **Duplicate column headers select the wrong column.** Every picker
   resolves the chosen name with `headers.index(name)`, which always
   returns the first occurrence. Two columns named "Amount" means the
   breakdown, reconciliation, anomaly and chart pages may silently analyze
   the wrong one. *(v2: pickers are index-based with `A — Amount` labels.)*

4. **Two of the four PDF export checkboxes do nothing.** "Category
   Breakdown" and "Charts" are displayed and checkable, but `_export_pdf`
   only implements the "Financial Summary" and "Anomalies" sections.
   *(v2: all sections implemented, plus a Macro Analysis section.)*

5. **Scrolling is Windows-only and globally hijacked.** The Financial page
   uses `canvas.bind_all('<MouseWheel>')` with `event.delta / 120`. On
   Linux wheel events are `Button-4/5` (nothing scrolls); on macOS the
   delta scale is wrong; and `bind_all` makes the wheel scroll that hidden
   canvas from every page. *(v2: pointer-scoped cross-platform handler.)*

6. **Silent data loss with formula cells.** `load_workbook(data_only=True)`
   returns the *cached* value of formulas — `None` if the file was
   generated programmatically and never opened in Excel. A workbook full of
   `=SUM(...)` cells can silently analyze as empty. *(v2: uncached-formula
   warning surfaced in the UI and the PDF.)*

7. **Silent truncation.** The data grid shows only the first 500 rows while
   the label reports the full count; Bar/Pie charts keep only the *first*
   30 rows. *(v2: truncation is labeled.)*

8. **Minor inconsistencies:** the PDF anomaly section hardcodes the 2.0σ
   threshold, ignoring the UI selector; booleans count as numeric during
   column detection (`float(True)`); `plt.close(fig_pdf)` is a no-op on a
   `Figure` not created through pyplot; `self.chart_figures` accumulates
   every figure ever generated (slow memory leak). *(All fixed in v2.)*

9. **`check_dependencies()` pip-installs at import time.** For an app whose
   docstring says "offline desktop application",
   `subprocess.check_call([..., 'pip', 'install', ...])` crashes the launch
   when there is no network, duplicates `launch.bat`'s job, and installs
   unpinned packages without asking. *(v2: reports missing dependencies
   with install instructions instead.)*

## Dead code

- Unused imports: `json`, `tempfile`, `get_column_letter`, `mticker`.
- `ExcelAnalyzer.detect_date_columns()` and `analyze_time_series()` are
  never called — the Line chart plots row order rather than dates despite
  date-handling code existing. *(v2: wired into the Line/Area charts;
  `json` now powers session autosave.)*
- `self.analyses = {}` is assigned and never touched again.

## Design observations

- The whole workbook is materialized into Python lists at load;
  openpyxl's `read_only=True` streaming mode handles large files better
  *(v2 uses it)*.
- Currency formatting is hardcoded to `$` with US separators, and
  `compute_ratios` matches revenue/expense columns by English keywords
  only — fragile for the tool's francophone origin.
- Ratio computation used only the *first* matched revenue and expense
  columns *(v2 sums across all matches)*.
- No tests, no type hints — though the UI-free `ExcelAnalyzer` is very
  testable *(v2 ships an 18-test suite, fixtures in 6 formats, a headless
  `--self-test` mode and a scripted `--smoke-test` GUI run)*.
