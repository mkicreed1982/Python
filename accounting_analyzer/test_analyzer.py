#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Headless test suite for the AccountingAnalyzer engine.

Run:  python make_fixtures.py && python test_analyzer.py
No display or tkinter needed — only the engine, session and PDF layers are
exercised. Exit code 0 = all tests passed.
"""

import os
import subprocess
import sys
import tempfile
from collections import Counter
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from compta_analyzer_en import (  # noqa: E402
    ExcelAnalyzer, SessionManager, build_pdf_report, run_self_test,
    coerce_text_cell, HAS_XLRD, HAS_OLETOOLS,
)
import make_fixtures  # noqa: E402

FIXTURES = os.path.join(HERE, 'fixtures')

PASSED = []
FAILED = []


def check(name, fn):
    try:
        fn()
        PASSED.append(name)
        print(f"  PASS  {name}")
    except Exception as e:
        FAILED.append((name, e))
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")


def load(filename):
    analyzer = ExcelAnalyzer()
    analyzer.load_file(os.path.join(FIXTURES, filename))
    return analyzer


def make_sheet(headers, rows):
    data = [headers] + rows
    return {
        'data': data, 'headers': headers, 'rows': rows,
        'num_rows': len(rows), 'num_cols': len(headers)
    }


# ---------------------------------------------------------------- loaders
def test_loader_parity():
    """Every format yields the same logical Ledger table."""
    files = ['ledger.xlsx', 'ledger.xml', 'ledger.csv', 'ledger.ods']
    if HAS_XLRD:
        files.append('ledger.xls')
    reference_amounts = [float(r[3]) for r in make_fixtures.LEDGER_ROWS]
    reference_date = make_fixtures.LEDGER_ROWS[0][0]

    for filename in files:
        analyzer = load(filename)
        first_sheet = analyzer.get_sheet_names()[0]
        sheet = analyzer.get_sheet_data(first_sheet)

        headers = [str(h) for h in sheet['headers']]
        assert headers == make_fixtures.LEDGER_HEADERS, \
            f"{filename}: headers {headers}"
        assert sheet['num_rows'] == len(make_fixtures.LEDGER_ROWS), \
            f"{filename}: {sheet['num_rows']} rows"

        amounts = [float(r[3]) for r in sheet['rows']]
        assert amounts == reference_amounts, f"{filename}: amounts differ"

        first_date = sheet['rows'][0][0]
        assert isinstance(first_date, datetime), \
            f"{filename}: date cell is {type(first_date).__name__}"
        assert first_date == reference_date, f"{filename}: {first_date}"


def test_multisheet_formats():
    for filename in ['ledger.xlsx', 'ledger.xml', 'ledger.ods'] + \
            (['ledger.xls'] if HAS_XLRD else []):
        analyzer = load(filename)
        names = analyzer.get_sheet_names()
        assert names[:3] == ['Ledger', 'Bank', 'PnL'], f"{filename}: {names}"


def test_formula_warning():
    analyzer = load('ledger.xlsx')
    assert any('Totals' in w for w in analyzer.warnings), analyzer.warnings


def test_csv_coercion():
    assert coerce_text_cell('$1,234.56') == 1234.56
    assert coerce_text_cell('(500.00)') == -500.0
    assert coerce_text_cell('2025-01-05') == datetime(2025, 1, 5)
    assert coerce_text_cell('  ') is None
    assert coerce_text_cell('Rent') == 'Rent'


# ---------------------------------------------------------------- macros
def test_macro_detection():
    assert load('ledger.xlsm').macro_info['has_macros'] is True
    assert load('ledger.xlsx').macro_info['has_macros'] is False
    assert load('ledger.csv').macro_info['has_macros'] is False


# ---------------------------------------------------------------- reconciliation
def test_reconciliation_duplicates_synthetic():
    """Two identical amounts on one side must not both match a single entry."""
    analyzer = ExcelAnalyzer()
    analyzer.sheets = {
        'A': make_sheet(['Amount'], [[500.0], [500.0]]),
        'B': make_sheet(['Amount'], [[500.0]]),
    }
    result = analyzer.bank_reconciliation('A', 0, 'B', 0)
    assert len(result['matched']) == 1, result['matched']
    assert len(result['unmatched_sheet1']) == 1, result['unmatched_sheet1']
    assert len(result['unmatched_sheet2']) == 0, result['unmatched_sheet2']


def test_reconciliation_fixture():
    analyzer = load('ledger.xlsx')
    result = analyzer.bank_reconciliation('Ledger', 3, 'Bank', 2)

    ledger_amounts = Counter(round(float(r[3]), 2) for r in make_fixtures.LEDGER_ROWS)
    bank_amounts = Counter(round(float(r[2]), 2) for r in make_fixtures.BANK_ROWS)
    expected_matches = sum((ledger_amounts & bank_amounts).values())

    assert len(result['matched']) == expected_matches, \
        f"{len(result['matched'])} != {expected_matches}"
    assert len(result['unmatched_sheet1']) == \
        len(make_fixtures.LEDGER_ROWS) - expected_matches
    assert len(result['unmatched_sheet2']) == \
        len(make_fixtures.BANK_ROWS) - expected_matches
    # the duplicated -500.00 rent: exactly one side matched
    unmatched_vals = [round(u['value'], 2) for u in result['unmatched_sheet1']]
    assert unmatched_vals.count(-500.0) == 1, unmatched_vals


def test_reconciliation_tolerance():
    analyzer = ExcelAnalyzer()
    analyzer.sheets = {
        'A': make_sheet(['Amount'], [[100.0], [200.005]]),
        'B': make_sheet(['Amount'], [[100.009], [200.0]]),
    }
    result = analyzer.bank_reconciliation('A', 0, 'B', 0, tolerance=0.01)
    assert len(result['matched']) == 2, result


# ---------------------------------------------------------------- analyses
def test_financial_summary():
    analyzer = load('ledger.xlsx')
    summary = analyzer.analyze_financial_summary('Ledger')
    amounts = [float(r[3]) for r in make_fixtures.LEDGER_ROWS]
    assert abs(summary['Amount']['total'] - sum(amounts)) < 1e-6
    assert summary['Amount']['count'] == len(amounts)
    assert summary['Amount']['max'] == 50000.0


def test_category_breakdown():
    analyzer = load('ledger.xlsx')
    results = analyzer.analyze_by_category('Ledger', 2, 3)
    expected = {}
    for row in make_fixtures.LEDGER_ROWS:
        expected[row[2]] = expected.get(row[2], 0) + float(row[3])
    assert set(results) == set(expected)
    for cat, total in expected.items():
        assert abs(results[cat]['total'] - total) < 1e-6, cat


def test_anomaly_detection():
    analyzer = load('ledger.xlsx')
    anomalies = analyzer.detect_anomalies('Ledger', 3, threshold=2.0)
    outlier_row = next(i for i, r in enumerate(make_fixtures.LEDGER_ROWS)
                       if r[3] == 50000.0) + 2
    assert any(a['row'] == outlier_row and a['value'] == 50000.0
               for a in anomalies), anomalies


def test_ratios():
    analyzer = load('ledger.xlsx')
    summary = analyzer.analyze_financial_summary('PnL')
    ratios = analyzer.compute_ratios(summary)
    revenue = sum(r[1] for r in make_fixtures.PNL_ROWS)
    expenses = sum(r[2] for r in make_fixtures.PNL_ROWS)
    expected_margin = (revenue - expenses) / revenue * 100
    assert abs(ratios['Margin (%)'] - expected_margin) < 1e-6, ratios


def test_time_series():
    analyzer = load('ledger.xlsx')
    series = analyzer.analyze_time_series('Ledger', 0, 3)
    assert len(series) == len(make_fixtures.LEDGER_ROWS)
    dates = [d for d, _ in series]
    assert dates == sorted(dates)


def test_detect_columns():
    analyzer = load('ledger.xlsx')
    numeric = {h for _, h in analyzer.detect_numeric_columns('Ledger')}
    assert numeric == {'Amount', 'Balance'}, numeric
    dates = {h for _, h in analyzer.detect_date_columns('Ledger')}
    assert dates == {'Date'}, dates
    cats = {h for _, h in analyzer.detect_category_columns('Ledger')}
    assert 'Category' in cats, cats


# ---------------------------------------------------------------- session
def test_session_roundtrip():
    with tempfile.TemporaryDirectory() as tmp:
        session = SessionManager(directory=tmp)
        assert session.load() is None
        state = {'file_path': '/tmp/x.xlsx', 'selections': {'fin_sheet': 2}}
        session.save(state)
        loaded = session.load()
        assert loaded['file_path'] == '/tmp/x.xlsx'
        assert loaded['selections']['fin_sheet'] == 2
        assert 'saved_at' in loaded
        session.clear()
        assert session.load() is None


# ---------------------------------------------------------------- PDF
def test_pdf_report():
    analyzer = load('ledger.xlsm')
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, 'report.pdf')
        build_pdf_report(analyzer, out, anomaly_threshold=2.0)
        assert os.path.getsize(out) > 5000, os.path.getsize(out)
        with open(out, 'rb') as f:
            assert f.read(5) == b'%PDF-'


# ---------------------------------------------------------------- self-test mode
def test_self_test_mode():
    assert run_self_test(FIXTURES) == 0


def test_self_test_cli():
    proc = subprocess.run(
        [sys.executable, os.path.join(HERE, 'compta_analyzer_en.py'),
         '--self-test', FIXTURES],
        capture_output=True, text=True, timeout=120)
    assert proc.returncode == 0, proc.stdout + proc.stderr
    assert 'PASS' in proc.stdout


def main():
    if not os.path.isdir(FIXTURES):
        print("fixtures/ missing — run: python make_fixtures.py")
        return 2

    print(f"xlrd available: {HAS_XLRD}   oletools available: {HAS_OLETOOLS}\n")

    check('loader parity across formats', test_loader_parity)
    check('multi-sheet formats', test_multisheet_formats)
    check('uncached-formula warning', test_formula_warning)
    check('csv value coercion', test_csv_coercion)
    check('macro detection', test_macro_detection)
    check('reconciliation: duplicates (synthetic)', test_reconciliation_duplicates_synthetic)
    check('reconciliation: fixture ledger vs bank', test_reconciliation_fixture)
    check('reconciliation: tolerance', test_reconciliation_tolerance)
    check('financial summary', test_financial_summary)
    check('category breakdown', test_category_breakdown)
    check('anomaly detection', test_anomaly_detection)
    check('ratios', test_ratios)
    check('time series', test_time_series)
    check('column detection', test_detect_columns)
    check('session save/load round-trip', test_session_roundtrip)
    check('PDF report generation', test_pdf_report)
    check('self-test mode (in-process)', test_self_test_mode)
    check('self-test mode (CLI)', test_self_test_cli)

    print(f"\n{len(PASSED)} passed, {len(FAILED)} failed")
    return 1 if FAILED else 0


if __name__ == '__main__':
    sys.exit(main())
