@echo off
echo ============================================
echo   AccountingAnalyzer v2 - Spreadsheet Analysis Tool
echo ============================================
echo.
echo Installing dependencies...
pip install openpyxl xlrd matplotlib reportlab oletools -q
echo.
echo Launching application...
python compta_analyzer_en.py
pause
