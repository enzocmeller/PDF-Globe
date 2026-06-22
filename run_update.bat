@echo off
REM ============================================================
REM  One command:  update the USDA CSVs, then refresh Excel.
REM  Double-click this file (or run it from a scheduled task).
REM ============================================================
cd /d "%~dp0"

echo === Step 1/2: pulling latest two USDA releases ===
"C:\Program Files\Python314\python.exe" "%~dp0update_psd.py"
if errorlevel 1 (
  echo.
  echo Python step FAILED -- Excel was NOT refreshed.
  pause
  exit /b 1
)

echo.
echo === Step 2/2: refreshing Excel workbook ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0refresh_excel.ps1"

echo.
echo === Step 3/3: exporting PowerPoint ===
for /f %%I in ('powershell -NoProfile -Command "(Get-Date).ToString('yyyy-MM-dd')"') do set "TODAY=%%I"
set "PPT_NAME=AUS SnD Commodities 2026 - %TODAY%.pptx"
set "OUTPUT_PATH=%USERPROFILE%\Desktop\%PPT_NAME%"
if exist "%OUTPUT_PATH%" del /f /q "%OUTPUT_PATH%"
"%~dp0.venv\Scripts\python.exe" "%~dp0export_pdfs.py" "%~dp0USDA_PSD.xlsx" "%OUTPUT_PATH%"
if errorlevel 1 (
  echo.
  echo PowerPoint export step FAILED.
  pause
  exit /b 1
)

echo.
echo PowerPoint generated: "%OUTPUT_PATH%"
echo.
echo Done - all tasks completed successfully!
pause