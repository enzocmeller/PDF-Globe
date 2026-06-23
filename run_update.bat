@echo off
REM ============================================================
REM  One command:  update the USDA CSVs, then refresh Excel.
REM  Double-click this file (or run it from a scheduled task).
REM ============================================================
cd /d "%~dp0"

REM Firewall-friendly: all dependencies live in .venv, so never reach out to
REM PyPI at run time (these env vars tell the download scripts to skip pip).
set "NO_PIP=1"
set "GET_GLOBE_NO_PIP=1"

REM Auto-bootstrap on a fresh machine: build .venv + install deps if missing.
if not exist "%~dp0.venv\Scripts\python.exe" (
  echo .venv not found -- running one-time setup ...
  call "%~dp0setup.bat"
)
if not exist "%~dp0.venv\Scripts\python.exe" (
  echo.
  echo ERROR: setup did not complete -- cannot continue. See messages above.
  pause
  exit /b 1
)

echo === Step 1/2: pulling latest two USDA releases ===
"%~dp0.venv\Scripts\python.exe" "%~dp0update_psd.py"
if errorlevel 1 (
  echo.
  echo Python step FAILED -- Excel was NOT refreshed.
  pause
  exit /b 1
)

echo.
echo === Step 2/4: refreshing Excel workbook ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0refresh_excel.ps1"
if errorlevel 1 (
  echo.
  echo Excel refresh FAILED.
  pause
  exit /b 1
)

echo.
echo === Step 3/4: downloading + trimming globe and El Nino images ===
"%~dp0.venv\Scripts\python.exe" "%~dp0final map.py"
if errorlevel 1 (
  echo.
  echo Globe download FAILED -- continuing with whatever map.png is on disk.
)
"%~dp0.venv\Scripts\python.exe" "%~dp0final_elnino.py"
if errorlevel 1 (
  echo.
  echo El Nino download FAILED -- continuing with whatever elnino.png is on disk.
)

echo.
echo === Step 3b/4: inserting images into Excel ===
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0insert_images.ps1"
if errorlevel 1 (
  echo.
  echo Image insertion FAILED.
  pause
  exit /b 1
)

echo.
echo === Step 4/4: exporting PDF deck ===
for /f %%I in ('powershell -NoProfile -Command "(Get-Date).ToString('yyyy-MM-dd')"') do set "TODAY=%%I"
set "PDF_NAME=AUS SnD Commodities 2026 - %TODAY%.pdf"
set "OUTPUT_PATH=%USERPROFILE%\Desktop\%PDF_NAME%"
if exist "%OUTPUT_PATH%" del /f /q "%OUTPUT_PATH%"
"%~dp0.venv\Scripts\python.exe" "%~dp0export_pdf_deck.py" "%~dp0USDA_PSD.xlsx" "%OUTPUT_PATH%"
if errorlevel 1 (
  echo.
  echo PDF export step FAILED.
  pause
  exit /b 1
)

echo.
echo PDF deck generated: "%OUTPUT_PATH%"
echo.
echo Done - all tasks completed successfully!
pause