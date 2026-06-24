@echo off
REM ============================================================
REM  One command:  update the USDA CSVs, then refresh Excel.
REM  Double-click this file (or run it from a scheduled task).
REM ============================================================
cd /d "%~dp0"

REM Download scripts never need PyPI at run time.
set "NO_PIP=1"
set "GET_GLOBE_NO_PIP=1"

REM --- Pick a Python interpreter ----------------------------------------------
REM 1) the bundled .venv if present;
REM 2) else any system Python that ALREADY has the required packages;
REM 3) else build .venv via setup.bat (needs the wheels\ folder or internet).
set "PY="
if exist "%PY%" set "PY=%~dp0.venv\Scripts\python.exe"

if not defined PY (
  for %%P in (
    "%LocalAppData%\Programs\Python\Python314\python.exe"
    "%LocalAppData%\Programs\Python\Python313\python.exe"
    "%LocalAppData%\Programs\Python\Python312\python.exe"
    "C:\Program Files\Python314\python.exe"
    "C:\Program Files\Python313\python.exe"
    "C:\Program Files\Python312\python.exe"
  ) do (
    if not defined PY if exist "%%~P" ( "%%~P" -c "import pandas,fitz,win32com.client,PIL,requests" 1>nul 2>nul && set "PY=%%~P" )
  )
)

if not defined PY (
  for /f "delims=" %%P in ('where python 2^>nul') do (
    if not defined PY ( "%%P" -c "import pandas,fitz,win32com.client,PIL,requests" 1>nul 2>nul && set "PY=%%P" )
  )
)

if not defined PY (
  echo No ready-to-use Python found -- running one-time setup to build .venv ...
  call "%~dp0setup.bat"
  if exist "%PY%" set "PY=%~dp0.venv\Scripts\python.exe"
)

if not defined PY (
  echo.
  echo ERROR: Could not find a Python that has the required packages
  echo        ^(pandas, pymupdf, pywin32, Pillow, requests^), and could not build
  echo        one. Install them into your Python, or provide the wheels\ folder.
  pause
  exit /b 1
)
echo Using Python: %PY%
echo.

echo === Step 1/2: pulling latest two USDA releases ===
"%PY%" "%~dp0update_psd.py"
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
REM If the image sites are firewall-blocked, set SKIP_IMAGE_DOWNLOAD=1 to skip
REM straight to using the map.png / elnino.png already on disk (no slow retries).
if defined SKIP_IMAGE_DOWNLOAD (
  echo SKIP_IMAGE_DOWNLOAD set -- using the existing map.png / elnino.png on disk.
) else (
  "%PY%" "%~dp0final map.py"
  if errorlevel 1 echo Globe download issue -- continuing with existing map.png.
  "%PY%" "%~dp0final_elnino.py"
  if errorlevel 1 echo El Nino download issue -- continuing with existing elnino.png.
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
"%PY%" "%~dp0export_pdf_deck.py" "%~dp0USDA_PSD.xlsx" "%OUTPUT_PATH%"
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