@echo off
REM ============================================================================
REM  setup.bat  --  one-time environment bootstrap (firewall-tolerant)
REM ----------------------------------------------------------------------------
REM  Creates a local .venv and installs all dependencies. Tries OFFLINE first
REM  (from the bundled wheels\ folder -- needs no internet at all), then falls
REM  back to ONLINE pip with corporate-firewall-friendly flags.
REM
REM  run_update.bat calls this automatically the first time if .venv is missing,
REM  so you normally don't need to run it by hand.
REM ============================================================================
setlocal enableextensions
cd /d "%~dp0"

echo ============================================================
echo  One-time setup: virtual environment + dependencies
echo ============================================================

REM ---- 1. Create .venv using whatever Python is available -------------------
if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment .venv ...
  py -3 -m venv .venv 2>nul
)
if not exist ".venv\Scripts\python.exe" python -m venv .venv 2>nul
if not exist ".venv\Scripts\python.exe" (
  for %%P in (
    "%LocalAppData%\Programs\Python\Python313\python.exe"
    "%LocalAppData%\Programs\Python\Python312\python.exe"
    "%LocalAppData%\Programs\Python\Python311\python.exe"
    "C:\Program Files\Python313\python.exe"
    "C:\Program Files\Python312\python.exe"
    "C:\Program Files\Python311\python.exe"
  ) do (
    if not exist ".venv\Scripts\python.exe" if exist "%%~P" "%%~P" -m venv .venv
  )
)
if not exist ".venv\Scripts\python.exe" (
  echo.
  echo ERROR: Could not find Python to create the environment.
  echo        Install Python 3.10+ 64-bit from https://www.python.org/downloads/
  echo        and tick "Add Python to PATH", then run setup.bat again.
  pause
  exit /b 1
)

set "VPY=.venv\Scripts\python.exe"
echo Virtual environment ready.

REM ---- 2a. OFFLINE install from bundled wheels (no internet needed) ---------
set "INSTALLED="
if exist "wheels" (
  echo.
  echo Installing dependencies OFFLINE from wheels\ ...
  "%VPY%" -m pip install --no-index --find-links wheels -r requirements.txt
  if not errorlevel 1 set "INSTALLED=1"
)

REM ---- 2b. ONLINE fallback (firewall-tolerant: trusted hosts + system proxy) -
if not defined INSTALLED (
  echo.
  echo Offline install not available / failed; trying ONLINE pip ...
  "%VPY%" -m pip install --upgrade pip ^
      --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
  "%VPY%" -m pip install -r requirements.txt ^
      --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
  if not errorlevel 1 set "INSTALLED=1"
)

if not defined INSTALLED (
  echo.
  echo ERROR: Dependency install failed both offline and online.
  echo   * If your network blocks PyPI, run "pip download -r requirements.txt -d wheels"
  echo     on a machine WITH internet ^(same Python 3.x 64-bit^) and copy the wheels\
  echo     folder next to this script, then re-run setup.bat.
  pause
  exit /b 1
)

REM ---- 3. Verify everything imports ----------------------------------------
echo.
echo Verifying dependencies ...
"%VPY%" -c "import pandas,requests,fitz,win32com.client,PIL,truststore; print('All dependencies OK')"
if errorlevel 1 (
  echo ERROR: dependencies installed but the import check failed.
  pause
  exit /b 1
)

echo.
echo ============================================================
echo  Setup complete.  Run run_update.bat to generate the deck.
echo ============================================================
endlocal
