@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [SCM] Creating local Python environment...
  python -m venv .venv
  if errorlevel 1 (
    echo [SCM] Unable to create .venv. Confirm Python is installed and on PATH.
    pause
    exit /b 1
  )
)

set "PY=.venv\Scripts\python.exe"

"%PY%" -c "import flask, openpyxl, pptx, matplotlib, waitress" >nul 2>&1
if errorlevel 1 (
  echo [SCM] Installing or repairing application dependencies...
  "%PY%" -m pip install --upgrade pip
  if errorlevel 1 exit /b 1
  "%PY%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo [SCM] Dependency installation failed. Check your internet connection and requirements.txt.
    pause
    exit /b 1
  )
)

if not exist ".env" (
  echo [SCM] .env not found. Running first-time local setup...
  powershell -ExecutionPolicy Bypass -File ".\setup_windows.ps1"
  if errorlevel 1 exit /b 1
)

"%PY%" app.py
endlocal
