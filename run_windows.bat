@echo off
setlocal
cd /d "%~dp0"
title SCM IDP Dashboard

if not exist .venv (
    echo Creating Python virtual environment...
    python -m venv .venv
    if errorlevel 1 goto :error
)

call .venv\Scripts\activate
if errorlevel 1 goto :error

python -c "import flask, pptx, matplotlib, waitress" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if errorlevel 1 goto :error
)

echo.
echo Starting SCM IDP Dashboard...
echo The browser will open automatically.
echo.
python app.py
if errorlevel 1 goto :error
goto :end

:error
echo.
echo Dashboard could not start. Copy the error shown above into ChatGPT.
pause

:end
endlocal
