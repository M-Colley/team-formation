@echo off
rem One-time setup: creates a private Python environment (.venv) in this folder
rem and installs the packages from requirements.txt. Nothing else on the PC is changed.
cd /d "%~dp0"
where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")
%PY% -c "import sys; sys.exit(sys.version_info < (3, 11))" >nul 2>nul || (
  echo Python 3.11 or newer is needed. Install it from https://www.python.org/downloads/
  echo and tick "Add python.exe to PATH" during installation, then run this file again.
  pause & exit /b 1
)
if not exist .venv (
  echo Creating .venv ...
  %PY% -m venv .venv || (echo Could not create the virtual environment. & pause & exit /b 1)
)
.venv\Scripts\python.exe -m pip install --upgrade pip >nul
.venv\Scripts\python.exe -m pip install -r requirements.txt || (echo Package install failed. & pause & exit /b 1)
echo.
echo Setup complete. Drag the filled-in intake workbook onto allocate.bat to form the teams.
pause
