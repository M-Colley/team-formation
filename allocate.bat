@echo off
rem Drag the filled-in intake workbook onto this file, or run:
rem   allocate.bat "C:\path\to\team_formation.xlsx" [--deterministic --seed 1 ...]
set "HERE=%~dp0"
if not exist "%HERE%.venv\Scripts\python.exe" (
  echo First run setup_windows.bat once.
  pause & exit /b 1
)
if "%~1"=="" (
  echo Drag the intake workbook onto allocate.bat, or pass its path as the first argument.
  pause & exit /b 1
)
"%HERE%.venv\Scripts\python.exe" "%HERE%allocate_teams.py" %*
pause
