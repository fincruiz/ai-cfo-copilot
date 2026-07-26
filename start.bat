@echo off
setlocal
title AI CFO Copilot 1.0
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creating the local Python environment...
  uv venv --python 3.12 .venv || goto :error
)
echo Checking application packages...
uv pip install --python ".venv\Scripts\python.exe" -r requirements.txt || goto :error
echo Starting AI CFO Copilot...
".venv\Scripts\python.exe" -m streamlit run app.py
goto :end
:error
echo.
echo Setup could not be completed. Confirm that uv and Python 3.12 are installed.
pause
:end
endlocal
