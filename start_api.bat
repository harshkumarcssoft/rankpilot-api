@echo off
cd /d "%~dp0"
echo Starting Rankpilot Python crawler API on http://127.0.0.1:8000
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m uvicorn api:app --host 127.0.0.1 --port 8000
) else (
  echo ERROR: Project virtual environment not found. Run setup_api.bat first.
  pause
  exit /b 1
)
pause
