@echo off
cd /d "%~dp0"
echo Creating project Python environment if needed...
if not exist ".venv\Scripts\python.exe" (
  py -3.13 -m venv .venv
  if errorlevel 1 (
    echo Could not create the Python 3.13 environment. Install Python 3.13 first.
    pause
    exit /b 1
  )
)
echo Upgrading pip...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :failed
echo Installing the Python API requirements...
".venv\Scripts\python.exe" -m pip install -r requirements-api.txt
if errorlevel 1 goto :failed
echo Setup complete. Double-click start_api.bat to start the API.
pause
exit /b 0
:failed
echo Installation failed. Please take a screenshot of this window.
pause
exit /b 1
