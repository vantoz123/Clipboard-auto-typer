@echo off
setlocal
cd /d "%~dp0"
echo Creating virtual environment if needed...
py -m venv .venv
if errorlevel 1 (
  echo Failed to create virtual environment. Make sure Python is installed and available as py.
  pause
  exit /b 1
)
echo Upgrading pip...
.venv\Scripts\python.exe -m pip install --upgrade pip
echo Installing requirements...
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
  echo Dependency installation failed.
  pause
  exit /b 1
)
echo.
echo Installation complete. You can now run run_app.bat.
pause
