@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Virtual environment not found. Running install_dependencies.bat first...
  call install_dependencies.bat
)
.venv\Scripts\python.exe diagnostics.py
pause
