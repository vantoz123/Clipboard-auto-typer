@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Virtual environment not found. Running install_dependencies.bat first...
  call install_dependencies.bat
)
echo Building Windows executable with PyInstaller...
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --windowed --onefile --name "Clipboard Auto Typer" main.py
if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)
echo.
echo Build complete. Find the executable in the dist folder.
pause
