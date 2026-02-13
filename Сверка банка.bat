@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
  echo [ERROR] .venv is not found.
  echo Run setup first:
  echo powershell -ExecutionPolicy Bypass -File .\setup.ps1
  pause
  exit /b 1
)

"%~dp0.venv\Scripts\pythonw.exe" "%~dp0compare_payments.py" --gui
if errorlevel 1 (
  echo.
  echo [ERROR] Script finished with an error.
  pause
  exit /b 1
)

endlocal
