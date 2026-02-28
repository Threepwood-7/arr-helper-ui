@echo off
echo Installing dependencies...
pip install -r "%~dp0requirements.txt"
echo.
echo Done. You can now run:
echo   run_checker.cmd  - Media Quality Checker (CLI)
echo   run_ui.cmd       - Sonarr UI Helper (GUI)
pause
