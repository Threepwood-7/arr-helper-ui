@echo off
CD /D %~dp0

echo Installing dependencies...
pip install -r requirements.txt
echo.
echo Done. You can now run:
echo   run_checker.cmd  - Media Quality Checker (CLI)
echo   run_ui.cmd       - Sonarr UI Helper (GUI)
pause
