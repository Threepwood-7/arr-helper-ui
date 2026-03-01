@echo off
cd /d "%~dp0"
START "" pythonw.exe sonarr_ui_helper.py %*
