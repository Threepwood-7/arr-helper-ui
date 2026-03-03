@ECHO OFF

CD /D %~dp0

PYTHON.EXE -m pytest -q
