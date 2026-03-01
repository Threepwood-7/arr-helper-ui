@ECHO OFF

CD /D %~dp0

START "" /MAX "C:\Program Files\Microsoft VS Code\Code.exe" -n %~dp0
