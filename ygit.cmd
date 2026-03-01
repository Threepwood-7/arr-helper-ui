@echo off

CD /D %~dp0

REM Quick git add and commit script
REM Usage: ygit.cmd

echo Adding all files...
git add -A

if errorlevel 1 (
    echo Error: Failed to add files
    exit /b 1
)

echo Committing with message "OK"...
git commit -m "OK"

if errorlevel 1 (
    echo Error: Commit failed or nothing to commit
    exit /b 1
)

git push

echo Done!
