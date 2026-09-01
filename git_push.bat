@echo off
REM Quick helper: commits any local changes with the message "quick fix", then
REM pushes. Saves typing the individual git commands each time.

cd /d "%~dp0"

echo == Checking for local changes ==
git add -A
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "quick fix"
) else (
    echo No local changes to commit.
)

echo.
echo == Pushing ==
git push

pause
