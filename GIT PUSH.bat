@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 goto :failed

rem Trust only this repository, whose location is determined by this script.
set "repoPath=%~dp0"
set "repoPath=%repoPath:~0,-1%"
set "repoPath=%repoPath:\=/%"
git config --global --get-all safe.directory | findstr /x /l /c:"%repoPath%" >nul
if errorlevel 1 (
  git config --global --add safe.directory "%repoPath%"
  if errorlevel 1 goto :failed
)

git rev-parse --show-toplevel
if errorlevel 1 goto :failed
for /f "delims=" %%B in ('git branch --show-current') do set "pushBranch=%%B"
if not defined pushBranch (
  echo ERROR: Detached HEAD. Switch to a branch first.
  goto :failed
)

echo Preparing changes on %pushBranch%...
git add .
if errorlevel 1 goto :failed

git diff --cached --quiet
if errorlevel 2 goto :failed
if errorlevel 1 (
  git commit -m "Update %date% %time%"
  if errorlevel 1 goto :failed
) else (
  echo No new changes to commit.
)

git push -u origin "%pushBranch%"
if errorlevel 1 goto :failed

echo Push completed successfully.
pause
exit /b 0

:failed
echo ERROR: Operation failed. Push was not confirmed.
pause
exit /b 1
