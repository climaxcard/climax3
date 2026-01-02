@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

set REPO_DIR=C:\Users\user\ClimaxGit\climax3
set BRANCH=main
set LOG=%REPO_DIR%\publish_log.txt

call :MAIN > "%LOG%" 2>&1
type "%LOG%"
echo.
echo ===== LOG FILE =====
echo %LOG%
pause
exit /b

:MAIN
echo ===== START publish =====
echo DATE : %date% %time%
echo.

cd /d "%REPO_DIR%" || (echo [ERR] repo dir not found & exit /b 1)

REM --- Git editor問題回避（vi起動させない）---
set GIT_EDITOR=true
git config core.editor true >nul 2>&1

echo ===== git fetch =====
git fetch origin
if errorlevel 1 (echo [ERR] fetch failed & exit /b 1)

echo ===== git pull --rebase --autostash =====
git pull --rebase --autostash origin %BRANCH%
if errorlevel 1 (
  echo [ERR] rebase failed. run: git status
  git status
  exit /b 1
)

echo.
echo ===== build static pages =====
python scripts\build_pokeka_static.py
if errorlevel 1 (echo [ERR] build failed & exit /b 1)

echo.
echo ===== stage docs only =====
git add docs

REM docsに変更が無ければ commit はしない
git diff --cached --quiet
if %errorlevel%==0 (
  echo [OK] no changes in docs
  goto :PUSH_ONLY
)

REM commit message
for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set D=%%a%%b%%c
for /f "tokens=1-3 delims=:. " %%a in ("%time%") do set T=%%a%%b%%c
set MSG=publish pokeca default %D%_%T%

echo.
echo ===== git commit =====
git commit -m "%MSG%"
if errorlevel 1 (
  echo [ERR] commit failed
  exit /b 1
)

:PUSH_ONLY
echo.
echo ===== git push =====
git push origin %BRANCH%
if errorlevel 1 (
  echo [ERR] push failed
  exit /b 1
)

echo.
echo [OK] publish SUCCESS
exit /b 0
