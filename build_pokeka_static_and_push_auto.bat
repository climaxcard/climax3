@echo off
setlocal

REM ============================================
REM  Pokeka static HTML build -> GitHub auto push
REM  - Place this file at:
REM      C:\Users\user\ClimaxGit\climax3
REM  - Python script: scripts\build_pokeka_static.py
REM ============================================

REM move to repo root
cd /d "%~dp0"

echo [1] run python (scripts\build_pokeka_static.py)

REM fixed Excel path / sheet / output dir
set "EXCEL_PATH=C:\Users\user\ClimaxGit\climax3\data\pokeca_rush.xlsm"
set "SHEET_NAME=Sheet1"
set "OUT_DIR=docs"

python "scripts\build_pokeka_static.py"
if errorlevel 1 (
    echo [ERROR] build_pokeka_static.py failed.
    goto :END
)

echo.
echo [2] git add -A
git add -A
if errorlevel 1 (
    echo [ERROR] git add failed.
    goto :END
)

echo.
echo [3] git commit

REM build simple auto message with date/time
set "MSG=auto: update pokeka static pages %date% %time%"
git commit -m "%MSG%"
if errorlevel 1 (
    echo [WARN] git commit failed (maybe no changes).
    goto :END
)

echo.
echo [4] git push
git push
if errorlevel 1 (
    echo [ERROR] git push failed.
    goto :END
)

echo.
echo [OK] build + commit + push completed.

:END
endlocal
