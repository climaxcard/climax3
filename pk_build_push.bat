@echo off

REM ==========================================
REM  Pokeka static build + GitHub push
REM  Repo : C:\Users\user\ClimaxGit\climax3
REM  Branch: main
REM ==========================================

cd /d C:\Users\user\ClimaxGit\climax3

echo [1] run python scripts\build_pokeka_static.py
set EXCEL_PATH=C:\Users\user\ClimaxGit\climax3\data\pokeca_rush.xlsm
set SHEET_NAME=Sheet1
set OUT_DIR=docs
python scripts\build_pokeka_static.py

echo.
echo [2] git add (docs only)
git add docs index.html scripts\build_pokeka_static.py

echo.
echo [3] git commit
set MSG=auto: update pokeka static pages %date% %time%
git commit -m "%MSG%"

echo.
echo [4] git push origin main
git push origin main

echo.
echo [OK] done.
pause
