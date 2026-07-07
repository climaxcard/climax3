@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

REM ============================================================
REM POKECA buylist update + SAFE deploy
REM - UTF-8 no BOM recommended
REM - Never sync S3 root
REM - Never use S3 --delete
REM - External JSON mode: upload buylist.json
REM - Embedded HTML mode: upload index.html only
REM ============================================================

set "ROOT=C:\Users\user\ClimaxGit\climax3"
set "SCRIPTS=%ROOT%\scripts"
set "PYTHON_EXE=python"

set "DOCS_DIR=%ROOT%\docs"
set "SCRIPTS_DOCS_DIR=%SCRIPTS%\docs"
set "INDEX_HTML="
set "BUYLIST_JSON="
set "NEEDS_JSON=0"

set "S3_BUCKET=climax-kaitori-static"
set "S3_PREFIX=pokeca"
set "CF_DIST_ID=E51XRDVR8AQAD"
set "PUBLIC_URL=https://kaitori.climax-card.com/pokeca/default/"

set "OUTPUT_DIR=C:\Users\user\OneDrive\ドキュメント\Desktop\ポケカラッシュ"
set "MYCA_CSV_OUT=%OUTPUT_DIR%"
set "MYCA_CSV_NAME=POKECA_Myca_upload.csv"
set "MYCA_CSV_PATH=%MYCA_CSV_OUT%\%MYCA_CSV_NAME%"
set "XLSM="

set "LOG_DIR=%ROOT%\logs"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%" >nul 2>&1
set "LOG_FILE=%LOG_DIR%\safe_pokeca_update_final.log"

> "%LOG_FILE%" echo ===== POKECA SAFE UPDATE START %date% %time% =====

call :CHECK_TOOL "%PYTHON_EXE%" || goto :END_FAIL
call :CHECK_TOOL git || goto :END_FAIL
call :CHECK_TOOL aws || goto :END_FAIL

if not exist "%ROOT%" call :DIE "ROOT not found: %ROOT%"
if not exist "%DOCS_DIR%" mkdir "%DOCS_DIR%" >> "%LOG_FILE%" 2>&1
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%" >> "%LOG_FILE%" 2>&1

call :RESOLVE_XLSM

call :LOG "[0/8] Optional git pull"
cd /d "%ROOT%" || call :DIE "cd failed: %ROOT%"
call :SAFE_GIT_PULL

call :LOG "[0.5/8] Python dependencies"
call :RUN "%PYTHON_EXE%" -m pip install -q beautifulsoup4 pandas openpyxl requests pillow lxml playwright

call :LOG "[1/8] Main update scripts"
cd /d "%ROOT%" || call :DIE "cd failed: %ROOT%"
if not exist "%SCRIPTS%\scrape_cardrush_and_update.py" call :DIE "Missing required script: %SCRIPTS%\scrape_cardrush_and_update.py"
call :RUN "%PYTHON_EXE%" "%SCRIPTS%\scrape_cardrush_and_update.py"

call :LOG "[2/8] Optional update scripts"
cd /d "%ROOT%" || call :DIE "cd failed: %ROOT%"
if exist "%SCRIPTS%\generate_shindan_buylist_png_only.py" (
  call :RUN "%PYTHON_EXE%" "%SCRIPTS%\generate_shindan_buylist_png_only.py"
) else (
  call :LOG "[WARN] Missing optional script. Skip: generate_shindan_buylist_png_only.py"
)

call :LOG "[3/8] Build static pages"
set "BUILD_DONE="
if not defined BUILD_DONE if exist "%SCRIPTS%\build_pokeka_static.py" (
  call :RUN "%PYTHON_EXE%" "%SCRIPTS%\build_pokeka_static.py"
  set "BUILD_DONE=1"
)
if not defined BUILD_DONE if exist "%ROOT%\generate_buylist.py" (
  call :RUN "%PYTHON_EXE%" "%ROOT%\generate_buylist.py"
  set "BUILD_DONE=1"
)
if not defined BUILD_DONE if exist "%ROOT%\gen_buylist.py" (
  call :RUN "%PYTHON_EXE%" "%ROOT%\gen_buylist.py"
  set "BUILD_DONE=1"
)
if not defined BUILD_DONE call :DIE "No build script found or build did not run."

call :LOG "[4/8] Normalize build output"
if exist "%SCRIPTS_DOCS_DIR%\default\index.html" (
  call :LOG "[INFO] scripts\docs output found. Copy to root docs."
  robocopy "%SCRIPTS_DOCS_DIR%" "%DOCS_DIR%" /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /NFL /NDL /NJH /NJS >> "%LOG_FILE%" 2>&1
  set "ROBO_RC=!ERRORLEVEL!"
  if !ROBO_RC! GEQ 8 call :DIE "robocopy scripts\docs to docs failed RC=!ROBO_RC!"
) else (
  call :LOG "[INFO] scripts\docs output not found. Use root docs."
)

call :LOG "[5/8] Verify outputs and detect data mode"
call :RESOLVE_INDEX_HTML
call :DETECT_DATA_MODE

call :LOG "[6/8] Git commit and push best effort"
cd /d "%ROOT%" || call :DIE "cd failed: %ROOT%"
git add docs >> "%LOG_FILE%" 2>&1
git add "%~nx0" >> "%LOG_FILE%" 2>&1
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "update pokeca buylist" >> "%LOG_FILE%" 2>&1
  if errorlevel 1 call :LOG "[WARN] git commit failed. Continue."
  git pull --rebase --autostash >> "%LOG_FILE%" 2>&1
  if errorlevel 1 call :LOG "[WARN] git pull after commit failed. Continue."
  git push >> "%LOG_FILE%" 2>&1
  if errorlevel 1 call :LOG "[WARN] git push failed. Continue."
) else (
  call :LOG "[INFO] No git changes."
)

call :LOG "[7/8] Safe S3 upload"
call :DEPLOY_SAFE

echo.
echo [OK] Done: %PUBLIC_URL%
echo [LOG] %LOG_FILE%
pause
exit /b 0


REM ============================================================
REM Functions
REM ============================================================

:RESOLVE_XLSM
set "XLSM="
for %%X in (
  "%ROOT%\data\pokeca_rush.xlsm"
  "%ROOT%\buylist.xlsm"
  "%ROOT%\data\buylist.xlsm"
) do (
  if not defined XLSM (
    if exist "%%~fX" set "XLSM=%%~fX"
  )
)
if defined XLSM (
  call :LOG "[INFO] XLSM=%XLSM%"
) else (
  call :LOG "[WARN] XLSM not found from candidates. Continue if scripts do not require it."
)
exit /b 0


:RESOLVE_INDEX_HTML
set "INDEX_HTML="
for %%I in (
  "%DOCS_DIR%\default\index.html"
  "%DOCS_DIR%\default\default\index.html"
  "%SCRIPTS_DOCS_DIR%\default\index.html"
  "%SCRIPTS_DOCS_DIR%\default\default\index.html"
) do (
  if not defined INDEX_HTML (
    if exist "%%~fI" set "INDEX_HTML=%%~fI"
  )
)
if not defined INDEX_HTML call :DIE "index.html not found in known output locations"
call :LOG "[INFO] INDEX_HTML=%INDEX_HTML%"
exit /b 0


:DETECT_DATA_MODE
set "BUYLIST_JSON="
set "NEEDS_JSON=0"

findstr /i "__BUYLIST_API__ buylist.json" "%INDEX_HTML%" >nul 2>&1
if not errorlevel 1 set "NEEDS_JSON=1"

for %%J in (
  "%DOCS_DIR%\buylist.json"
  "%DOCS_DIR%\default\buylist.json"
  "%DOCS_DIR%\default\default\buylist.json"
  "%ROOT%\buylist.json"
  "%SCRIPTS_DOCS_DIR%\buylist.json"
  "%SCRIPTS_DOCS_DIR%\default\buylist.json"
  "%SCRIPTS_DOCS_DIR%\default\default\buylist.json"
) do (
  if not defined BUYLIST_JSON (
    if exist "%%~fJ" set "BUYLIST_JSON=%%~fJ"
  )
)

if defined BUYLIST_JSON (
  call :LOG "[INFO] External JSON mode. BUYLIST_JSON=%BUYLIST_JSON%"
  call :FIX_INDEX_API
  exit /b 0
)

if "%NEEDS_JSON%"=="1" (
  call :DIE "index.html requires buylist.json, but buylist.json was not found."
)

call :LOG "[INFO] Embedded HTML mode. buylist.json not required."
exit /b 0


:FIX_INDEX_API
call :LOG "[INFO] Fix index.html JSON path to ../buylist.json"
"%PYTHON_EXE%" -c "from pathlib import Path; import re, os; p=Path(os.environ['INDEX_HTML']); s=p.read_text(encoding='utf-8'); s=re.sub(r'window\.__BUYLIST_API__\s*=\s*[\x22\x27][^\x22\x27]+[\x22\x27];','window.__BUYLIST_API__=\x22../buylist.json\x22;',s); p.write_text(s,encoding='utf-8')" >> "%LOG_FILE%" 2>&1
if errorlevel 1 call :DIE "Failed to fix JSON path in index.html"
exit /b 0


:DEPLOY_SAFE
aws sts get-caller-identity >> "%LOG_FILE%" 2>&1
if errorlevel 1 call :DIE "AWS credential invalid. Run aws configure."

REM Upload only exact public files. No S3 root sync. No --delete.
aws s3 cp "%INDEX_HTML%" "s3://%S3_BUCKET%/%S3_PREFIX%/default/index.html" --cache-control "no-cache, no-store, must-revalidate" --content-type "text/html; charset=utf-8" >> "%LOG_FILE%" 2>&1
if errorlevel 1 call :DIE "S3 upload failed: default/index.html"

if defined BUYLIST_JSON (
  aws s3 cp "%BUYLIST_JSON%" "s3://%S3_BUCKET%/%S3_PREFIX%/buylist.json" --cache-control "no-cache, no-store, must-revalidate" --content-type "application/json; charset=utf-8" >> "%LOG_FILE%" 2>&1
  if errorlevel 1 call :DIE "S3 upload failed: buylist.json"
) else (
  call :LOG "[INFO] No buylist.json. Skip JSON upload."
)

if exist "%DOCS_DIR%\price_asc\index.html" (
  aws s3 cp "%DOCS_DIR%\price_asc\index.html" "s3://%S3_BUCKET%/%S3_PREFIX%/price_asc/index.html" --cache-control "no-cache, no-store, must-revalidate" --content-type "text/html; charset=utf-8" >> "%LOG_FILE%" 2>&1
  if errorlevel 1 call :DIE "S3 upload failed: price_asc/index.html"
)

if exist "%DOCS_DIR%\price_desc\index.html" (
  aws s3 cp "%DOCS_DIR%\price_desc\index.html" "s3://%S3_BUCKET%/%S3_PREFIX%/price_desc/index.html" --cache-control "no-cache, no-store, must-revalidate" --content-type "text/html; charset=utf-8" >> "%LOG_FILE%" 2>&1
  if errorlevel 1 call :DIE "S3 upload failed: price_desc/index.html"
)

REM Safe asset upload: limited to this category prefix and no --delete.
if exist "%DOCS_DIR%\assets" (
  aws s3 sync "%DOCS_DIR%\assets" "s3://%S3_BUCKET%/%S3_PREFIX%/assets/" --size-only --cache-control "public,max-age=31536000,immutable" --only-show-errors >> "%LOG_FILE%" 2>&1
  if errorlevel 1 call :DIE "S3 upload failed: assets"
)

aws cloudfront create-invalidation --distribution-id "%CF_DIST_ID%" --paths "/%S3_PREFIX%/default/*" "/%S3_PREFIX%/buylist.json" "/%S3_PREFIX%/price_asc/*" "/%S3_PREFIX%/price_desc/*" "/%S3_PREFIX%/assets/*" >> "%LOG_FILE%" 2>&1
if errorlevel 1 call :DIE "CloudFront invalidation failed"

exit /b 0


:SAFE_GIT_PULL
git rev-parse --is-inside-work-tree >> "%LOG_FILE%" 2>&1
if errorlevel 1 exit /b 0

git diff --quiet
set "D1=%ERRORLEVEL%"
git diff --cached --quiet
set "D2=%ERRORLEVEL%"

if not "%D1%%D2%"=="00" (
  call :LOG "[WARN] Working tree has local changes. Skip git pull."
  exit /b 0
)

git pull --rebase --autostash >> "%LOG_FILE%" 2>&1
if errorlevel 1 call :LOG "[WARN] git pull failed. Continue."

exit /b 0


:CHECK_TOOL
where %~1 >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Tool not found: %~1
  echo [ERROR] Tool not found: %~1>> "%LOG_FILE%"
  exit /b 1
)
exit /b 0


:RUN
echo [RUN] %*>> "%LOG_FILE%"
%* >> "%LOG_FILE%" 2>&1
if errorlevel 1 call :DIE "Command failed: %*"
exit /b 0


:LOG
echo %~1
echo %~1>> "%LOG_FILE%"
exit /b 0


:DIE
echo.
echo [ERROR] %~1
echo [ERROR] %~1>> "%LOG_FILE%"
echo [LOG] %LOG_FILE%
start "" notepad "%LOG_FILE%"
pause
exit /b 1


:END_FAIL
echo.
echo [ERROR] Setup failed.
echo [LOG] %LOG_FILE%
pause
exit /b 1
