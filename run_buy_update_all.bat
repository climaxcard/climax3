@echo off
chcp 65001 > nul
setlocal EnableExtensions EnableDelayedExpansion

REM ==================================================
REM 設定
REM ==================================================
set "SCRIPTS=C:\Users\user\ClimaxGit\climax3\scripts"
set "ROOT=C:\Users\user\ClimaxGit\climax3"

REM PNG / CSV 保存先
set "OUTPUT_DIR=C:\Users\user\OneDrive\ドキュメント\Desktop\ポケカラッシュ"

REM PNG設定
set "PNG_OUT=%OUTPUT_DIR%"
set "PNG_BASE=ポケカ新弾買取表"

REM Excel
set "XLSM=C:\Users\user\ClimaxGit\climax3\data\pokeca_rush.xlsm"

REM MYCAアップロード用CSV
set "MYCA_CSV_OUT=%OUTPUT_DIR%"
set "MYCA_CSV_NAME=ポケカラッシュ_Mycaアップロード用.csv"
set "MYCA_CSV_PATH=%MYCA_CSV_OUT%\%MYCA_CSV_NAME%"

REM ==================================================
REM まずはROOTへ移動して最新化（push拒否対策）
REM ==================================================
cd /d "%ROOT%"
if errorlevel 1 (
  echo [ERROR] ルートへ移動できません: %ROOT%
  pause
  exit /b 1
)

echo =========================================
echo [0/9] git pull --rebase（push拒否対策）
echo =========================================
git pull --rebase --autostash
if errorlevel 1 (
  echo.
  echo [ERROR] git pull --rebase でエラー（競合の可能性）
  echo        git status を見てコンフリクト解消してから再実行してください
  pause
  exit /b 1
)

REM ==================================================
REM scriptsへ移動
REM ==================================================
cd /d "%SCRIPTS%"
if errorlevel 1 (
  echo [ERROR] scriptsへ移動できません: %SCRIPTS%
  pause
  exit /b 1
)

echo.
echo =========================================
echo [1/9] CardRush データ取得・Excel更新
echo =========================================
python scrape_cardrush_and_update.py
if errorlevel 1 (
  echo.
  echo [ERROR] scrape_cardrush_and_update.py でエラー
  pause
  exit /b 1
)

echo.
echo =========================================
echo [2/9] MYCAアップロード用CSV生成
echo =========================================

if not exist "%MYCA_CSV_OUT%" (
  mkdir "%MYCA_CSV_OUT%"
)

set "CSV_PY=%TEMP%\export_myca_csv.py"

> "%CSV_PY%" echo import csv
>> "%CSV_PY%" echo import openpyxl
>> "%CSV_PY%" echo from pathlib import Path
>> "%CSV_PY%" echo xlsm = Path(r"%XLSM%")
>> "%CSV_PY%" echo out = Path(r"%MYCA_CSV_PATH%")
>> "%CSV_PY%" echo out.parent.mkdir(parents=True, exist_ok=True)
>> "%CSV_PY%" echo if not xlsm.exists():
>> "%CSV_PY%" echo     raise FileNotFoundError(f"Excel file not found: {xlsm}")
>> "%CSV_PY%" echo wb = openpyxl.load_workbook(xlsm, data_only=True, keep_vba=True)
>> "%CSV_PY%" echo if "Sheet1" not in wb.sheetnames:
>> "%CSV_PY%" echo     raise RuntimeError(f"Sheet1 not found. sheets={wb.sheetnames}")
>> "%CSV_PY%" echo ws = wb["Sheet1"]
>> "%CSV_PY%" echo with out.open("w", newline="", encoding="utf-8-sig") as f:
>> "%CSV_PY%" echo     writer = csv.writer(f)
>> "%CSV_PY%" echo     for row in ws.iter_rows():
>> "%CSV_PY%" echo         values = []
>> "%CSV_PY%" echo         for cell in row:
>> "%CSV_PY%" echo             v = cell.value
>> "%CSV_PY%" echo             values.append("" if v is None else v)
>> "%CSV_PY%" echo         writer.writerow(values)
>> "%CSV_PY%" echo print("[OK] MYCA CSV saved:", out)

python "%CSV_PY%"
if errorlevel 1 (
  echo.
  echo [ERROR] MYCAアップロード用CSV生成でエラー
  del "%CSV_PY%" 2>nul
  pause
  exit /b 1
)

del "%CSV_PY%" 2>nul

echo.
echo [OK] MYCA CSV保存先:
echo %MYCA_CSV_PATH%

echo.
echo =========================================
echo [3/9] Excel分類チェック（メガゲッコウガ確認）
echo =========================================

set "CHECK_PY=%TEMP%\check_mega_gekkouga.py"

> "%CHECK_PY%" echo import openpyxl
>> "%CHECK_PY%" echo from pathlib import Path
>> "%CHECK_PY%" echo p = Path(r"%XLSM%")
>> "%CHECK_PY%" echo wb = openpyxl.load_workbook(p, data_only=True, keep_vba=True)
>> "%CHECK_PY%" echo ws = wb["Sheet1"]
>> "%CHECK_PY%" echo found = False
>> "%CHECK_PY%" echo print("--- メガゲッコウガ rows ---")
>> "%CHECK_PY%" echo for r in range(3, ws.max_row + 1):
>> "%CHECK_PY%" echo     name = str(ws.cell(r, 3).value or "")
>> "%CHECK_PY%" echo     if "メガゲッコウガ" in name:
>> "%CHECK_PY%" echo         found = True
>> "%CHECK_PY%" echo         print("row=" + str(r) + " C=" + str(ws.cell(r, 3).value) + " E=" + str(ws.cell(r, 5).value) + " O=" + str(ws.cell(r, 15).value) + " Q=" + str(ws.cell(r, 17).value))
>> "%CHECK_PY%" echo if not found:
>> "%CHECK_PY%" echo     print("該当なし")
>> "%CHECK_PY%" echo else:
>> "%CHECK_PY%" echo     print("--- check end ---")

python "%CHECK_PY%"
if errorlevel 1 (
  echo.
  echo [ERROR] Excel分類チェックでエラー
  del "%CHECK_PY%" 2>nul
  pause
  exit /b 1
)

del "%CHECK_PY%" 2>nul

REM ==================================================
REM scriptsへ戻る
REM ==================================================
cd /d "%SCRIPTS%"
if errorlevel 1 (
  echo [ERROR] scriptsへ移動できません: %SCRIPTS%
  pause
  exit /b 1
)

echo.
echo =========================================
echo [4/9] ポケカ買取表 静的HTML生成
echo =========================================
python build_pokeka_static.py
if errorlevel 1 (
  echo.
  echo [ERROR] build_pokeka_static.py でエラー
  pause
  exit /b 1
)

echo.
echo =========================================
echo [5/9] docs 同期（安全上書き） scripts\docs -^> root\docs
echo =========================================
set "SRC=%SCRIPTS%\docs"
set "DST=%ROOT%\docs"

if not exist "%SRC%" (
  echo [ERROR] コピー元が存在しません: %SRC%
  pause
  exit /b 1
)

if not exist "%DST%" (
  mkdir "%DST%"
)

REM 削除しない。上書きのみ。
robocopy "%SRC%" "%DST%" /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /NFL /NDL /NJH /NJS
set "RC=%errorlevel%"

REM robocopyの戻り値は特殊。0-7は概ね成功。
if %RC% GEQ 8 (
  echo [ERROR] robocopy failed with code %RC%
  pause
  exit /b 1
)

echo.
echo =========================================
echo [6/9] 古いPNG削除
echo =========================================
if not exist "%PNG_OUT%" (
  mkdir "%PNG_OUT%"
)

echo 削除対象: "%PNG_OUT%\%PNG_BASE%-*.png"
del /q "%PNG_OUT%\%PNG_BASE%-*.png" 2>nul

echo OK: 古いPNGを削除しました

echo.
echo =========================================
echo [7/9] ポケカ新弾買取表 PNG生成
echo =========================================
cd /d "%SCRIPTS%"
if errorlevel 1 (
  echo [ERROR] scriptsへ移動できません: %SCRIPTS%
  pause
  exit /b 1
)

python generate_shindan_buylist_png_only.py --download-images
if errorlevel 1 (
  echo.
  echo [ERROR] generate_shindan_buylist_png_only.py でエラー
  pause
  exit /b 1
)

echo.
echo =========================================
echo [8/9] git add docs
echo =========================================
cd /d "%ROOT%"
if errorlevel 1 (
  echo [ERROR] ルートへ移動できません: %ROOT%
  pause
  exit /b 1
)

git add docs
if errorlevel 1 (
  echo [ERROR] git add でエラー
  pause
  exit /b 1
)

REM bat自体もGit管理している場合は反映
if exist "%~nx0" (
  git add "%~nx0" 2>nul
)

REM 変更がないなら終了
git diff --cached --quiet
if %errorlevel%==0 (
  echo.
  echo [INFO] docs に変更がありません（commit/push不要）
  echo [INFO] PNG は %PNG_OUT% に出力済みです
  echo [INFO] MYCA CSV は %MYCA_CSV_PATH% に出力済みです
  pause
  exit /b 0
)

echo.
echo =========================================
echo [9/9] git commit / push
echo =========================================

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set "D=%%a-%%b-%%c"
for /f "tokens=1-2 delims=: " %%a in ("%time%") do set "T=%%a:%%b"
set "MSG=update pokeca %D% %T%"

git commit -m "%MSG%"
if errorlevel 1 (
  echo [ERROR] git commit でエラー
  pause
  exit /b 1
)

git push
if errorlevel 1 (
  echo.
  echo [ERROR] git push でエラー
  echo        もう一度このバッチを実行してください（pull-^>rebaseしてからpushします）
  pause
  exit /b 1
)

echo.
echo ✅ publish 完了！（pull/rebase -^> CSV生成 -^> PNG生成 -^> commit -^> push 済み）
echo ✅ PNG保存先: %PNG_OUT%
echo ✅ MYCA CSV保存先: %MYCA_CSV_PATH%
pause
endlocal
exit /b 0