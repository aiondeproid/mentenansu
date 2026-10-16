@echo off
rem 設備保全アプリを起動する。ダブルクリックで実行
rem この画面を閉じるとアプリが止まる
title 設備保全アプリ（この画面を閉じると止まります）
cd /d "%~dp0"

rem 5000 番は別のアプリ（管理表）が使っているので 5055 にする
if "%MAINT_PORT%"=="" set MAINT_PORT=5055

where python >nul 2>nul
if errorlevel 1 (
  echo [エラー] Python が見つかりません。Python 3.12 をインストールしてください。
  echo         インストール時に「Add python.exe to PATH」にチェックを入れること。
  pause
  exit /b 1
)

if not exist "data\maintenance.db" (
  echo [エラー] data\maintenance.db がありません。
  echo         最初に一度だけ、次のコマンドで Excel の履歴を取り込んでください。
  echo           pip install -r requirements.txt
  echo           python import_xls.py
  pause
  exit /b 1
)

rem 同じポートを別のアプリが使っていたら止める（Windows では二重に起動できてしまうため）
powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort %MAINT_PORT% -State Listen -ErrorAction SilentlyContinue) { exit 1 }"
if errorlevel 1 (
  echo [エラー] ポート %MAINT_PORT% は別のアプリが使っています。
  echo         設備保全アプリがもう起動していないか確認してください。
  echo         別のポートで動かすときは、環境変数 MAINT_PORT に番号を設定してから実行します。
  pause
  exit /b 1
)

echo ============================================================
echo  設備保全アプリを起動します
echo.
echo  このPCで開く    : http://localhost:%MAINT_PORT%/
echo  現場の端末で開く: http://（下のIPv4アドレス）:%MAINT_PORT%/
ipconfig | findstr /c:"IPv4"
echo.
echo  止めるときは この画面で Ctrl + C を押すか、画面を閉じてください
echo ============================================================
echo.

rem 起動を少し待ってからブラウザを開く
start "" /min cmd /c "timeout /t 3 /nobreak >nul & start "" http://localhost:%MAINT_PORT%/"

python app.py

echo.
echo アプリが停止しました。
pause
