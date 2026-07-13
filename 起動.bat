@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ==================================================
echo   DPS 空調デマンド制御 試算アプリ
echo ==================================================
echo.

rem --- Python を探す（py -3 優先、無ければ python）---
set "PYCMD="
where py >nul 2>nul && set "PYCMD=py -3"
if not defined PYCMD (
  where python >nul 2>nul && set "PYCMD=python"
)
if not defined PYCMD (
  echo [エラー] Python が見つかりませんでした。
  echo   先に Python 3.9 以上をインストールしてください：
  echo   https://www.python.org/downloads/
  echo   ※インストール時「Add Python to PATH」に必ずチェックを入れてください。
  echo.
  pause
  exit /b 1
)

rem --- 必要ライブラリの確認（無ければ初回セットアップ）---
%PYCMD% -c "import streamlit" >nul 2>nul
if errorlevel 1 (
  echo 初回セットアップ中です。必要なライブラリを取得します（数分かかります）...
  echo.
  %PYCMD% -m pip install --upgrade pip
  %PYCMD% -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo [エラー] ライブラリのインストールに失敗しました。ネット接続をご確認ください。
    pause
    exit /b 1
  )
  echo.
  echo セットアップが完了しました。
  echo.
)

echo アプリを起動します。ブラウザが自動で開きます。
echo   ・ブラウザが開かない場合： http://localhost:8501 を開いてください。
echo   ・終了するには、この黒い画面で Ctrl+C を押すか、画面を閉じてください。
echo.
%PYCMD% -m streamlit run app.py

echo.
echo アプリを終了しました。
pause
