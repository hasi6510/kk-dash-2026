@echo off
chcp 65001 > nul
echo.
echo ============================================
echo   家計ダッシュボード 更新ツール
echo ============================================
echo.
echo MoneyForward ME から CSV をダウンロードして
echo mf_csv フォルダに入れてから実行してください。
echo.
echo 手順:
echo   1. moneyforward.com にログイン
echo   2. 家計簿 ^> 収支内訳 ^> CSVダウンロード
echo   3. mf_csv フォルダに移動
echo   4. このファイルをダブルクリック
echo.
echo 更新中...
echo.

cd /d "%~dp0"

where node >nul 2>&1
if %errorlevel% neq 0 (
  echo [エラー] Node.js が見つかりません。
  echo https://nodejs.org/ からインストールしてください。
  echo.
  pause
  exit /b 1
)

node update.js

if %errorlevel% neq 0 (
  echo.
  echo [エラー] 更新中に問題が発生しました。
  echo 上記のエラーメッセージを確認してください。
  echo.
  pause
  exit /b 1
)

echo.
echo ============================================
echo   データ更新完了！GitHub にアップロード中...
echo ============================================
echo.

git add data_inline.js
git commit -m "update: %date% のデータを更新"
git push origin main

if %errorlevel% neq 0 (
  echo.
  echo [警告] GitHub へのアップロードに失敗しました。
  echo ネットワーク接続を確認してください。
  echo.
) else (
  echo.
  echo ============================================
  echo   完了！以下のURLでスマホからも確認できます
  echo   https://hasi6510.github.io/kk-dash-2026/
  echo ============================================
)
echo.
pause
