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
echo   3. Excel で開き「CSV UTF-8」で保存し直す
echo   4. mf_csv フォルダに移動
echo   5. このファイルをダブルクリック
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
echo   完了！index.html をブラウザで開いてください
echo ============================================
echo.
pause
