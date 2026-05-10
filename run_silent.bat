@echo off
chcp 65001 > nul
cd /d "%~dp0"
where node >nul 2>&1
if %errorlevel% neq 0 exit /b 1
node update.js
if %errorlevel% neq 0 exit /b 1
git add data_inline.js
git commit -m "auto-update: %date%"
git push origin main
exit /b %errorlevel%
