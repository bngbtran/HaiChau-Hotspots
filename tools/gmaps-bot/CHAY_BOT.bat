@echo off
chcp 65001 >nul
cd /d "%~dp0"
where node >nul 2>nul || (echo Chua cai Node.js. Tai ban LTS tai https://nodejs.org roi chay lai. & pause & exit /b)
if not exist node_modules (echo Dang cai thu vien lan dau, cho 1-2 phut... & call npm install --silent)
node bot.js run %*
pause
