@echo off
chcp 65001 >nul
echo ========================================
echo   Wastewater Log Generator Server
echo   http://localhost:8789
echo ========================================
cd /d "%~dp0"
start http://localhost:8789
python server.py
pause
