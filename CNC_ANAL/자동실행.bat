@echo off
chcp 65001 > nul
cd /d "%~dp0"
python soltri_auto.py
pause
