@echo off
chcp 65001 > nul
cd /d "%~dp0"
python capture_buttons.py
pause
