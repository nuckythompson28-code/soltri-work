@echo off
chcp 65001 > nul
cd /d "%~dp0"
python find_coords.py
pause
