@echo off
chcp 65001 > nul
echo ================================
echo  CNC 자동화 패키지 설치
echo ================================
echo.
pip install pyautogui keyboard pyperclip pygetwindow pywin32 openpyxl
echo.
echo ================================
echo  설치 완료!
echo  다음 순서:
echo  1. soltri_config.json 열어서 비밀번호/경로 수정
echo  2. SOLTRI-1 실행 후 좌표세팅.bat 실행
echo  3. 자동실행.bat 으로 사용
echo ================================
pause
