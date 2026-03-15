@echo off
chcp 65001 > nul
echo 폐수배출시설 운영일지 생성기 실행 중...

:: Python 경로 자동 탐색
where python > nul 2>&1
if %errorlevel% == 0 (
    python "%~dp0water_app.py"
    goto end
)

where python3 > nul 2>&1
if %errorlevel% == 0 (
    python3 "%~dp0water_app.py"
    goto end
)

:: Python 일반 설치 경로 시도
for %%p in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
) do (
    if exist %%p (
        %%p "%~dp0water_app.py"
        goto end
    )
)

:: Python 없음
echo.
echo [오류] Python이 설치되어 있지 않거나 PATH에 등록되지 않았습니다.
echo.
echo 해결 방법:
echo   1. https://www.python.org/downloads/ 에서 Python 설치
echo   2. 설치 시 "Add Python to PATH" 반드시 체크
echo   3. 설치 후 이 파일 다시 실행
echo.
pause
exit /b 1

:end
