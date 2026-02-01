@echo off
chcp 65001 >nul
echo ================================
echo 재고관리 자동화 스크립트
echo ================================
echo.

REM Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo    https://www.python.org/downloads/ 에서 Python을 설치해주세요.
    pause
    exit /b 1
)

REM 필요한 라이브러리 설치 확인
echo 필요한 라이브러리 확인 중...
pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo openpyxl 설치 중...
    pip install openpyxl
)

pip show pandas >nul 2>&1
if errorlevel 1 (
    echo pandas 설치 중...
    pip install pandas
)

echo.
echo ================================
echo 스크립트 실행 중...
echo ================================
echo.

REM 스크립트 실행 (배치 파일과 같은 폴더의 재고관리.xlsx 사용)
python "%~dp0update_joheung_inventory.py" "%~dp0재고관리.xlsx"

echo.
pause
