@echo off
REM -- 신청서 데이터셋 만들기 GUI 실행 --
REM 폼에 신청자 정보를 입력해 실전 양식 xlsx 를 생성합니다.
REM (크롬/9222 불필요 -- 엑셀 생성 전용. 만든 xlsx 는 run_app.bat 로 자동입력)

cd /d "%~dp0"

REM 1) 가상환경(.venv)이 있으면 우선 사용, 없으면 PATH 의 python 사용
set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

REM 2) 파이썬 자체가 있는지 확인
"%PY%" --version >nul 2>&1
if errorlevel 1 (
  echo [오류] 파이썬을 찾지 못했습니다: "%PY%"
  echo        Python 3.11+ 설치 후 다시 실행하거나, 가상환경을 만드세요:
  echo          python -m venv .venv
  echo          .venv\Scripts\python -m pip install -r requirements.txt
  pause
  exit /b 1
)

REM 3) 필수 패키지가 있으면 그대로 실행, 없을 때만 설치 시도
"%PY%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo [안내] 필요한 패키지가 없어 설치를 시도합니다...
  "%PY%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo [오류] 자동 설치 실패. 직접 실행하세요:
    echo          "%PY%" -m pip install -r requirements.txt
    pause
    exit /b 1
  )
)

"%PY%" make_gui.py
if errorlevel 1 (
  echo.
  echo [앱이 오류로 종료되었습니다. 위 메시지를 확인하세요.]
  pause
)
