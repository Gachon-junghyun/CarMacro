@echo off
REM ── 신청서 데이터셋 만들기 GUI 실행 ──
REM 폼에 신청자 정보를 입력해 실전 양식 xlsx 를 생성합니다.
REM (크롬/9222 불필요 — 엑셀 생성 전용. 만든 xlsx 는 run_app.bat 로 자동입력)

cd /d "%~dp0"

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" (
  echo [오류] 가상환경 파이썬을 못 찾았습니다: "%PY%"
  echo        먼저 아래로 환경을 만드세요:
  echo          python -m venv .venv
  echo          .venv\Scripts\python -m pip install -r requirements.txt
  pause
  exit /b 1
)

"%PY%" make_gui.py
if errorlevel 1 (
  echo.
  echo [앱이 오류로 종료되었습니다. 위 메시지를 확인하세요.]
  pause
)
