@echo off
REM ── CarMacro 자동입력 앱(GUI) 실행 ──
REM 먼저 run_chrome.bat 로 크롬(포트 9222) 띄우고 ev.or.kr 신청서 폼까지 진입해 두세요.

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

"%PY%" app.py
if errorlevel 1 (
  echo.
  echo [앱이 오류로 종료되었습니다. 위 메시지를 확인하세요.]
  pause
)
