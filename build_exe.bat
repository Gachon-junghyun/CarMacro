@echo off
REM ── app.py 를 단일 실행파일(exe)로 빌드 ──
REM 결과: dist\CarMacro-app.exe  (파이썬·selenium 통째로 포함 → 타겟 PC에 파이썬/pip 불필요)
REM 타겟 PC 요건: Windows + Chrome 설치 + 인터넷(최초 실행 시 chromedriver 자동 다운로드)
REM pandas/numpy 는 제외(세로형 양식은 불필요) → 가볍고 안정적.

cd /d "%~dp0"

python -m pip install pyinstaller
python -m PyInstaller --onefile --windowed --name CarMacro-app ^
  --collect-all selenium ^
  --exclude-module pandas --exclude-module numpy --exclude-module matplotlib ^
  --noconfirm app.py

echo.
echo ============================================================
echo  빌드 완료: dist\CarMacro-app.exe
echo  이 exe 하나만 다른 PC로 복사하면 됩니다(파이썬 설치 불필요).
echo  함께 챙길 것: run_chrome.bat, 신청서 xlsx, 첨부 PDF
echo ============================================================
pause
