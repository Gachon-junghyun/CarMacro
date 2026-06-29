@echo off
REM ── EV/H2 보조금 자동입력용 크롬을 디버그 포트 9222로 띄웁니다 ──
REM 이미 떠 있는 크롬이 있으면 모두 끈 뒤 실행하세요(같은 프로필이면 포트가 안 열립니다).

set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if not exist "%CHROME%" (
  echo [오류] chrome.exe 를 못 찾았습니다. 경로를 직접 수정하세요.
  pause
  exit /b 1
)

start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-debug-carmacro"
echo 크롬을 9222 포트로 띄웠습니다. 이 창에서 ev.or.kr 로그인 후 신청서 폼까지 진입하세요.
