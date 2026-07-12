@echo off
REM Launch Chrome with remote debugging on port 9222 for EV/H2 subsidy auto-fill.
REM Close all running Chrome windows first (same profile blocks the port).

set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

if not exist "%CHROME%" (
  echo [ERROR] chrome.exe not found. Edit the path in this script.
  pause
  exit /b 1
)

start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-debug-carmacro"
echo Chrome launched on port 9222. Log in to ev.or.kr and open the application form.