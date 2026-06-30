#!/usr/bin/env bash
# ── CarMacro 자동입력 앱(GUI) 실행 (macOS) ──
# 먼저 ./run_chrome.sh 로 크롬(포트 9222) 띄우고 ev.or.kr 신청서 폼까지 진입해 두세요.
set -euo pipefail

# 스크립트가 있는 폴더로 이동
cd "$(dirname "$0")"

PY=".venv/bin/python"
if [ ! -x "$PY" ]; then
  echo "[오류] 가상환경 파이썬을 못 찾았습니다: $PY" >&2
  echo "       먼저 아래로 환경을 만드세요 (tkinter 포함된 python 필요):" >&2
  echo "         /opt/homebrew/bin/python3.13 -m venv .venv" >&2
  echo "         .venv/bin/python -m pip install -r requirements.txt" >&2
  exit 1
fi

# tkinter 사용 가능 여부 사전 점검(없으면 GUI가 안 뜸)
if ! "$PY" -c "import tkinter" 2>/dev/null; then
  echo "[오류] 이 가상환경의 파이썬에 tkinter 가 없습니다." >&2
  echo "       tkinter 포함된 파이썬(예: brew의 python3.13)으로 .venv 를 다시 만드세요:" >&2
  echo "         rm -rf .venv && /opt/homebrew/bin/python3.13 -m venv .venv" >&2
  echo "         .venv/bin/python -m pip install -r requirements.txt" >&2
  exit 1
fi

"$PY" app.py
