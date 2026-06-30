#!/usr/bin/env bash
# ── EV/H2 보조금 자동입력용 크롬을 디버그 포트 9222로 띄웁니다 (macOS) ──
# 이미 떠 있는 크롬이 같은 프로필을 쓰면 포트가 안 열립니다.
# 이 스크립트는 별도 프로필(~/chrome-debug-carmacro)을 쓰므로
# 평소 쓰는 크롬과 분리되어 함께 떠 있어도 됩니다.
set -euo pipefail

CHROME="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
if [ ! -x "$CHROME" ]; then
  CHROME="$HOME/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
fi

if [ ! -x "$CHROME" ]; then
  echo "[오류] Google Chrome 을 못 찾았습니다. CHROME 경로를 직접 수정하세요." >&2
  exit 1
fi

PROFILE="$HOME/chrome-debug-carmacro"

echo "크롬을 9222 포트로 띄웁니다 (프로필: $PROFILE)"
echo "이 크롬에서 ev.or.kr 로그인 후 신청서 작성 폼까지 진입하세요."
"$CHROME" --remote-debugging-port=9222 --user-data-dir="$PROFILE" >/dev/null 2>&1 &

echo "완료. 크롬 창이 떴는지 확인하세요. (포트 9222)"
