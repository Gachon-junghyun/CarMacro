# -*- coding: utf-8 -*-
"""app.py 범용 흐름 통합 테스트: 세로형 xlsx 파싱 → 라이브 폼 입력 → 폐차정보 → 검증.
load_vertical 과 do_fill 의 핵심 경로를 그대로 재현. 저장/제출/보안코드 안 함.
사용: python tools/test_app_flow.py [세로형.xlsx]  (기본: examples/real_sample_ev.xlsx)"""
import os
import sys
import time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from selenium import webdriver
from parser import parse_vertical
from app import (fill_one, fill_address, fill_text, fill_exchange_scrap,
                 verify_fill, verify_exchange_scrap)

XLSX = sys.argv[1] if len(sys.argv) > 1 else "examples/real_sample_ev.xlsx"
print("XLSX:", XLSX)
res = parse_vertical(XLSX)
row = dict(res["mapped"])
ts = set()
for k, v in res["special"].items():
    row[k] = v
    ts.add(k)
row["_text_select"] = ts
print("row keys:", sorted(row.keys()))


def log(m, color=None):
    print(("[R]" if color == "red" else "[G]" if color == "green" else "   ") + " " + m)


o = webdriver.ChromeOptions()
o.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
o.set_capability("unhandledPromptBehavior", "ignore")
d = webdriver.Chrome(options=o)
for h in d.window_handles:
    d.switch_to.window(h)
    if d.current_url.endswith("sellerApplyform"):
        break

print("\n== fill_one ==")
fill_one(d, row, log=log)

print("\n== 주소 ==")
if str(row.get("addr") or "").strip():
    st = fill_address(d, row.get("addr"), log=log)
    if st == "ok":
        fill_text(d, "addr_detail", row.get("addr_detail"))

print("\n== 전환지원금 폐차정보 ==")
if str(row.get("exchange_3year_yn") or "").strip() == "Y" and row.get("exchange_scrap"):
    fill_exchange_scrap(d, row, log=log)

print("\n== 검증(되읽기) ==")
time.sleep(0.4)
mm = verify_fill(d, row) + verify_exchange_scrap(d, row)
if not mm:
    print("  ✅ 전 항목 일치 (불일치 0건)")
else:
    print("  불일치 %d건:" % len(mm))
    for fid, want, got in mm:
        print("    · %s: 기대 '%s' / 실제 '%s'" % (fid, want, got))
print("\n(저장/제출/보안코드 안 함.)")
