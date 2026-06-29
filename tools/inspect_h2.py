# -*- coding: utf-8 -*-
"""열린 모든 창/탭을 훑어 신청서 폼(전기/수소)의 필드 구성을 비교."""
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
import app  # FIELDS 재사용

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

EXPECT_IDS = [f["id"] for f in app.FIELDS] + ["local_nm", "zipno", "addr"]

report = []
for h in d.window_handles:
    try:
        d.switch_to.window(h)
    except Exception:
        continue
    url = d.current_url
    if "sellerApplyform" not in url and "Applyform" not in url:
        report.append({"url": url, "form": False})
        continue
    # 필드 존재 여부
    present, missing = [], []
    for fid in EXPECT_IDS:
        if d.execute_script("return !!document.getElementById(arguments[0]);", fid):
            present.append(fid)
        else:
            missing.append(fid)
    # model_cd 옵션, 라디오 그룹, car_type
    models = d.execute_script(
        "var e=document.getElementById('model_cd');"
        "return e?[...e.options].map(o=>[o.value,(o.textContent||'').trim()]):[];")
    radios = d.execute_script(
        "var s=new Set();document.querySelectorAll('input[type=radio]').forEach(r=>{if(r.name)s.add(r.name)});"
        "return [...s];")
    car_type = d.execute_script(
        "var e=document.querySelector('[name=car_type]');return e?e.value:'';")
    title_step = d.execute_script(
        "var e=document.querySelector('.on, .active, .step.on');return e?e.textContent.trim().slice(0,20):'';")
    report.append({"url": url, "form": True, "car_type": car_type,
                   "present": present, "missing": missing,
                   "model_count": len([m for m in models if m[0]]),
                   "model_sample": [m for m in models if m[0]][:4],
                   "radios": radios})

with open(r"C:\Users\fivep\OneDrive\Desktop\CarMacro\h2.json", "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

for r in report:
    print("URL:", r["url"])
    if not r.get("form"):
        print("  (신청서 폼 아님)\n")
        continue
    print("  car_type:", r["car_type"], "| model수:", r["model_count"])
    print("  있는 필드:", len(r["present"]), "/", len(EXPECT_IDS))
    print("  없는 필드:", r["missing"])
    print("  라디오 그룹:", r["radios"])
    print("  차종 샘플:", r["model_sample"])
    print()
