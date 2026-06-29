# -*- coding: utf-8 -*-
"""현재 크롬(9222) 폼의 select 옵션 라벨과 라디오 라벨을 UTF-8 파일로 덤프."""
import json
from selenium import webdriver
from selenium.webdriver.common.by import By

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

out = {"url": d.current_url, "selects": {}, "radios": {}, "texts": []}

# select: id -> [(value, label)]
for el in d.find_elements(By.TAG_NAME, "select"):
    sid = el.get_attribute("id") or el.get_attribute("name")
    if not sid:
        continue
    opts_list = []
    for o in el.find_elements(By.TAG_NAME, "option"):
        opts_list.append([o.get_attribute("value"), o.text.strip()])
    out["selects"][sid] = opts_list

# radio: name -> [(value, 옆 라벨 텍스트)]
seen = set()
for el in d.find_elements(By.XPATH, "//input[@type='radio']"):
    name = el.get_attribute("name")
    if not name:
        continue
    val = el.get_attribute("value")
    rid = el.get_attribute("id")
    # 라벨 텍스트: for=id 인 label, 없으면 부모/형제 텍스트
    label = ""
    try:
        if rid:
            labs = d.find_elements(By.XPATH, f"//label[@for='{rid}']")
            if labs:
                label = labs[0].text.strip()
    except Exception:
        pass
    out["radios"].setdefault(name, []).append([rid, val, label])

# 주요 텍스트 input id 목록(있는지 확인용)
for el in d.find_elements(By.XPATH, "//input[@type='text'] | //textarea"):
    tid = el.get_attribute("id")
    if tid:
        out["texts"].append(tid)

path = r"C:\Users\fivep\OneDrive\Desktop\CarMacro\labels.json"
with open(path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print("written:", path)
