# -*- coding: utf-8 -*-
"""주소입력 팝업을 잠깐 열어 DOM 구조(검색칸·결과링크)를 따고 닫는다. (선택/제출 안 함)"""
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

main = d.current_window_handle
before = set(d.window_handles)

# zipno 근처의 첫 '주소입력' 버튼 클릭
btns = d.find_elements(By.XPATH, "//button[contains(.,'주소입력')]")
print("주소입력 버튼 수:", len(btns))
if not btns:
    raise SystemExit("버튼 없음")
d.execute_script("arguments[0].scrollIntoView({block:'center'});", btns[0])
time.sleep(0.3)
btns[0].click()   # 네이티브 클릭 = 사용자 제스처 → 팝업 차단 회피

# 새 창 대기
popup = None
for _ in range(20):
    time.sleep(0.3)
    new = set(d.window_handles) - before
    if new:
        popup = new.pop()
        break

if not popup:
    print("팝업 창이 안 떴습니다 (레이어형일 수 있음). 메인에서 iframe 탐색.")
    for fr in d.find_elements(By.TAG_NAME, "iframe"):
        print("  iframe src:", fr.get_attribute("src"))
    raise SystemExit

d.switch_to.window(popup)
time.sleep(0.6)
info = {"url": d.current_url, "title": d.title, "inputs": [], "buttons": [], "iframes": []}
for el in d.find_elements(By.XPATH, "//input"):
    info["inputs"].append({"id": el.get_attribute("id"), "name": el.get_attribute("name"),
                           "type": el.get_attribute("type"), "ph": el.get_attribute("placeholder")})
for el in d.find_elements(By.XPATH, "//button | //a[contains(@onclick,'')]")[:15]:
    info["buttons"].append({"tag": el.tag_name, "text": (el.text or "").strip()[:30],
                            "onclick": (el.get_attribute("onclick") or "")[:120]})
for fr in d.find_elements(By.TAG_NAME, "iframe"):
    info["iframes"].append(fr.get_attribute("src"))

path = r"C:\Users\fivep\OneDrive\Desktop\CarMacro\addr_popup.json"
with open(path, "w", encoding="utf-8") as f:
    json.dump(info, f, ensure_ascii=False, indent=2)
print("URL:", info["url"], "| TITLE:", info["title"])
print("inputs:", json.dumps(info["inputs"], ensure_ascii=False))
print("iframes:", info["iframes"])
print("buttons:", json.dumps(info["buttons"], ensure_ascii=False)[:600])

# 팝업 닫고 메인 복귀
d.close()
d.switch_to.window(main)
print("\n팝업 닫고 메인 복귀 완료. written:", path)
