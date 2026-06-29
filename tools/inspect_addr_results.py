# -*- coding: utf-8 -*-
"""주소 팝업에서 검색만 실행해 결과 목록 구조와 개수를 확인 (선택/제출 안 함)."""
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By

KEYWORD = "세종특별자치시 시청대로 78"

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

main = d.current_window_handle
before = set(d.window_handles)

btns = d.find_elements(By.XPATH, "//button[contains(.,'주소입력')]")
d.execute_script("arguments[0].scrollIntoView({block:'center'});", btns[0])
time.sleep(0.3)
btns[0].click()

popup = None
for _ in range(20):
    time.sleep(0.3)
    new = set(d.window_handles) - before
    if new:
        popup = new.pop()
        break
if not popup:
    raise SystemExit("팝업 안 뜸")

d.switch_to.window(popup)
time.sleep(0.6)

# 검색어 입력 후 검색
kw = d.find_element(By.ID, "keyword")
kw.clear()
kw.send_keys(KEYWORD)
d.execute_script("$('#currentPage').val(1); searchUrlJuso();")
time.sleep(1.5)

# 결과 영역의 a/링크 덤프
anchors = []
for a in d.find_elements(By.XPATH, "//a"):
    oc = a.get_attribute("onclick") or ""
    href = a.get_attribute("href") or ""
    txt = (a.text or "").strip()
    if "juso" in oc.lower() or "Parent" in oc or "javascript" in href.lower():
        if txt or oc:
            anchors.append({"text": txt[:60], "onclick": oc[:160], "href": href[:120]})

# 결과 행(li/tr) 텍스트도
rows = []
for el in d.find_elements(By.XPATH, "//ul//li | //table//tr"):
    t = (el.text or "").strip().replace("\n", " ")
    if "시청대로" in t or t:
        if t:
            rows.append(t[:90])

# 결과 앵커의 정확한 HTML/선택자 파악
detail = d.execute_script(r"""
const res=[];
for (const a of document.querySelectorAll('a')){
  const t=(a.textContent||'').trim();
  if (t.includes('시청대로')){
    res.push({outer: a.outerHTML.slice(0,260),
              parentId: a.closest('[id]') ? a.closest('[id]').id : '',
              cls: a.className});
  }
}
return res;
""")
print("\n=== 결과 앵커 상세 ===")
for x in detail:
    print("  parentId:", x["parentId"], "| class:", x["cls"])
    print("  outer:", x["outer"])

info = {"keyword": KEYWORD, "url": d.current_url, "anchors": anchors,
        "rows_sample": rows[:12], "detail": detail}
path = r"C:\Users\fivep\OneDrive\Desktop\CarMacro\addr_results.json"
with open(path, "w", encoding="utf-8") as f:
    json.dump(info, f, ensure_ascii=False, indent=2)

print("결과 앵커 수:", len(anchors))
for a in anchors[:12]:
    print(f"  text={a['text']!r}")
    print(f"    onclick={a['onclick']!r}")
print("\n결과행 샘플:")
for r in rows[:12]:
    print("  ", r)

d.close()
d.switch_to.window(main)
print("\n팝업 닫고 복귀. written:", path)
