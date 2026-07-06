# -*- coding: utf-8 -*-
"""특정 라디오 값을 클릭한 뒤, 관심 의존필드들의 '보임' 여부가 어떻게
바뀌는지 정밀 확인. 각 트리거 테스트 후 원상복구."""
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By

o = webdriver.ChromeOptions()
o.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
o.set_capability("unhandledPromptBehavior", "ignore")
d = webdriver.Chrome(options=o)
for h in d.window_handles:
    d.switch_to.window(h)
    if d.current_url.endswith("sellerApplyform"):
        break

# 관심 의존필드(대표 id 하나씩) — 이게 보이면 해당 섹션이 열린 것
WATCH = ["busi_no", "pri_busi_nm", "corp_no",
         "improve_fd_detail1", "improve_fd_detail2",
         "social_kind",
         "ls_user_kind1", "ls_user_nm", "ls_user_sex1", "ls_user_nm_chk1",
         "disposal_yn1", "old_car_no", "old_car_vin"]

VIS_JS = r"""
const ids = arguments[0]; const out = {};
ids.forEach(id=>{
  const el=document.getElementById(id);
  if(!el){ out[id]='X'; return; }
  const s=getComputedStyle(el); const r=el.getBoundingClientRect();
  out[id] = !(s.display==='none'||s.visibility==='hidden'||(r.width===0&&r.height===0)) ? 'V':'.';
});
return JSON.stringify(out);
"""


def vis():
    return json.loads(d.execute_script(VIS_JS, WATCH))


def cur_radio(name):
    for el in d.find_elements(By.CSS_SELECTOR, "input[type=radio][name='%s']" % name):
        if el.is_selected():
            return el.get_attribute("value")
    return None


def set_radio(name, value):
    for el in d.find_elements(By.CSS_SELECTOR, "input[type=radio][name='%s']" % name):
        if el.get_attribute("value") == value:
            d.execute_script("arguments[0].click();", el)
            d.execute_script(
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
            return True
    return False


def show(tag, before, after):
    diff = [k for k in WATCH if before.get(k) != after.get(k)]
    line = "  변화: " + (", ".join("%s %s→%s" % (k, before[k], after[k]) for k in diff)
                        if diff else "(없음)")
    print(tag)
    print(line)


TRIGGERS = [
    ("lease_kind", "P"), ("lease_kind", "G"),
    ("improve_fd_yn", "Y"),
    ("social_yn", "Y"),
    ("ls_user_yn", "Y"),
    ("truck_yn", "Y"),
    ("exchange_3year_yn", "Y"),
    ("taxi_yn", "Y"),
    ("bms_yn", "Y"),
    ("in_facility_yn", "Y"),
]

print("WATCH 기준선:", json.dumps(vis(), ensure_ascii=False))
for name, val in TRIGGERS:
    orig = cur_radio(name)
    before = vis()
    if not set_radio(name, val):
        print("\n● %s=%s : 옵션 없음" % (name, val))
        continue
    time.sleep(0.6)
    after = vis()
    show("\n● %s=%s (원래 %s)" % (name, val, orig), before, after)
    # 복구
    if orig is not None:
        set_radio(name, orig)
        time.sleep(0.3)
print("\n원상복구 완료.")
