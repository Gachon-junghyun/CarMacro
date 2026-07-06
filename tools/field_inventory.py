# -*- coding: utf-8 -*-
"""9222 폼의 모든 입력 필드(id/type/label/보임)를 조사.
전환지원금(exchange_3year_yn)=Y, 신청유형 개인 선택 후 열리는 칸까지 포함.
읽기 위주 — 저장/제출 안 함. 조사 후 라디오는 원상복구하지 않음(다음 단계에서 채울 것)."""
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

INV_JS = r"""
function labelOf(el){
  const tr=el.closest('tr');
  let lab='';
  if(tr){ const th=tr.querySelector('th'); if(th) lab=(th.textContent||'').replace(/\s+/g,' ').trim(); }
  // 같은 셀(td) 안 라벨 p
  const td=el.closest('td');
  if(td){ const p=td.querySelector('p,label'); if(p){ const t=(p.textContent||'').replace(/\s+/g,' ').trim(); if(t&&t.length<20) lab=lab?lab+' | '+t:t; } }
  return lab;
}
function vis(el){const s=getComputedStyle(el);const r=el.getBoundingClientRect();
  return !(s.display==='none'||s.visibility==='hidden'||(r.width===0&&r.height===0));}
const out=[];
document.querySelectorAll('input,select,textarea').forEach(el=>{
  const type=(el.type||el.tagName).toLowerCase();
  if(type==='hidden') return;
  if(type==='radio'||type==='checkbox') return; // 라디오는 따로 봄
  let v='';
  if(el.tagName==='SELECT'){const o=el.options[el.selectedIndex];v=(o?o.text:'');}
  else v=el.value||'';
  out.push({id:el.id||'',name:el.name||'',type:type,label:labelOf(el),
            value:v,vis:vis(el),ro:!!el.readOnly,dis:!!el.disabled});
});
return JSON.stringify(out);
"""


def inv():
    return json.loads(d.execute_script(INV_JS))


def set_radio(name, value):
    for el in d.find_elements(By.CSS_SELECTOR, "input[type=radio][name='%s']" % name):
        if el.get_attribute("value") == value:
            d.execute_script("arguments[0].click();"
                             "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
            return True
    return False


def dump(tag):
    print("\n===== %s =====" % tag)
    for f in inv():
        flag = ("V" if f["vis"] else ".") + ("R" if f["ro"] else " ") + ("D" if f["dis"] else " ")
        print("  [%s] %-18s %-8s %-28s = %s"
              % (flag, f["id"] or f["name"], f["type"], f["label"][:28], f["value"][:20]))


dump("기본 상태")
# 개인 + 전환지원금 Y 켜고 재조사
print("\n>>> lease_kind=P, exchange_3year_yn=Y 적용")
set_radio("lease_kind", "P")
time.sleep(0.4)
set_radio("exchange_3year_yn", "Y")
time.sleep(0.8)
dump("개인 + 전환지원금 Y")

# 선택형 옵션 목록도 저장
sel = {}
for el in d.find_elements(By.TAG_NAME, "select"):
    sid = el.get_attribute("id") or el.get_attribute("name")
    if not sid:
        continue
    sel[sid] = [[o.get_attribute("value"), o.text.strip()]
                for o in el.find_elements(By.TAG_NAME, "option")]
json.dump({"inv": inv(), "selects": sel},
          open("dom_captures/field_inventory.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=2)
print("\nsaved dom_captures/field_inventory.json")
