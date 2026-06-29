# -*- coding: utf-8 -*-
"""차종(model_cd) 옵션 메뉴를 따고, 각 텍스트칸의 라벨 추정값을 덤프.
계약번호/휴대폰 등 매핑할 필드 id 를 찾는 용도."""
import json
from selenium import webdriver

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

# 1) model_cd 옵션 전체
js_model = r"""
const el = document.getElementById('model_cd');
if (!el) return [];
return [...el.options].map(o => [o.value, (o.textContent||'').trim()]);
"""
models = d.execute_script(js_model)

# 2) 텍스트칸 라벨 추정: input 가장 가까운 th/label/앞 텍스트
js_inputs = r"""
function labelFor(inp){
  if (inp.id){ const l=document.querySelector(`label[for="${inp.id}"]`); if(l&&l.textContent.trim()) return l.textContent.trim(); }
  // 같은 행(tr)의 th
  let tr = inp.closest('tr');
  if (tr){ const th = tr.querySelector('th'); if(th&&th.textContent.trim()) return th.textContent.trim().replace(/\s+/g,' '); }
  // 부모 td 의 이전 형제
  let td = inp.closest('td');
  if (td && td.previousElementSibling){ const t=td.previousElementSibling.textContent.trim(); if(t) return t.replace(/\s+/g,' '); }
  return '';
}
const out=[];
for (const inp of document.querySelectorAll('input[type=text], input:not([type]), textarea')){
  out.push([inp.id||'', labelFor(inp), inp.placeholder||'']);
}
return out;
"""
inputs = d.execute_script(js_inputs)

out = {"url": d.current_url, "model_cd": models, "inputs": inputs}
path = r"C:\Users\fivep\OneDrive\Desktop\CarMacro\map.json"
with open(path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("model_cd 옵션 수:", len(models))
for v, t in models:
    print(f"  {v!r:28} {t}")
# 행 전체 텍스트 기반으로 계약번호 칸 재탐색
js_rowtext = r"""
const out=[];
for (const inp of document.querySelectorAll('input[type=text], input:not([type]), textarea')){
  let tr = inp.closest('tr');
  let rowtext = tr ? (tr.textContent||'').replace(/\s+/g,' ').trim() : '';
  out.push([inp.id||'', rowtext]);
}
return out;
"""
rowtexts = d.execute_script(js_rowtext)
print("\n계약/관리(계약)번호 칸 (행 텍스트 기준):")
for iid, rt in rowtexts:
    if "계약" in rt or "관리(계약)" in rt or "제조수입사 관리" in rt:
        print(f"  id={iid!r:22} row={rt[:60]!r}")
print("\n휴대폰/연락처 칸:")
for iid, lbl, ph in inputs:
    if "휴대폰" in lbl or "전화" in lbl or "mobile" in iid or "phone" in iid:
        print(f"  id={iid!r:22} label={lbl!r} ph={ph!r}")
print("\nwritten:", path)
