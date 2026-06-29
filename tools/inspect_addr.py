# -*- coding: utf-8 -*-
"""주소입력 버튼의 동작 방식 파악: onclick 핸들러, zipno readonly, 주소 관련 함수 탐지."""
import json
from selenium import webdriver

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

js = r"""
const out = {buttons: [], zipno: null, addr_fields: {}, funcs: [], scripts_hit: []};

// '주소입력' 텍스트를 가진 버튼/링크
for (const el of document.querySelectorAll('button, a, input[type=button], span')) {
  const t = (el.textContent||el.value||'').trim();
  if (t === '주소입력' || t.includes('주소입력')) {
    out.buttons.push({tag: el.tagName, text: t, onclick: el.getAttribute('onclick')||'',
                      outer: el.outerHTML.slice(0,200)});
  }
}

// zipno / addr 속성
for (const id of ['zipno','addr','addr_detail']) {
  const el = document.getElementById(id);
  if (el) out.addr_fields[id] = {readonly: el.readOnly, onclick: el.getAttribute('onclick')||'',
                                 onfocus: el.getAttribute('onfocus')||''};
}

// 주소 관련 전역 함수 이름 후보
for (const name of ['fnPost','fn_post','openPost','execDaumPostcode','daumPost',
                    'jusoCallBack','goPopup','fnAddr','fn_addr','searchAddr','fnSearchAddr',
                    'fnZip','zipPop','addrPop','fnPopup']) {
  if (typeof window[name] === 'function') out.funcs.push(name);
}

// 스크립트 본문에서 주소 팝업 단서
const keys = ['daum.Postcode','jusoCallBack','juso.go.kr','window.open','postcode','우편번호','roadAddr'];
for (const s of document.querySelectorAll('script')) {
  const txt = s.textContent||'';
  for (const k of keys) {
    if (txt.includes(k)) {
      const i = txt.indexOf(k);
      out.scripts_hit.push({key: k, near: txt.slice(Math.max(0,i-60), i+120).replace(/\s+/g,' ')});
    }
  }
}
return out;
"""
out = d.execute_script(js)
path = r"C:\Users\fivep\OneDrive\Desktop\CarMacro\addr.json"
with open(path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print("=== 주소입력 버튼 ===")
for b in out["buttons"]:
    print(" ", b["tag"], "| onclick:", b["onclick"])
    print("   outer:", b["outer"])
print("\n=== zipno/addr 필드 ===")
print(json.dumps(out["addr_fields"], ensure_ascii=False, indent=2))
print("\n=== 전역 주소함수 후보 ===", out["funcs"])
print("\n=== 스크립트 단서 ===")
for h in out["scripts_hit"][:15]:
    print(f"  [{h['key']}] {h['near']}")
print("\nwritten:", path)
