# -*- coding: utf-8 -*-
import json
from selenium import webdriver
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
js=r"""
const out={saveBtns:[], randomUse:[], funcs:[]};
// 저장/임시저장/신청 버튼 + onclick
document.querySelectorAll("button,a,input[type=button],input[type=submit]").forEach(b=>{
  const t=(b.textContent||b.value||'').trim();
  if(/임시저장|저장|신청서?\s*저장|제출|등록/.test(t) && t.length<25)
    out.saveBtns.push({t:t.slice(0,20), oc:(b.getAttribute('onclick')||'').slice(0,140)});
});
// 스크립트에서 randomVal / random / 보안 / 역순 단서
document.querySelectorAll('script').forEach(s=>{
  const x=s.textContent||'';
  ['randomVal','역순','reverse','보안','random','captcha'].forEach(k=>{
    let i=x.indexOf(k);
    while(i>=0 && out.randomUse.length<25){
      out.randomUse.push({k, near:x.slice(Math.max(0,i-50),i+90).replace(/\s+/g,' ')});
      i=x.indexOf(k,i+1);
    }
  });
});
// 전역 함수 이름 후보
for(const n of ['fnSave','save','fn_save','goSave','fnApply','submitForm','fnRandom','randomCheck','fnConfirm','chkRandom','fnTempSave','tempSave'])
  if(typeof window[n]==='function') out.funcs.push(n);
return out;
"""
r=d.execute_script(js)
print("URL:", d.current_url)
print("\n[저장/제출 버튼]"); 
for b in r["saveBtns"]: print("  ", b)
print("\n[random/보안/역순 스크립트 단서]")
for u in r["randomUse"]: print("  -", u["k"], "::", u["near"])
print("\n[전역 함수]", r["funcs"])
