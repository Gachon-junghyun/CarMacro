# -*- coding: utf-8 -*-
import time
from selenium import webdriver
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
# 수소 폼 창 기준 (opener 컨텍스트는 inspection엔 불필요, 쿠키만 있으면 됨)
for h in d.window_handles:
    d.switch_to.window(h)
    if d.current_url.endswith("sellerApplyform"): break
base = "https://ev.or.kr/ev_ps/h2/comm/popupSellerRandomChk"
d.switch_to.new_window('tab')
d.get(base)
time.sleep(1.0)
print("팝업 URL:", d.current_url, "| title:", d.title)
info=d.execute_script(r"""
const out={bodyText:(document.body.innerText||'').replace(/\s+/g,' ').slice(0,400), inputs:[], buttons:[], imgs:[]};
document.querySelectorAll('input,textarea').forEach(i=>out.inputs.push({id:i.id,name:i.name,type:i.type,val:i.value,ph:i.placeholder}));
document.querySelectorAll('button,a,input[type=button]').forEach(b=>out.buttons.push({t:(b.textContent||b.value||'').trim().slice(0,20),oc:(b.getAttribute('onclick')||'').slice(0,160)}));
document.querySelectorAll('img').forEach(im=>{ if(/random|secur|보안|captcha/i.test(im.src||'')) out.imgs.push((im.src||'').slice(0,140)); });
// 화면에 보이는 코드성 요소(span/div id에 rand/code/str 포함)
out.codeEls=[];
document.querySelectorAll('[id]').forEach(e=>{ if(/rand|code|str|secur|num/i.test(e.id) && e.children.length===0){ out.codeEls.push({id:e.id, text:(e.textContent||'').trim().slice(0,30), tag:e.tagName}); }});
return out;
""")
print("BODY:", info["bodyText"])
print("INPUTS:", info["inputs"])
print("BUTTONS:", info["buttons"])
print("IMGS:", info["imgs"])
print("CODE ELEMENTS:", info["codeEls"])
# 팝업 닫기
d.close()
d.switch_to.window(d.window_handles[0])
print("닫고 복귀 OK")
