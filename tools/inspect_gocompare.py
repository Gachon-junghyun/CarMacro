# -*- coding: utf-8 -*-
import time, re
from selenium import webdriver
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
for h in d.window_handles:
    d.switch_to.window(h)
    if d.current_url.endswith("sellerApplyform"): break
d.switch_to.new_window('tab')
d.get("https://ev.or.kr/ev_ps/h2/comm/popupSellerRandomChk")
time.sleep(1.0)
# 코드 요소 후보: 6~12자 영숫자 텍스트를 가진 말단 요소
data=d.execute_script(r"""
const out={cands:[], goCompare:''};
document.querySelectorAll('body *').forEach(e=>{
  if(e.children.length) return;
  const t=(e.textContent||'').trim();
  if(/^[A-Za-z0-9]{6,14}$/.test(t)) out.cands.push({tag:e.tagName,id:e.id,cls:e.className,text:t});
});
try{ out.goCompare=goCompare.toString(); }catch(e){ out.goCompare='(없음)'; }
return out;
""")
print("코드 후보 요소:")
for c in data["cands"]: print("  ", c)
print("\ngoCompare():")
print(data["goCompare"][:1600])
d.close(); d.switch_to.window(d.window_handles[0])
print("\n닫고 복귀 OK")
