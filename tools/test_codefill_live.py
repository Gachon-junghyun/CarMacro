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
# 앱과 동일 로직: goCompare 소스에서 코드 추출
src=d.execute_script("try{return goCompare.toString()}catch(e){return ''}")
m=re.search(r"=\s*'([0-9A-Za-z]{6,16})'\s*\.split", src)
code=m.group(1) if m else ""
rev=code[::-1]
d.execute_script("var e=document.getElementById('randeomChk');e.value=arguments[0];e.dispatchEvent(new Event('input',{bubbles:true}));", rev)
got=d.execute_script("return document.getElementById('randeomChk').value")
expect=re.search(r"=\s*'([0-9A-Za-z]+)'\.split", src).group(1)[::-1]
print("화면코드:", code, "| 넣은 역순:", got, "| 기대:", expect, "| 일치:", got==expect)
d.close(); d.switch_to.window(d.window_handles[0])
print("확인 클릭 안 함 → 저장 안 됨")
