# -*- coding: utf-8 -*-
from selenium import webdriver
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
for fn in ["goSave","goPriSave","goSavePay"]:
    src=d.execute_script(f"try{{return {fn}.toString()}}catch(e){{return '(없음)'}}")
    print("="*70)
    print(f"function {fn}:")
    print(src[:1400])
# randomVal 을 건드리는 다른 전역함수 탐색
hits=d.execute_script(r"""
const res=[];
for(const k of Object.keys(window)){
  try{ if(typeof window[k]==='function'){ const s=window[k].toString();
    if(/randomVal|역순|reverse|보안문자|보안코드/.test(s)) res.push(k); } }catch(e){}
}
return res;
""")
print("="*70)
print("randomVal/역순/보안 건드리는 함수들:", hits)
