# -*- coding: utf-8 -*-
from selenium import webdriver
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
for fn in ["setReverseStr","goSave_1"]:
    src=d.execute_script(f"try{{return {fn}.toString()}}catch(e){{return '(없음)'}}")
    print("="*70); print(f"function {fn}:"); print(src[:1800])
# setReverseStr 가 참조하는 요소들의 현재 상태도
print("="*70)
info=d.execute_script(r"""
function g(id){var e=document.getElementById(id);return e?{val:e.value,tag:e.tagName,type:e.type,vis:e.offsetParent!==null}:null;}
return {randomVal:g('randomVal'), clickChk:g('clickChk'), clickChkMsg:g('clickChkMsg')};
""")
print("관련 필드 상태:", info)
