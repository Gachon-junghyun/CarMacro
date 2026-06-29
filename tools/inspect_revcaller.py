# -*- coding: utf-8 -*-
from selenium import webdriver
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
# setReverseStr 를 호출하는 함수들 + 'cud'/보안팝업 여는 함수들 찾기
res=d.execute_script(r"""
const out={callers:[], openers:[]};
for(const k of Object.keys(window)){
  try{ if(typeof window[k]!=='function') continue;
    const s=window[k].toString();
    if(/setReverseStr\s*\(/.test(s) && k!=='setReverseStr') out.callers.push(k);
    if(/randomPop|보안|random_pop|reverseLayer|popupRandom|layer.*random|openRandom/i.test(s)) out.openers.push(k);
  }catch(e){}
}
return out;
""")
print("setReverseStr 호출 함수:", res["callers"])
print("보안팝업 여는 함수 후보:", res["openers"])
for fn in res["callers"][:4]:
    s=d.execute_script(f"return {fn}.toString()")
    print("="*70); print(fn, ":"); print(s[:1500])
# goSave_1 끝부분(보안팝업 호출 지점) 보기
s=d.execute_script("return goSave_1.toString()")
print("="*70); print("goSave_1 TAIL:"); print(s[-1500:])
