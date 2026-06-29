# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from selenium import webdriver
import app
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
for h in d.window_handles:
    d.switch_to.window(h)
    if "/h2/" in d.current_url and d.current_url.endswith("sellerApplyform"): break
print("폼:", d.current_url)
st=app.select_value(d,"req_kind","B")
print("req_kind 'B' 결과:", st)
if st=="missing":
    st2=app.select_value(d,"req_kind","P")
    print("  대체 'P' 결과:", st2, "| 현재값:", app.get_value(d,"req_kind"))
app.fill_text(d,"birth1","1982-03-22")
print("birth1 입력 후 값:", repr(app.get_value(d,"birth1")))
