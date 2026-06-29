# -*- coding: utf-8 -*-
from selenium import webdriver
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
for fn in ["checkFinish"]:
    s=d.execute_script(f"try{{return {fn}.toString()}}catch(e){{return '(없음)'}}")
    print(f"function {fn}:"); print(s[:2000])
