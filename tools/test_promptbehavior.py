import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from selenium import webdriver
o=webdriver.ChromeOptions()
o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
o.set_capability("unhandledPromptBehavior","ignore")
d=webdriver.Chrome(options=o)
print("unhandledPromptBehavior:", d.capabilities.get("unhandledPromptBehavior"))
# get_value 가 알림 떠 있어도 죽지 않는지(여기선 일반 동작만 확인)
for h in d.window_handles:
    d.switch_to.window(h)
    if d.current_url.endswith("sellerApplyform"): break
print("폼 읽기 OK, req_nm:", repr(d.execute_script("var e=document.getElementById('req_nm');return e?e.value:'(none)'")))
