import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from selenium import webdriver
import app
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
ev=None
for h in d.window_handles:
    d.switch_to.window(h)
    u=d.current_url
    if u.endswith("sellerApplyform") and "/h2/" not in u: ev=h; break
if ev:
    print("EV 폼:", d.current_url)
    app.fill_text(d,"birth1","1978-06-11")
    print("EV birth1 값:", repr(app.get_value(d,"birth1")))
    print("EV req_kind 'B':", app.select_value(d,"req_kind","B"), "| 값:", app.get_value(d,"req_kind"))
else:
    print("열린 EV 폼 없음 (수소만 떠 있음) — EV는 위젯 동일하니 동일 동작 예상")
