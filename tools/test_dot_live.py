import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from selenium import webdriver
import app
o=webdriver.ChromeOptions(); o.add_experimental_option("debuggerAddress","127.0.0.1:9222")
d=webdriver.Chrome(options=o)
for h in d.window_handles:
    d.switch_to.window(h)
    if d.current_url.endswith("sellerApplyform"): break
app.fill_text(d,"phone",".")
app.fill_text(d,"email",".")
print("phone:", repr(app.get_value(d,"phone")), "| email:", repr(app.get_value(d,"email")))
