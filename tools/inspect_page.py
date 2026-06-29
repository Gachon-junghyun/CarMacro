# -*- coding: utf-8 -*-
"""현재 떠 있는 크롬(포트 9222)에 붙어서 폼의 input/select/radio id·name을 덤프."""
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

print("URL:", d.current_url)
print("TITLE:", d.title)

# iframe 안에 폼이 있을 수 있으니 모든 프레임을 훑는다
def dump(ctx_name):
    print(f"\n===== context: {ctx_name} =====")
    els = d.find_elements(By.XPATH, "//input | //select | //textarea")
    for el in els:
        tag = el.tag_name
        t = el.get_attribute("type") or ""
        i = el.get_attribute("id") or ""
        n = el.get_attribute("name") or ""
        v = el.get_attribute("value") or ""
        ph = el.get_attribute("placeholder") or ""
        if tag == "select":
            opts_txt = []
            for o in el.find_elements(By.TAG_NAME, "option"):
                opts_txt.append(f"{o.get_attribute('value')!r}:{o.text.strip()!r}")
            print(f"[select] id={i!r} name={n!r} options={opts_txt}")
        else:
            print(f"[{tag}:{t}] id={i!r} name={n!r} value={v!r} ph={ph!r}")

# 메인 문서
d.switch_to.default_content()
dump("main")

# 모든 iframe
frames = d.find_elements(By.TAG_NAME, "iframe")
for idx in range(len(frames)):
    try:
        d.switch_to.default_content()
        fr = d.find_elements(By.TAG_NAME, "iframe")[idx]
        fid = fr.get_attribute("id") or fr.get_attribute("name") or f"#{idx}"
        d.switch_to.frame(fr)
        dump(f"iframe {fid}")
    except Exception as e:
        print(f"iframe {idx} 접근 실패: {e}")

d.switch_to.default_content()
print("\n done.")
