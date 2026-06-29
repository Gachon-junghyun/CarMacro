# -*- coding: utf-8 -*-
"""정리(경고/잔여팝업) 후 주소 자동화 전체 흐름을 견고하게 테스트."""
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoAlertPresentException
import app

KEYWORD = "세종특별자치시 시청대로 78"

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)


def dismiss_alert():
    try:
        al = d.switch_to.alert
        print("  남은 경고 닫음:", al.text)
        al.accept()
    except NoAlertPresentException:
        pass
    except Exception:
        pass


# 1) 정리: 메인(sellerApplyform) 창만 남기고 나머지 닫기
main = None
for h in list(d.window_handles):
    try:
        d.switch_to.window(h)
        dismiss_alert()
        if d.current_url.endswith("sellerApplyform"):
            main = h
        elif "addrPopup" in d.current_url:
            d.close()
    except Exception as e:
        print("  정리 중:", e)
if main is None:
    main = d.window_handles[0]
d.switch_to.window(main)
dismiss_alert()
print("메인 복귀. zipno(전):", repr(app.get_value(d, "zipno")))

# 2) 주소 버튼 클릭 → 팝업
before = set(d.window_handles)
btn = app._find_addr_button(d)
print("주소버튼:", btn.get_attribute("onclick") if btn else None)
btn.click()
popup = None
for _ in range(20):
    time.sleep(0.3)
    new = set(d.window_handles) - before
    if new:
        popup = new.pop()
        break
print("팝업:", bool(popup))
d.switch_to.window(popup)

# 3) keyword 가 나타날 때까지 대기 후 JS 주입
kwval = ""
for _ in range(15):
    try:
        kwval = d.execute_script("return document.getElementById('keyword') ? 'Y':'N';")
        if kwval == "Y":
            break
    except Exception:
        pass
    time.sleep(0.2)
d.execute_script(
    "var k=document.getElementById('keyword'); k.value=arguments[0];"
    "k.dispatchEvent(new Event('input',{bubbles:true}));", KEYWORD)
print("검색어 확인:", repr(d.execute_script("return document.getElementById('keyword').value;")))

# 4) 검색
d.execute_script("$('#currentPage').val(1); searchUrlJuso();")
# 결과 폴링
cands = []
want = app._squash(KEYWORD)
for _ in range(15):
    time.sleep(0.3)
    dismiss_alert()
    try:
        cands = [a for a in d.find_elements(By.XPATH, "//a") if a.text and want in app._squash(a.text)]
    except Exception:
        cands = []
    if cands:
        break
print("매칭 앵커 수:", len(cands))
if cands:
    a = cands[0]
    print("앵커 outer:", a.get_attribute("outerHTML")[:200])
    try:
        print("setMaping 본문:", d.execute_script("return setMaping.toString();")[:600])
    except Exception as e:
        print("  setMaping 본문 못읽음:", e)
    # 클릭 대신 JS로 직접 호출
    d.execute_script("setMaping('1');")
    time.sleep(0.8)
    dismiss_alert()
    # 선택 후 팝업 내부 상태 진단
    diag = d.execute_script("""
      function g(id){var e=document.getElementById(id);return e?e.value:'(none)';}
      return {road:g('roadFullAddr'), zip:g('zipNo'), rtRoad:g('rtRoadAddr'), rtZip:g('rtZipNo'),
              inputYn:g('inputYn'),
              opener: (typeof opener!=='undefined' && opener) ? 'Y':'N',
              openerCb: (typeof opener!=='undefined' && opener && typeof opener.jusoCallBack==='function')?'Y':'N'};
    """)
    print("선택후 진단:", diag)
    try:
        print("setParent 본문:", d.execute_script("return setParent.toString();")[:500])
    except Exception as e:
        print("  setParent 본문 못읽음:", e)
    # 선택 후 부모로 넘기기: setParent() 호출 (있으면)
    try:
        has = d.execute_script("return typeof setParent === 'function' ? 'Y':'N';")
        print("setParent 존재:", has, "| 현재 URL:", d.current_url)
        if has == "Y":
            d.execute_script("setParent();")
            time.sleep(1.0)
            dismiss_alert()
    except Exception as e:
        print("  setParent 호출 중:", e)

# 5) 메인 복귀 결과
try:
    if popup in d.window_handles:
        d.switch_to.window(popup); d.close()
except Exception:
    pass
d.switch_to.window(main)
dismiss_alert()
print("RESULT zipno(후):", repr(app.get_value(d, "zipno")), "addr:", repr(app.get_value(d, "addr")))
