# -*- coding: utf-8 -*-
"""첨부 팝업(popupAttachFile) 내부 구조 진단.
첨부 페이지가 떠 있는 상태에서 실행하면 실제 '첨부파일 등록' 버튼을 클릭해
'attachFile' 팝업 창을 띄우고 내부(파일 input, 업로드/저장 버튼)를 덤프한다.

  .venv\\Scripts\\python tools\\inspect_attach_popup.py [코드]
   코드 미지정 시 'A'. (예: A2, A3)

* execute_script 로 popupAttachFile 를 호출하면 크롬 팝업차단에 막힌다.
  → 사용자 제스처(실제 버튼 .click())로 열어야 팝업이 뜬다.
* 업로드/저장은 절대 누르지 않는다 — 구조만 본 뒤 팝업을 닫는다.
"""
import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By

JS = r"""
const out = {url: location.href, title: document.title,
             fileInputs: [], buttons: [], forms: [], funcs: [], iframes: []};
document.querySelectorAll('input[type=file]').forEach((el,i)=>{
  const cs = getComputedStyle(el);
  out.fileInputs.push({i, id: el.id||'', name: el.name||'', accept: el.accept||'',
    multiple: !!el.multiple, disabled: !!el.disabled,
    hidden: (el.offsetParent===null)||cs.display==='none'||cs.visibility==='hidden',
    onchange:(el.getAttribute('onchange')||'').slice(0,140)});
});
document.querySelectorAll('button,a,input[type=button],input[type=submit]').forEach(b=>{
  const t=(b.textContent||b.value||'').trim();
  if(t && t.length<24){
    out.buttons.push({t:t.slice(0,22), tag:b.tagName, id:b.id||'',
      oc:(b.getAttribute('onclick')||'').slice(0,160), cls:(b.className||'').slice(0,60)});
  }
});
document.querySelectorAll('form').forEach(f=>{
  out.forms.push({name:f.name||'', id:f.id||'', action:(f.action||'').slice(0,140),
    enctype:f.enctype||'', method:f.method||''});
});
const cand=['goUpload','fnUpload','upload','fn_save','goSave','save','fnAdd','add',
  'fnFileUpload','fileUpload','doUpload','fnAttach','attach','submitFile','fileSave',
  'fnFileSave','goFileUpload','fn_upload','fnConfirm','confirm2'];
cand.forEach(n=>{ if(typeof window[n]==='function') out.funcs.push(n); });
document.querySelectorAll('iframe').forEach(f=>out.iframes.push({id:f.id||'',name:f.name||'',src:(f.src||'').slice(0,140)}));
return out;
"""


def dump(d, tag):
    r = d.execute_script(JS)
    print(f"\n========== [{tag}] {r['url']} ==========")
    print("title:", r["title"])
    print(f"\n[파일 input {len(r['fileInputs'])}개]")
    for f in r["fileInputs"]:
        print(f"  #{f['i']} id={f['id']!r} name={f['name']!r} hidden={f['hidden']} "
              f"disabled={f['disabled']} multiple={f['multiple']} accept={f['accept']!r}")
        if f["onchange"]:
            print(f"       onchange={f['onchange']!r}")
    print(f"\n[버튼 {len(r['buttons'])}개]")
    for b in r["buttons"]:
        print(f"  <{b['tag']}> '{b['t']}' id={b['id']!r} cls={b['cls']!r}")
        if b["oc"]:
            print(f"       onclick={b['oc']!r}")
    print(f"\n[form] {r['forms']}")
    print(f"\n[전역 함수 후보] {r['funcs']}")
    print(f"\n[iframe] {r['iframes']}")
    return r


def find_attach_button(d, code):
    """popupAttachFile('<code>') 를 호출하는 버튼을 정확히 찾는다."""
    # onclick 에 'popupAttachFile('A');' 처럼 세미콜론까지 포함시켜 A vs A2 구분
    needle = f"popupAttachFile('{code}');"
    xp = f"//button[contains(@onclick, \"{needle}\")]"
    els = d.find_elements(By.XPATH, xp)
    return els[0] if els else None


def main():
    code = sys.argv[1] if len(sys.argv) > 1 else "A"
    o = webdriver.ChromeOptions()
    o.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    o.set_capability("unhandledPromptBehavior", "ignore")
    d = webdriver.Chrome(options=o)
    base = os.path.dirname(os.path.abspath(__file__))

    # 첨부 버튼이 있는 창/탭으로 전환
    target_handle = None
    for h in d.window_handles:
        d.switch_to.window(h)
        if find_attach_button(d, code):
            target_handle = h
            break
    if not target_handle:
        print(f"popupAttachFile('{code}') 버튼을 가진 창을 못 찾았습니다. 첨부 페이지가 떠 있는지 확인하세요.")
        return

    main_handle = target_handle
    before = set(d.window_handles)
    print(f"메인 창: {d.current_url}")

    btn = find_attach_button(d, code)
    print(f"'{code}' 첨부 버튼 클릭… (onclick={btn.get_attribute('onclick')!r})")
    try:
        d.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        btn.click()  # 실제 클릭 = 사용자 제스처 → 팝업차단 우회
    except Exception as e:
        print("버튼 클릭 오류:", e)

    # 새 팝업 창 대기
    popup = None
    for _ in range(20):
        time.sleep(0.3)
        new = set(d.window_handles) - before
        if new:
            popup = new.pop()
            break

    if not popup:
        print("새 팝업 창이 안 떴습니다. (인라인 레이어일 수 있음 → 메인 창에서 파일 input 재검색)")
        dump(d, "메인창(팝업없음)")
        with open(os.path.join(base, "_attach_popup_dump.html"), "w", encoding="utf-8") as f:
            f.write(d.execute_script("return document.documentElement.outerHTML;"))
        print("\n>> HTML 저장: tools/_attach_popup_dump.html")
        return

    d.switch_to.window(popup)
    time.sleep(0.6)
    dump(d, f"팝업({code})")
    # iframe 내부도 확인
    try:
        frames = d.find_elements(By.TAG_NAME, "iframe")
        for fi, fr in enumerate(frames):
            try:
                d.switch_to.frame(fr)
                dump(d, f"팝업 iframe{fi+1}")
                d.switch_to.window(popup)
            except Exception:
                d.switch_to.window(popup)
    except Exception:
        pass

    with open(os.path.join(base, "_attach_popup_dump.html"), "w", encoding="utf-8") as f:
        f.write(d.execute_script("return document.documentElement.outerHTML;"))
    print("\n>> 팝업 HTML 저장: tools/_attach_popup_dump.html")

    # 팝업 닫고 메인 복귀(업로드 안 함)
    try:
        d.close()
        d.switch_to.window(main_handle)
        print(">> 팝업 닫고 메인 창 복귀 완료(업로드는 하지 않음).")
    except Exception:
        pass


if __name__ == "__main__":
    main()
