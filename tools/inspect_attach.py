# -*- coding: utf-8 -*-
"""임시저장 후 뜨는 '첨부파일' 페이지 구조 진단.
크롬(9222)에서 그 페이지를 띄워둔 채 실행하세요.

  .venv\\Scripts\\python tools\\inspect_attach.py

- 파일 input(숨김 포함), 첨부/등록 버튼, 각 첨부행 라벨, 전역함수, iframe 을 출력
- 전체 HTML 을 tools/_attach_dump.html 로 저장(필요시 사람이/도구가 직접 확인)
"""
import os
import sys
import json
from selenium import webdriver
from selenium.webdriver.common.by import By

JS = r"""
function rowText(el){
  let n = el;
  for(let h=0; h<8 && n; h++){
    const tr = n.closest ? n.closest('tr') : null;
    if(tr){ return (tr.innerText||'').replace(/\s+/g,' ').trim().slice(0,160); }
    n = n.parentElement;
  }
  // tr 이 없으면 부모 컨테이너 텍스트
  let p = el.parentElement;
  for(let h=0; h<4 && p; h++){
    const t = (p.innerText||'').replace(/\s+/g,' ').trim();
    if(t) return t.slice(0,160);
    p = p.parentElement;
  }
  return '';
}
const out = {url: location.href, title: document.title,
             fileInputs: [], buttons: [], funcs: [], iframes: [], labels: []};

document.querySelectorAll('input[type=file]').forEach((el,i)=>{
  const cs = getComputedStyle(el);
  out.fileInputs.push({
    i, id: el.id||'', name: el.name||'', accept: el.accept||'',
    disabled: !!el.disabled,
    hidden: (el.offsetParent===null) || cs.display==='none' || cs.visibility==='hidden',
    onchange: (el.getAttribute('onchange')||'').slice(0,120),
    near: rowText(el)
  });
});

document.querySelectorAll('button,a,input[type=button],input[type=submit],span,label,img').forEach(b=>{
  const t = (b.textContent||b.value||b.alt||'').trim();
  if(/첨부|등록|파일|찾아보기|업로드|선택|삭제|변경/.test(t) && t.length<30){
    out.buttons.push({t:t.slice(0,28), tag:b.tagName, id:b.id||'',
      oc:(b.getAttribute('onclick')||'').slice(0,180), cls:(b.className||'').slice(0,70),
      near: rowText(b)});
  }
});

// 첨부 항목 라벨 후보(행 전체 텍스트)
const seen = new Set();
document.querySelectorAll('th,td,label,div,li').forEach(e=>{
  const t=(e.textContent||'').replace(/\s+/g,' ').trim();
  if(/증빙|계약서|신청서|구비서류|자격|첨부|우선순위|서류/.test(t) && t.length<80){
    if(!seen.has(t)){ seen.add(t); out.labels.push(t); }
  }
});

const cand = ['fnFileUpload','fileUpload','goFile','addFile','fn_addFile','fnAttach',
  'attachFile','fnAddFile','fnFile','goUpload','upload','fnUpload','fileAdd','fnFileAdd',
  'openFilePop','filePop','fnFilePop','goAttach','fn_file_add'];
cand.forEach(n=>{ if(typeof window[n]==='function') out.funcs.push(n); });

document.querySelectorAll('iframe').forEach(f=>{
  out.iframes.push({id:f.id||'', name:f.name||'', src:(f.src||'').slice(0,140)});
});
return out;
"""


def dump_window(d, tag=""):
    r = d.execute_script(JS)
    print(f"\n========== [{tag}] {r['url']} ==========")
    print("title:", r["title"])
    print(f"\n[파일 input  {len(r['fileInputs'])}개]")
    for f in r["fileInputs"]:
        print(f"  #{f['i']} id={f['id']!r} name={f['name']!r} hidden={f['hidden']} "
              f"disabled={f['disabled']} accept={f['accept']!r}")
        print(f"       onchange={f['onchange']!r}")
        print(f"       행텍스트: {f['near']}")
    print(f"\n[첨부/등록 관련 버튼  {len(r['buttons'])}개]")
    for b in r["buttons"]:
        print(f"  <{b['tag']}> '{b['t']}' id={b['id']!r} cls={b['cls']!r}")
        if b["oc"]:
            print(f"       onclick={b['oc']!r}")
        if b["near"]:
            print(f"       행텍스트: {b['near']}")
    print(f"\n[첨부 항목 라벨 후보]")
    for t in r["labels"]:
        print("   -", t)
    print(f"\n[전역 함수 후보] {r['funcs']}")
    print(f"\n[iframe] {r['iframes']}")
    return r


def main():
    o = webdriver.ChromeOptions()
    o.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    try:
        d = webdriver.Chrome(options=o)
    except Exception as e:
        print("크롬 연결 실패(포트 9222 확인):", e)
        sys.exit(1)

    base = os.path.dirname(os.path.abspath(__file__))

    # 현재 창 + 다른 탭들 중 파일 input 이 있는 창 모두 진단
    handles = d.window_handles
    best_html = None
    for idx, h in enumerate(handles):
        try:
            d.switch_to.window(h)
            r = dump_window(d, tag=f"창{idx+1}/{len(handles)}")
            if r["fileInputs"]:
                best_html = d.execute_script("return document.documentElement.outerHTML;")
        except Exception as e:
            print(f"  (창{idx+1} 검사 실패: {e})")

    # iframe 내부도 한 번 들여다보기(현재 창 기준)
    try:
        frames = d.find_elements(By.TAG_NAME, "iframe")
        for fi, fr in enumerate(frames):
            try:
                d.switch_to.default_content()
                d.switch_to.frame(fr)
                r = dump_window(d, tag=f"iframe{fi+1}")
                if r["fileInputs"] and not best_html:
                    best_html = d.execute_script("return document.documentElement.outerHTML;")
            except Exception:
                pass
        d.switch_to.default_content()
    except Exception:
        pass

    if best_html:
        out_path = os.path.join(base, "_attach_dump.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(best_html)
        print(f"\n>> 파일 input 이 있는 페이지 HTML 저장: {out_path}")
    else:
        # 그래도 현재 창 HTML 저장
        out_path = os.path.join(base, "_attach_dump.html")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(d.execute_script("return document.documentElement.outerHTML;"))
        print(f"\n>> (파일 input 미발견) 현재 창 HTML 저장: {out_path}")


if __name__ == "__main__":
    main()
