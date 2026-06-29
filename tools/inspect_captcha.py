# -*- coding: utf-8 -*-
"""현재 폼에서 보안문자/암호/역순(캡차) 관련 요소를 탐색해 어디에 입력하는지 찾는다."""
import json
from selenium import webdriver

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

# 메인 + 모든 iframe 훑기
def scan(label):
    js = r"""
    const KW = ['보안','암호','역순','거꾸로','자동입력','방지','캡차','captcha','인증','random','secur','확인문자'];
    const out = {texts:[], imgs:[], inputs:[], buttons:[]};
    // 키워드가 들어간 텍스트 노드(짧게)
    const wal="";
    document.querySelectorAll('body *').forEach(el=>{
      if (el.children.length) return;
      const t=(el.textContent||'').trim();
      if (t && t.length<60 && KW.some(k=>t.toLowerCase().includes(k.toLowerCase())))
        out.texts.push(t);
    });
    // 이미지(캡차 후보)
    document.querySelectorAll('img').forEach(im=>{
      const s=(im.src||'')+' '+(im.alt||'')+' '+(im.id||'');
      if (KW.some(k=>s.toLowerCase().includes(k.toLowerCase())) || /captcha|secur|random|보안|catpcha/i.test(im.src||''))
        out.imgs.push({src:(im.src||'').slice(0,140), alt:im.alt||'', id:im.id||''});
    });
    // 입력칸 전부(id,name,라벨,placeholder)
    document.querySelectorAll('input,textarea').forEach(inp=>{
      const id=inp.id||'', nm=inp.name||'', ph=inp.placeholder||'', ty=inp.type||'';
      let lab='';
      if (id){const l=document.querySelector('label[for="'+id+'"]'); if(l)lab=l.textContent.trim();}
      const blob=(id+' '+nm+' '+ph+' '+lab).toLowerCase();
      if (KW.some(k=>blob.includes(k.toLowerCase())) || ty==='password')
        out.inputs.push({id,nm,ph,ty,lab,hidden:(inp.type==='hidden'||inp.offsetParent===null)});
    });
    // 버튼(저장/제출/확인 등 — 캡차 트리거 후보)
    document.querySelectorAll('button,a,input[type=button],input[type=submit]').forEach(b=>{
      const t=(b.textContent||b.value||'').trim();
      if (/저장|제출|신청|확인|임시/.test(t))
        out.buttons.push({text:t.slice(0,20), onclick:(b.getAttribute('onclick')||'').slice(0,100)});
    });
    return out;
    """
    return d.execute_script(js)

print("URL:", d.current_url)
main = scan("main")
print("\n[보안/암호 관련 텍스트]", json.dumps(main["texts"], ensure_ascii=False))
print("\n[캡차 후보 이미지]", json.dumps(main["imgs"], ensure_ascii=False))
print("\n[관련 입력칸 / password]", json.dumps(main["inputs"], ensure_ascii=False))
print("\n[저장/제출 버튼]", json.dumps(main["buttons"], ensure_ascii=False)[:800])

# iframe 안도
frames = d.find_elements("tag name", "iframe")
for i in range(len(frames)):
    try:
        d.switch_to.default_content()
        fr = d.find_elements("tag name", "iframe")[i]
        fid = fr.get_attribute("id") or fr.get_attribute("src") or f"#{i}"
        d.switch_to.frame(fr)
        info = scan(f"iframe {fid}")
        if info["texts"] or info["imgs"] or info["inputs"]:
            print(f"\n=== iframe {fid} ===")
            print("  texts:", json.dumps(info["texts"], ensure_ascii=False))
            print("  imgs:", json.dumps(info["imgs"], ensure_ascii=False))
            print("  inputs:", json.dumps(info["inputs"], ensure_ascii=False))
    except Exception as e:
        print("iframe", i, "err", e)
d.switch_to.default_content()
