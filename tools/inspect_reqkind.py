# -*- coding: utf-8 -*-
"""열린 폼들의 req_kind / priority_type 옵션·현재값·disabled 상태 비교."""
import json
from selenium import webdriver

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

js = r"""
function selInfo(id){
  var e=document.getElementById(id); if(!e) return null;
  return {value:e.value, disabled:e.disabled, onchange:(e.getAttribute('onchange')||'').slice(0,120),
          options:[...e.options].map(o=>[o.value,(o.textContent||'').trim()])};
}
function radioInfo(name){
  var rs=document.querySelectorAll('input[name="'+name+'"]'); if(!rs.length) return null;
  var out=[]; rs.forEach(r=>out.push([r.id, r.value, r.checked,
    (document.querySelector('label[for="'+r.id+'"]')||{}).textContent ?
     document.querySelector('label[for="'+r.id+'"]').textContent.trim():'']));
  return out;
}
return {url:location.href, req_kind:selInfo('req_kind'),
        priority_type:radioInfo('priority_type'),
        pri_business_yn: radioInfo('pri_business_yn'),
        profit_yn: radioInfo('profit_yn')};
"""

for h in d.window_handles:
    try:
        d.switch_to.window(h)
    except Exception:
        continue
    if "sellerApplyform" not in d.current_url:
        continue
    info = d.execute_script(js)
    kind = "수소" if "/h2/" in info["url"] else "전기"
    print("=" * 60)
    print(f"[{kind}] {info['url']}")
    rk = info["req_kind"]
    if rk:
        print("  req_kind 현재값:", repr(rk["value"]), "| disabled:", rk["disabled"])
        print("  req_kind onchange:", rk["onchange"])
        print("  req_kind 옵션:", rk["options"])
    else:
        print("  req_kind 없음")
    pt = info["priority_type"]
    print("  priority_type:", pt)
