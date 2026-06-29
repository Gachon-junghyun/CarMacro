# -*- coding: utf-8 -*-
"""JS textContent 로 select 옵션·라디오 라벨을 안정적으로 덤프."""
import json
from selenium import webdriver

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
d = webdriver.Chrome(options=opts)

js = r"""
const out = {url: location.href, selects: {}, radios: {}};
for (const s of document.querySelectorAll('select')) {
  const key = s.id || s.name; if (!key) continue;
  out.selects[key] = [...s.options].map(o => [o.value, (o.textContent||'').trim()]);
}
for (const r of document.querySelectorAll('input[type=radio]')) {
  const name = r.name; if (!name) continue;
  let label = '';
  if (r.id) { const l = document.querySelector(`label[for="${r.id}"]`); if (l) label = (l.textContent||'').trim(); }
  if (!label && r.parentElement) label = (r.parentElement.textContent||'').trim();
  (out.radios[name] = out.radios[name] || []).push([r.id, r.value, label]);
}
return out;
"""
out = d.execute_script(js)
path = r"C:\Users\fivep\OneDrive\Desktop\CarMacro\labels2.json"
with open(path, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print("written:", path)
print("req_kind:", out["selects"].get("req_kind"))
print("model_cd count:", len(out["selects"].get("model_cd", [])))
print("priority_type:", out["radios"].get("priority_type"))
