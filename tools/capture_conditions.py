# -*- coding: utf-8 -*-
"""9222 크롬의 신청서 폼에서
  1) 현재 DOM 전체(outerHTML)를 저장
  2) 라디오 그룹(신청 조건 Y/N 등) 구조를 파악
  3) 각 Y/N 라디오를 눌러가며 폼이 어떻게 바뀌는지(숨김칸 노출/활성화) 캡처
끝나면 원래 선택값으로 복원한다. 저장/제출은 절대 안 함.

STEP=explore : DOM 저장 + 라디오 구조만 덤프(폼 안 건드림)
STEP=toggle  : 각 Y/N 눌러가며 변화 캡처(원상복구)
"""
import os
import sys
import re
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By

OUTDIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                      "dom_captures")
os.makedirs(OUTDIR, exist_ok=True)

opts = webdriver.ChromeOptions()
opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
opts.set_capability("unhandledPromptBehavior", "ignore")
d = webdriver.Chrome(options=opts)

# 신청서 폼 창으로 전환
for h in d.window_handles:
    d.switch_to.window(h)
    if d.current_url.endswith("sellerApplyform"):
        break
url = d.current_url
print("URL:", url)


# ── 현재 화면에서 "보이는 폼 컨트롤" 스냅샷을 찍는 JS ─────────────
# 각 input/select/textarea 의 id/name/type/보임여부/disabled/현재값,
# 그리고 눈에 보이는 라벨성 텍스트(th/label)를 함께 수집.
SNAP_JS = r"""
function vis(el){
  const r = el.getBoundingClientRect();
  const s = getComputedStyle(el);
  return !(s.display==='none' || s.visibility==='hidden' || (r.width===0 && r.height===0));
}
const out = [];
document.querySelectorAll('input,select,textarea').forEach(el=>{
  const type = (el.type||el.tagName).toLowerCase();
  if(type==='hidden') return;
  let value = '';
  if(el.tagName==='SELECT'){ const o=el.options[el.selectedIndex]; value=(o?o.text:''); }
  else if(type==='radio'||type==='checkbox'){ value = el.checked?'[x]':'[ ]'; }
  else { value = el.value||''; }
  out.push({
    id: el.id||'', name: el.name||'', type: type,
    value: value, vis: vis(el), disabled: !!el.disabled
  });
});
return JSON.stringify(out);
"""


def snap():
    return json.loads(d.execute_script(SNAP_JS))


def snap_key(rows):
    """비교용: 보이고 활성화된 컨트롤의 (name/id + type) 집합."""
    s = set()
    for r in rows:
        if r["vis"] and not r["disabled"]:
            s.add((r["name"] or r["id"], r["type"]))
    return s


# ── 라디오 그룹 구조 덤프 ──────────────────────────────────────
RADIO_JS = r"""
function labelFor(el){
  // for=id 라벨 우선, 없으면 부모 라벨/형제 텍스트
  if(el.id){
    const l=document.querySelector("label[for='"+el.id+"']");
    if(l) return (l.textContent||'').trim();
  }
  const p=el.closest('label'); if(p) return (p.textContent||'').trim();
  const sib=el.nextSibling; if(sib&&sib.textContent) return sib.textContent.trim();
  return '';
}
function rowLabel(el){
  // 이 컨트롤이 속한 행(tr)의 th/첫 셀 텍스트 = 항목명
  const tr=el.closest('tr');
  if(tr){ const th=tr.querySelector('th'); if(th) return (th.textContent||'').replace(/\s+/g,' ').trim(); }
  return '';
}
const groups={};
document.querySelectorAll("input[type=radio]").forEach(el=>{
  const n=el.name||el.id; if(!n) return;
  (groups[n]=groups[n]||{name:n, row:rowLabel(el), opts:[]}).opts.push({
    value: el.value||'', id: el.id||'', label: labelFor(el),
    checked: el.checked, disabled: !!el.disabled
  });
});
return JSON.stringify(Object.values(groups));
"""
radio_groups = json.loads(d.execute_script(RADIO_JS))

step = os.environ.get("STEP", "explore")

# 1) DOM 전체 저장 (항상)
html = d.execute_script("return document.documentElement.outerHTML;")
dom_path = os.path.join(OUTDIR, "form_dom.html")
with open(dom_path, "w", encoding="utf-8") as f:
    f.write("<!-- URL: %s -->\n" % url + html)
print("saved DOM:", dom_path, "(%d KB)" % (len(html) // 1024))

# 라디오 구조 저장
with open(os.path.join(OUTDIR, "radio_groups.json"), "w", encoding="utf-8") as f:
    json.dump({"url": url, "groups": radio_groups}, f, ensure_ascii=False, indent=2)

print("\n[ 라디오 그룹 %d개 ]" % len(radio_groups))
for g in radio_groups:
    checked = next((o["value"] for o in g["opts"] if o["checked"]), "")
    optstr = " / ".join("%s=%s%s" % (o["value"], o["label"] or "?",
                                     "(disabled)" if o["disabled"] else "")
                        for o in g["opts"])
    print("  · %-18s [%s]  현재=%s  | %s"
          % (g["name"], g["row"], checked or "(없음)", optstr))

if step == "explore":
    print("\n(explore 모드 — 폼은 건드리지 않음. 변화 캡처는 STEP=toggle 로 실행)")
    sys.exit(0)


# ── 2) Y/N 라디오 토글하며 변화 캡처 ──────────────────────────
def set_radio(name, value):
    els = d.find_elements(By.CSS_SELECTOR,
                          "input[type=radio][name='%s']" % name)
    for el in els:
        if el.get_attribute("value") == value:
            if not el.is_enabled():
                return "disabled"
            d.execute_script("arguments[0].click();", el)
            return "ok"
    return "missing"


baseline = snap()
base_key = snap_key(baseline)
report = {"url": url, "changes": {}}

# 신청 조건성 Y/N 그룹만 대상(값이 Y/N 위주)
targets = [g for g in radio_groups
           if {o["value"] for o in g["opts"]} & {"Y", "N"}]
print("\n[ 토글 대상 %d개 ]" % len(targets),
      ", ".join(g["name"] for g in targets))

# 원상복구용 원래값
original = {g["name"]: next((o["value"] for o in g["opts"] if o["checked"]), None)
           for g in targets}

for g in targets:
    name = g["name"]
    g_report = {"row": g["row"], "states": {}}
    for val in ("Y", "N"):
        if val not in {o["value"] for o in g["opts"]}:
            continue
        st = set_radio(name, val)
        if st != "ok":
            g_report["states"][val] = {"click": st}
            continue
        time.sleep(0.5)  # 조건부 섹션 노출 대기
        after = snap()
        after_key = snap_key(after)
        appeared = sorted(after_key - base_key)
        disappeared = sorted(base_key - after_key)
        g_report["states"][val] = {
            "click": "ok",
            "appeared": ["%s(%s)" % (n, t) for n, t in appeared],
            "disappeared": ["%s(%s)" % (n, t) for n, t in disappeared],
        }
    report["changes"][name] = g_report
    # 이 그룹 복원 후 다음 그룹 (baseline 기준 유지)
    if original.get(name):
        set_radio(name, original[name])
        time.sleep(0.3)

with open(os.path.join(OUTDIR, "condition_changes.json"), "w", encoding="utf-8") as f:
    json.dump(report, f, ensure_ascii=False, indent=2)

print("\n[ Y/N 토글 → 폼 변화 ]")
for name, gr in report["changes"].items():
    print("\n● %s  (%s)" % (name, gr["row"]))
    for val, info in gr["states"].items():
        if info["click"] != "ok":
            print("   %s: 클릭 %s" % (val, info["click"]))
            continue
        app_ = info["appeared"]
        dis_ = info["disappeared"]
        print("   %s → 나타남: %s | 사라짐: %s"
              % (val, ", ".join(app_) or "-", ", ".join(dis_) or "-"))

print("\n원상복구 완료. saved:", os.path.join(OUTDIR, "condition_changes.json"))
