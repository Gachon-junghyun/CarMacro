# -*- coding: utf-8 -*-
"""9222 디버그 크롬에 읽기전용 CDP로 붙어 신청서 폼 상태를 진단.
- 폼 필드 현재값 / 빈 필수칸
- '임시저장' 버튼 후보(개수·텍스트·onclick·표시여부·disabled)
- 저장을 막는 흔한 원인 힌트
드라이버를 새로 띄우지 않으므로 실행 중인 앱과 충돌하지 않는다."""
import json
import urllib.request
import websocket  # websocket-client

FIELD_IDS = [
    "req_kind", "contract_day", "req_nm", "birth1", "birth2", "busi_no",
    "pri_busi_nm", "req_sex", "model_cd", "req_cnt", "delivery_sch_day",
    "zipno", "addr", "addr_detail", "phone", "mobile", "email",
    "social_yn", "social_kind", "priority_type",
    "contact_nm", "contact_mobile", "seller_mgrid",
]
REQUIRED = {"req_kind", "req_nm", "req_sex", "model_cd", "req_cnt", "mobile"}

JS = r"""
(function(){
  function val(id){
    var el = document.getElementById(id);
    if(!el){ return {exists:false}; }
    var v = null;
    if(el.tagName==='SELECT'){
      var o = el.options[el.selectedIndex];
      v = el.value; var txt = o? o.text : '';
      return {exists:true, tag:'select', value:v, text:txt};
    }
    if(el.type==='radio'){
      var g = document.getElementsByName(el.name);
      var checked=''; for(var i=0;i<g.length;i++){ if(g[i].checked) checked=g[i].value; }
      return {exists:true, tag:'radio', value:checked};
    }
    return {exists:true, tag:(el.tagName||'').toLowerCase(), value:(el.value||'')};
  }
  var ids = %s;
  var fields = {}; ids.forEach(function(id){ fields[id]=val(id); });

  // 임시저장 버튼 후보
  var nodes = document.querySelectorAll("button, a, input[type=button], input[type=submit]");
  var saves = [];
  nodes.forEach(function(el){
    var t = (el.innerText || el.value || '').trim();
    if(t.indexOf('임시저장')<0) return;
    var r = el.getBoundingClientRect();
    var visible = !!(el.offsetParent!==null || r.width||r.height);
    saves.push({
      text:t, onclick:(el.getAttribute('onclick')||''),
      disabled: !!el.disabled, visible:visible,
      w:Math.round(r.width), h:Math.round(r.height)
    });
  });

  // 첨부 화면 신호
  var attachBtn = document.querySelector("button[onclick*=\"popupAttachFile('A')\"]");
  return JSON.stringify({
    url: location.href,
    fields: fields,
    saves: saves,
    hasAttach: !!attachBtn
  });
})()
"""


def evaluate(ws_url, expr):
    ws = websocket.create_connection(ws_url, max_size=None, timeout=8,
                                     suppress_origin=True)
    try:
        ws.send(json.dumps({"id": 1, "method": "Runtime.enable"}))
        ws.recv()
        ws.send(json.dumps({
            "id": 2, "method": "Runtime.evaluate",
            "params": {"expression": expr, "returnByValue": True, "awaitPromise": True},
        }))
        while True:
            msg = json.loads(ws.recv())
            if msg.get("id") == 2:
                return msg
    finally:
        ws.close()


def main():
    targets = json.load(urllib.request.urlopen("http://127.0.0.1:9222/json", timeout=3))
    pages = [t for t in targets if t.get("type") == "page" and "sellerApplyform" in (t.get("url") or "")]
    if not pages:
        pages = [t for t in targets if t.get("type") == "page"]
    if not pages:
        print("no page target")
        return
    page = pages[0]
    expr = JS % json.dumps(FIELD_IDS)
    res = evaluate(page["webSocketDebuggerUrl"], expr)
    r = res.get("result", {}).get("result", {})
    if r.get("type") != "string":
        print("EVAL_ERR", json.dumps(res, ensure_ascii=False)[:500])
        return
    data = json.loads(r["value"])

    lines = []
    lines.append("URL: " + data["url"])
    lines.append("\n[ 필드 값 ]")
    empty_req = []
    for fid in FIELD_IDS:
        f = data["fields"].get(fid, {})
        if not f.get("exists"):
            mark = "   (필드 없음)"
            v = ""
        else:
            v = f.get("value", "")
            if f.get("tag") == "select" and f.get("text"):
                v = "%s (%s)" % (v, f.get("text"))
            mark = ""
        req = " *필수" if fid in REQUIRED else ""
        if fid in REQUIRED and f.get("exists") and not str(f.get("value", "")).strip():
            empty_req.append(fid)
        lines.append("  %-16s = %s%s%s" % (fid, v, req, mark))

    lines.append("\n[ 빈 필수칸 ] " + (", ".join(empty_req) if empty_req else "없음"))

    lines.append("\n[ '임시저장' 버튼 후보 ] %d개" % len(data["saves"]))
    for s in data["saves"]:
        lines.append("  · text='%s' visible=%s disabled=%s size=%dx%d onclick=%s"
                     % (s["text"], s["visible"], s["disabled"], s["w"], s["h"], s["onclick"]))

    lines.append("\n[ 첨부(팝업) 화면 신호 ] " + ("있음 → 이미 임시저장 넘어간 화면일 수 있음" if data["hasAttach"] else "없음(입력 폼 단계)"))

    # 진단 힌트
    lines.append("\n[ 진단 ]")
    goSave = [s for s in data["saves"] if "goSave(" in s["onclick"]
              and s["visible"] and not s["disabled"]
              and not any(a in s["text"] for a in ("신청", "제출", "최종", "지급", "우선순위", "취소", "삭제", "보완", "등록"))]
    if not data["saves"]:
        lines.append("  ⛔ 화면에 '임시저장' 버튼 자체가 없음 → 첨부화면이거나 폼이 아님")
    elif len(goSave) == 0:
        lines.append("  ⛔ goSave() 조건을 만족하는 클릭가능 버튼이 없음 (숨김/비활성/onclick 불일치)")
        lines.append("     → app.py _find_temp_save_buttons 가 후보 0개로 판단해 'notfound' 반환")
    elif len(goSave) > 1:
        lines.append("  ⛔ goSave() 후보가 %d개 → app.py 는 안전상 자동클릭 보류('ambiguous')" % len(goSave))
    else:
        lines.append("  ✅ goSave() 임시저장 버튼 1개 정상 → 자동클릭 가능 상태")
        if empty_req:
            lines.append("  ⚠ 단, 빈 필수칸(%s) 때문에 클릭해도 유효성검사 알럿으로 저장 차단될 수 있음" % ", ".join(empty_req))

    out = "\n".join(lines)
    open("_diag.txt", "w", encoding="utf-8").write(out)
    print("written _diag.txt")


if __name__ == "__main__":
    main()
