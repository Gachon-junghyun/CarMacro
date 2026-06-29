import re
import time
import difflib
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

from parser import parse_vertical

URL_SUFFIX = "sellerApplyform"

# ── 필드 사양 (검증·입력·되읽기 대조에 공통으로 사용) ────────────────
# kind: text / select / radio
# required: 필수 입력 여부
# enum: 허용 코드값 (정적 검증). model_cd 는 라이브 옵션으로 별도 대조.
# norm: 되읽기 대조 시 정규화 방식 ("digits"=숫자만, "lower"=소문자, None=공백정리)
FIELDS = [
    {"id": "req_kind",         "kind": "select", "required": True,  "enum": ["P", "B", "G"]},
    {"id": "contract_day",     "kind": "text",   "required": False, "norm": "digits"},
    {"id": "req_nm",           "kind": "text",   "required": True},
    {"id": "birth1",           "kind": "text",   "required": False, "norm": "digits"},
    {"id": "birth2",           "kind": "text",   "required": False, "norm": "digits"},
    {"id": "busi_no",          "kind": "text",   "required": False},
    {"id": "pri_busi_nm",      "kind": "text",   "required": False},
    {"id": "req_sex",          "kind": "radio",  "required": True,  "enum": ["M", "F"]},
    {"id": "model_cd",         "kind": "select", "required": True},   # 라이브 옵션 대조
    {"id": "req_cnt",          "kind": "text",   "required": True,  "norm": "digits"},
    {"id": "delivery_sch_day", "kind": "text",   "required": False, "norm": "digits"},
    # zipno/addr 은 주소검색 팝업으로 채움(직접 입력 불가, readonly) → FIELDS 에서 제외
    {"id": "addr_detail",      "kind": "text",   "required": False},
    {"id": "phone",            "kind": "text",   "required": False, "norm": "digits"},
    {"id": "mobile",           "kind": "text",   "required": True,  "norm": "digits"},
    {"id": "email",            "kind": "text",   "required": False, "norm": "lower"},
    {"id": "improve_fd_yn",    "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "first_buy_yn",     "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "social_yn",        "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "social_kind",      "kind": "select", "required": False},   # 텍스트 매칭(소상공인 등)
    {"id": "school_bus_yn",    "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "in_facility_yn",   "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "disaster_scrap_yn","kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "bms_yn",           "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "priority_type",    "kind": "radio",  "required": False,
     "enum": ["10", "20", "30", "40", "50", "00"]},
    {"id": "contact_nm",       "kind": "text",   "required": False},
    {"id": "contact_mobile",   "kind": "text",   "required": False, "norm": "digits"},
    {"id": "seller_mgrid",     "kind": "text",   "required": False},
]
FIELD_BY_ID = {f["id"]: f for f in FIELDS}
KNOWN_COLS = set(FIELD_BY_ID) | {"local_nm"}  # local_nm = (선택) 지자체 일치 확인용

MOBILE_RE = re.compile(r"^01[016789]-?\d{3,4}-?\d{4}$")
BIRTH6_RE = re.compile(r"^\d{6}$")
RRN_RE = re.compile(r"^\d{6}-?\d{7}$")
DATE_RE = re.compile(r"^\d{4}-?\d{2}-?\d{2}$")
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def detect_format(path):
    """세로형(라벨이 셀에 박힌 양식) vs 표형식 판별."""
    if str(path).lower().endswith(".csv"):
        return "tabular"
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        labels = {"계약일자", "신청유형", "성명", "신청차종", "신청대수", "출고예정일자"}
        found = 0
        for r in ws.iter_rows(values_only=True):
            for v in r:
                if v is not None and str(v).strip() in labels:
                    found += 1
                    if found >= 2:
                        return "vertical"
        return "tabular"
    except Exception:
        return "tabular"


def _norm(kind, value):
    s = str(value or "")
    if kind == "digits":
        return re.sub(r"\D", "", s)
    if kind == "lower":
        return s.strip().lower()
    return re.sub(r"\s+", " ", s.strip())


# ── 엑셀 정적 검증 ────────────────────────────────────────────
def validate_rows(rows, columns):
    """반환: (errors, warnings) — 각 항목은 사람이 읽는 문자열."""
    errors, warnings = [], []

    unknown = [c for c in columns if c not in KNOWN_COLS]
    if unknown:
        warnings.append(f"알 수 없는 열(무시됨): {', '.join(unknown)}")

    for i, row in enumerate(rows, start=1):
        def g(fid):
            return str(row.get(fid, "") or "").strip()

        # 필수값
        for f in FIELDS:
            if f["required"] and not g(f["id"]):
                # birth1/birth2 는 둘 중 하나만 있으면 됨 → 아래서 별도 처리
                errors.append(f"{i}행: 필수값 '{f['id']}' 비어있음")

        # 생년월일은 birth1 또는 birth2 중 하나 필요
        if not g("birth1") and not g("birth2"):
            errors.append(f"{i}행: 생년월일(birth1 또는 birth2) 둘 다 비어있음")

        # enum (코드값) 검증
        for f in FIELDS:
            v = g(f["id"])
            if v and "enum" in f and v not in f["enum"]:
                errors.append(
                    f"{i}행: '{f['id']}'='{v}' 는 허용값 {f['enum']} 아님")

        # 형식 검증
        if g("mobile") and not MOBILE_RE.match(g("mobile")):
            errors.append(f"{i}행: 휴대폰 형식 이상 '{g('mobile')}' (예 010-1234-5678)")
        if g("birth1") and not BIRTH6_RE.match(g("birth1")):
            warnings.append(f"{i}행: birth1 '{g('birth1')}' 6자리 숫자 아님")
        if g("birth2") and not RRN_RE.match(g("birth2")):
            warnings.append(f"{i}행: birth2 '{g('birth2')}' 주민번호 형식 아님")
        for df in ("contract_day", "delivery_sch_day"):
            if g(df) and not DATE_RE.match(g(df)):
                warnings.append(f"{i}행: {df} '{g(df)}' 날짜형식(YYYY-MM-DD) 아님")
        if g("email") and not EMAIL_RE.match(g("email")):
            warnings.append(f"{i}행: 이메일 형식 이상 '{g('email')}'")
        if g("req_cnt") and not g("req_cnt").isdigit():
            errors.append(f"{i}행: 신청대수 req_cnt '{g('req_cnt')}' 숫자 아님")

    return errors, warnings


# ── Selenium: 읽기 ────────────────────────────────────────────
def get_value(driver, fid):
    try:
        return (driver.find_element(By.ID, fid).get_attribute("value") or "").strip()
    except Exception:
        return ""


def get_select_options(driver, fid):
    try:
        el = driver.find_element(By.ID, fid)
        return [o.get_attribute("value") for o in el.find_elements(By.TAG_NAME, "option")]
    except Exception:
        return []


def get_radio_checked(driver, name):
    try:
        for el in driver.find_elements(By.NAME, name):
            if el.is_selected():
                return el.get_attribute("value")
    except Exception:
        pass
    return ""


# ── Selenium: 쓰기 (실패해도 예외 안 올림) ──────────────────────
_JS_SET = (
    "arguments[0].value=arguments[1];"
    "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
    "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));"
)


def fill_text(driver, fid, value):
    if not str(value or "").strip():
        return
    try:
        el = driver.find_element(By.ID, fid)
    except Exception:
        return
    v = str(value)
    try:
        if el.get_attribute("readonly") or not el.is_enabled():
            driver.execute_script(_JS_SET, el, v)
        else:
            el.clear()
            el.send_keys(v)
            driver.execute_script(
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
    except Exception:
        try:
            driver.execute_script(_JS_SET, el, v)
        except Exception:
            pass


def select_value(driver, fid, value):
    """반환: 'ok'|'skip'|'missing'(폼에 그 값 없음)|'error'."""
    if not str(value or "").strip():
        return "skip"
    try:
        el = driver.find_element(By.ID, fid)
        vals = [o.get_attribute("value") for o in el.find_elements(By.TAG_NAME, "option")]
        if str(value) not in vals:
            return "missing"
        Select(el).select_by_value(str(value))
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
        return "ok"
    except Exception:
        return "error"


def pick_radio(driver, name, value):
    """반환: 'ok'|'skip'|'missing'(그 값 없음)|'disabled'(비활성)|'error'."""
    if not str(value or "").strip():
        return "skip"
    try:
        found = False
        for el in driver.find_elements(By.NAME, name):
            if el.get_attribute("value") == str(value):
                found = True
                if not el.is_enabled():
                    return "disabled"
                driver.execute_script("arguments[0].click();", el)
                return "ok"
        return "missing" if not found else "error"
    except Exception:
        return "error"


def _squash(s):
    return re.sub(r"\s+", "", str(s or "")).lower()


def _option_list(driver, fid):
    """[(value, textContent)] 라이브 옵션."""
    try:
        el = driver.find_element(By.ID, fid)
        return [(o.get_attribute("value"),
                 (o.get_attribute("textContent") or "").strip())
                for o in el.find_elements(By.TAG_NAME, "option")]
    except Exception:
        return []


def match_option(opts, text):
    """옵션목록[(value,text)] 에서 한글 text 와 유일 매칭되는 (value, text) 반환.
    0개 또는 2개 이상이면 None (엉뚱한 선택 방지). 순수 함수 — 테스트용."""
    want = _squash(text)
    cands = [(v, t) for (v, t) in opts if v and want and want in _squash(t)]
    if len(cands) != 1:
        toks = [x for x in str(text).split() if x.strip()]
        cands = [(v, t) for (v, t) in opts
                 if v and toks and all(_squash(x) in _squash(t) for x in toks)]
    return cands[0] if len(cands) == 1 else None


def closest_option(opts, text):
    """가장 비슷한 옵션 (value, text, 유사도0~1). 옵션 없으면 None."""
    best, best_score = None, -1.0
    for v, t in opts:
        if not v:
            continue
        score = difflib.SequenceMatcher(None, _squash(text), _squash(t)).ratio()
        if score > best_score:
            best, best_score = (v, t), score
    return (best[0], best[1], best_score) if best else None


def _do_select(driver, fid, value):
    try:
        el = driver.find_element(By.ID, fid)
        Select(el).select_by_value(value)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
        return True
    except Exception:
        return False


def select_by_text(driver, fid, text):
    """한글 텍스트로 옵션 선택.
    반환: (status, 선택텍스트, 유사도, 옵션목록)
      status: 'ok'(정확) | 'fuzzy'(가장 비슷한 것) | 'none'(옵션 없음) | 'error'."""
    opts = _option_list(driver, fid)
    hit = match_option(opts, text)
    if hit is not None:
        return ("ok" if _do_select(driver, fid, hit[0]) else "error", hit[1], 1.0, opts)
    cl = closest_option(opts, text)
    if cl is None:
        return ("none", None, 0.0, opts)
    return ("fuzzy" if _do_select(driver, fid, cl[0]) else "error", cl[1], cl[2], opts)


def selected_option_text(driver, fid):
    try:
        el = driver.find_element(By.ID, fid)
        return (Select(el).first_selected_option.get_attribute("textContent") or "").strip()
    except Exception:
        return ""


def _find_addr_button(driver):
    """등록주소지용 '주소입력' 버튼(psPopupZip) 중 화면에 보이는 것."""
    btns = driver.find_elements(By.XPATH, "//button[contains(.,'주소입력')]")
    for b in btns:
        oc = b.get_attribute("onclick") or ""
        if "psPopupZip" in oc:
            try:
                if b.is_displayed():
                    return b
            except Exception:
                pass
    for b in btns:
        try:
            if b.is_displayed():
                return b
        except Exception:
            pass
    return None


def _dismiss_alert(driver):
    try:
        driver.switch_to.alert.accept()
        return True
    except Exception:
        return False


def fill_address(driver, keyword, log=None):
    """주소검색 팝업 자동화. 결과가 정확히 1개일 때만 자동 선택.
    반환: 'ok' | 'manual'(직접) | 'skip' | 'error'.
    행안부 juso 팝업: 검색어는 JS로 주입, 결과 선택은 setMaping(idx)+setParent() 호출."""
    if not str(keyword or "").strip():
        return "skip"
    main = driver.current_window_handle
    before = set(driver.window_handles)
    btn = _find_addr_button(driver)
    if btn is None:
        if log:
            log("   ⚠ '주소입력' 버튼을 못 찾음 → 주소는 직접 입력하세요")
        return "error"
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        time.sleep(0.2)
        btn.click()  # 네이티브 클릭(팝업 차단 회피)
    except Exception as e:
        if log:
            log(f"   ⚠ 주소입력 클릭 실패: {e} → 직접")
        return "error"
    # 팝업 창 대기
    popup = None
    for _ in range(20):
        time.sleep(0.3)
        new = set(driver.window_handles) - before
        if new:
            popup = new.pop()
            break
    if not popup:
        if log:
            log("   ⚠ 주소 팝업이 안 떴습니다 → 직접")
        return "error"
    try:
        driver.switch_to.window(popup)
        # 검색칸 등장 대기
        ready = False
        for _ in range(15):
            if driver.execute_script("return !!document.getElementById('keyword');"):
                ready = True
                break
            time.sleep(0.2)
        if not ready:
            raise RuntimeError("주소 검색칸을 못 찾음")
        # 검색어 JS 주입(send_keys 가 가끔 비어서 'E0005 검색어 없음' 경고 남)
        driver.execute_script(
            "var k=document.getElementById('keyword');k.value=arguments[0];"
            "k.dispatchEvent(new Event('input',{bubbles:true}));", str(keyword))
        driver.execute_script("$('#currentPage').val(1); searchUrlJuso();")
        # 결과 폴링: setMaping 링크 중 검색어와 일치하는 것
        want = _squash(keyword)
        idx, matched = None, 0
        for _ in range(15):
            time.sleep(0.3)
            _dismiss_alert(driver)
            cands = []
            for a in driver.find_elements(By.XPATH, "//a[contains(@href,'setMaping')]"):
                if a.text and want in _squash(a.text):
                    cands.append(a)
            if cands:
                matched = len(cands)
                m = re.search(r"setMaping\(\s*'?(\d+)'?\s*\)",
                              cands[0].get_attribute("href") or "")
                if m:
                    idx = m.group(1)
                break
        if idx is not None and matched == 1:
            driver.execute_script("setMaping(arguments[0]);", idx)  # 결과 선택
            time.sleep(0.4)
            driver.execute_script("setParent();")                   # 부모폼에 반영
            time.sleep(0.8)
            _dismiss_alert(driver)
            try:
                if popup in driver.window_handles:
                    driver.switch_to.window(popup)
                    driver.close()
            except Exception:
                pass
            driver.switch_to.window(main)
            z = get_value(driver, "zipno")
            ad = get_value(driver, "addr")
            if z or ad:
                if log:
                    log(f"   ✓ 주소 자동 선택 완료 (우편번호 {z}, {ad})")
                return "ok"
            if log:
                log("   ⚠ 주소 선택했으나 폼에 반영 안 됨 → 직접 확인")
            return "error"
        else:
            if log:
                log(f"   ⚠ 주소 검색결과 {matched}개 → 팝업에서 직접 선택하세요(자동 보류)")
            try:
                driver.switch_to.window(main)
            except Exception:
                pass
            return "manual"
    except Exception as e:
        _dismiss_alert(driver)
        if log:
            log(f"   ⚠ 주소 자동화 오류: {e} → 직접")
        try:
            driver.switch_to.window(main)
        except Exception:
            pass
        return "error"


def fill_one(driver, row, log=None):
    text_sel = row.get("_text_select", set())
    for f in FIELDS:
        fid = f["id"]
        v = row.get(fid)
        if not str(v or "").strip():
            continue
        if fid in text_sel:
            status, chosen, score, opts = select_by_text(driver, fid, v)
            if status == "fuzzy" and log:
                log(f"❌ '{fid}' 정확 일치 없음: '{v}' → 가장 비슷한 '{chosen}' "
                    f"(유사도 {score:.0%}) 선택함. 반드시 직접 확인!", color="red")
            elif status == "none" and log:
                log(f"❌ '{fid}' 선택할 옵션이 없음: '{v}' → 직접 선택 필요", color="red")
            elif status == "error" and log:
                log(f"❌ '{fid}' 선택 오류: '{v}' → 직접 선택 필요", color="red")
        elif f["kind"] == "text":
            fill_text(driver, fid, v)
        elif f["kind"] == "select":
            st = select_value(driver, fid, v)
            if st == "missing" and log:
                log(f"❌ '{fid}' 값 '{v}' 이(가) 이 폼 선택지에 없음 → 직접 확인", color="red")
            elif st == "error" and log:
                log(f"❌ '{fid}' 선택 오류 '{v}' → 직접 확인", color="red")
        else:
            st = pick_radio(driver, fid, v)
            if st == "disabled" and log:
                log(f"❌ '{fid}' 값 '{v}' 이(가) 이 폼에서 비활성(선택 불가·마감 등) "
                    f"→ 직접 확인", color="red")
            elif st == "missing" and log:
                log(f"❌ '{fid}' 값 '{v}' 이(가) 이 폼에 없음 → 직접 확인", color="red")
        # 조건부 섹션(사회계층 등) 노출 대기
        if fid == "social_yn" and str(v).strip() == "Y":
            time.sleep(0.4)


# ── 되읽기 대조 (핵심 안전장치) ────────────────────────────────
def verify_fill(driver, row):
    """폼에 실제로 들어간 값을 읽어 엑셀 값과 대조. 반환: 불일치 리스트."""
    text_sel = row.get("_text_select", set())
    mismatches = []
    for f in FIELDS:
        fid = f["id"]
        want_raw = str(row.get(fid, "") or "").strip()
        if not want_raw:
            continue  # 빈 값은 건너뜀(채우지 않았으므로)
        if fid in text_sel:
            # 텍스트매칭 필드: 선택된 옵션 텍스트가 원하는 한글을 포함하는지
            got_txt = selected_option_text(driver, fid)
            want = _squash(want_raw)
            toks = [x for x in want_raw.split() if x.strip()]
            ok = bool(got_txt) and (
                (want and want in _squash(got_txt))
                or (toks and all(_squash(x) in _squash(got_txt) for x in toks)))
            if not ok:
                mismatches.append((fid, want_raw, got_txt or "(선택 안됨)"))
            continue
        if f["kind"] == "select":
            got = get_value(driver, fid)
        elif f["kind"] == "radio":
            got = get_radio_checked(driver, fid)
        else:
            got = get_value(driver, fid)
        nk = f.get("norm")
        if f["kind"] in ("select", "radio"):
            ok = (got == want_raw)
        else:
            ok = (_norm(nk, got) == _norm(nk, want_raw))
        if not ok:
            mismatches.append((fid, want_raw, got or "(빈칸)"))
    return mismatches


# ── 워커 ──────────────────────────────────────────────────────
class Worker(threading.Thread):
    def __init__(self, rows, log_q, state_q, cmd_q, auto_flag, addr_flag,
                 code_flag, code_confirm_flag):
        super().__init__(daemon=True)
        self.rows = rows
        self.log_q = log_q
        self.state_q = state_q
        self.cmd_q = cmd_q
        self.auto_flag = auto_flag
        self.addr_flag = addr_flag
        self.code_flag = code_flag            # 확인코드 자동입력 on/off
        self.code_confirm_flag = code_confirm_flag  # 확인(저장)까지 자동 on/off
        self.idx = 0
        self.running = True
        self.driver = None
        self.blank_handled = False
        self.main_handle = None
        self._known_handles = set()    # 이미 본 창들(새 창만 검사 → 포커스 도둑질 방지)
        self._popup_focus = None       # 사용자가 [확인] 누를 보안팝업(포커스 유지)

    def log(self, msg, color=None):
        self.log_q.put((time.strftime("[%H:%M:%S] ") + msg, color))

    def set_state(self, **kw):
        self.state_q.put(kw)

    def on_form(self):
        try:
            return self.driver.current_url.endswith(URL_SUFFIX)
        except Exception:
            return False

    def do_fill(self, reason):
        if self.idx >= len(self.rows):
            self.log("더 입력할 건이 없습니다.")
            return
        row = self.rows[self.idx]
        self.blank_handled = True  # 재시도 스팸 방지

        # 0) 폼 페이지인지
        if not self.on_form():
            self.log("⚠ 신청서 작성 폼이 아닙니다. 폼으로 이동 후 다시 시도하세요.")
            self.set_state(status="중단 — 신청서 폼 아님")
            return

        # 0-1) 전기/수소 폼-데이터 일치 확인 (열린 폼 URL: /h2/ = 수소)
        want_kind = str(row.get("car_kind", "") or "").strip()
        if want_kind:
            try:
                form_kind = "수소" if "/h2/" in self.driver.current_url else "전기"
            except Exception:
                form_kind = ""
            if form_kind and want_kind != form_kind:
                self.log(f"⛔ 데이터는 '{want_kind}'인데 열린 폼은 '{form_kind}' → 입력 거부! "
                         f"올바른 {want_kind} 신청서 폼으로 이동하세요.", color="red")
                self.set_state(status=f"⛔ 차종구분 불일치 ({want_kind}≠{form_kind})")
                return

        # 1) 지자체 일치 확인 (엑셀에 local_nm 열이 있으면)
        detected = get_value(self.driver, "local_nm")
        want_local = str(row.get("local_nm", "") or "").strip()
        if want_local and detected and want_local != detected:
            self.log(f"⛔ 지자체 불일치! 화면='{detected}' / 엑셀='{want_local}' → 입력 거부")
            self.set_state(status=f"⛔ 지자체 불일치로 중단 ({detected}≠{want_local})")
            return

        # 2) 차종 코드가 실제 옵션에 있는지 (코드값 모드일 때만; 텍스트매칭은 fill에서 처리)
        text_sel = row.get("_text_select", set())
        want_model = str(row.get("model_cd", "") or "").strip()
        if want_model and "model_cd" not in text_sel:
            opts = get_select_options(self.driver, "model_cd")
            if opts and want_model not in opts:
                self.log(f"⛔ 차종코드 '{want_model}' 가 이 페이지 선택목록에 없음 → 입력 거부")
                self.set_state(status="⛔ 차종코드 없음으로 중단")
                return

        # 3) 입력
        self.log(f"{reason}: {self.idx+1}번째 입력 시작 (지자체 {detected or '(미선택)'})")
        self.set_state(status=f"입력 중… ({detected})")
        try:
            fill_one(self.driver, row, log=self.log)
        except Exception as e:
            self.log(f"입력 오류: {e}")
            self.set_state(status="입력 오류 (로그 확인)")
            return

        # 3-1) 주소검색 팝업 자동화 (도로명 주소가 있고 옵션 켜졌을 때)
        if self.addr_flag[0] and str(row.get("addr", "") or "").strip():
            st = fill_address(self.driver, row.get("addr"), log=self.log)
            if st == "ok":
                # 상세주소는 팝업 후 다시 채움(덮어쓰기 방지)
                fill_text(self.driver, "addr_detail", row.get("addr_detail"))

        # 4) 되읽기 대조
        time.sleep(0.3)
        mm = verify_fill(self.driver, row)
        self.idx += 1
        if mm:
            self.log(f"❌ {self.idx}번째 검증 실패 — 불일치 {len(mm)}건:")
            for fid, want, got in mm:
                self.log(f"    · {fid}: 기대 '{want}' / 실제 '{got}'")
            self.set_state(status=f"❌ 검증 실패 {len(mm)}건 — 저장 말고 직접 확인!",
                           progress=f"{self.idx}/{len(self.rows)}")
        else:
            self.log(f"✅ {self.idx}번째 입력·검증 통과. 검토 후 저장/제출은 직접.")
            self.set_state(status="✅ 검증 통과 — 검토 후 다음 폼으로",
                           progress=f"{self.idx}/{len(self.rows)}")

    def run(self):
        # 연결 검증
        try:
            opts = webdriver.ChromeOptions()
            opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            self.driver = webdriver.Chrome(options=opts)
        except Exception as e:
            self.log(f"크롬 연결 실패: {e}")
            self.set_state(status="크롬 연결 실패 (포트 9222 확인)")
            return

        # 작업 대상(메인 폼) 창 기억 — 보안팝업 처리 후 복귀용
        self.main_handle = self.driver.current_window_handle
        if not self.on_form():
            for h in self.driver.window_handles:
                try:
                    self.driver.switch_to.window(h)
                    if self.driver.current_url.endswith(URL_SUFFIX):
                        self.main_handle = h
                        break
                except Exception:
                    pass
            try:
                self.driver.switch_to.window(self.main_handle)
            except Exception:
                pass

        self.log(f"연결 완료. URL: {self.driver.current_url}")
        if not self.on_form():
            self.log("⚠ 지금 화면이 신청서 작성 폼이 아닙니다. 지자체 선택 후 폼으로 이동하세요.")
        if get_value(self.driver, "local_nm"):
            self.log(f"지자체 감지: {get_value(self.driver, 'local_nm')}")
        self.log(f"총 {len(self.rows)}건 대기 중.")
        self.set_state(status="대기 중 — 신청서 폼에서 입력 준비",
                       progress=f"{self.idx}/{len(self.rows)}")
        self._known_handles = set(self.driver.window_handles)
        self.precheck_models()

        while self.running:
            # 보안 확인코드 팝업: 새 창이 생겼을 때만 검사(평소엔 창 전환 0회)
            if self.code_flag[0]:
                try:
                    self.handle_random_popup()
                except Exception:
                    pass

            # 사용자가 [확인] 누를 보안팝업이 떠 있으면 폼 폴링/전환을 멈춰 포커스 양보
            if self._popup_focus:
                try:
                    if self._popup_focus not in self.driver.window_handles:
                        self._popup_focus = None
                except Exception:
                    self._popup_focus = None
                if self._popup_focus:
                    self.set_state(status="🔐 보안팝업: 역순 입력됨 — [확인] 직접 누르세요")
                    time.sleep(0.5)
                    continue

            try:
                while True:
                    cmd = self.cmd_q.get_nowait()
                    if cmd == "fill":
                        self.do_fill("수동 입력")
                    elif cmd == "reverify":
                        self.reverify()
            except queue.Empty:
                pass

            local = get_value(self.driver, "local_nm")
            req_nm = get_value(self.driver, "req_nm")
            self.set_state(local=local or "(미선택)")

            blank_form = self.on_form() and local != "" and req_nm == ""
            if not blank_form:
                self.blank_handled = False

            if (self.auto_flag[0] and blank_form and not self.blank_handled
                    and self.idx < len(self.rows)):
                self.do_fill("자동 감지(빈 폼)")

            time.sleep(0.5)

    def handle_random_popup(self):
        """새로 뜬 창만 검사. 보안 확인코드 팝업이면 코드 역순을 자동 입력.
        평소(새 창 없음)엔 창 전환을 하지 않아 포커스를 뺏지 않는다."""
        try:
            handles = set(self.driver.window_handles)
        except Exception:
            return
        new = handles - self._known_handles
        if not new:
            self._known_handles &= handles   # 닫힌 창 정리
            return
        self._known_handles = handles
        stay_on = None
        for h in new:
            try:
                self.driver.switch_to.window(h)
                url = self.driver.current_url
            except Exception:
                continue
            if "RandomChk" not in url and "popupSellerRandom" not in url:
                continue  # 사용자가 연 다른 창 — 건드리지 않음
            # 화면 코드: goCompare 소스의 실제 코드 우선(검증이 쓰는 값)
            code = ""
            src = self.driver.execute_script(
                "try{return goCompare.toString()}catch(e){return ''}")
            m = re.search(r"=\s*'([0-9A-Za-z]{6,16})'\s*\.split", src)
            if m:
                code = m.group(1)
            else:
                cands = self.driver.execute_script(
                    "return [...document.querySelectorAll('span.guide,span,div,b')]"
                    ".map(e=>(e.textContent||'').trim())"
                    ".filter(t=>/^[0-9A-Za-z]{6,16}$/.test(t));")
                code = cands[0] if cands else ""
            if not code:
                self.log("🔐 확인코드 팝업 감지했으나 코드를 못 읽음 → 직접 입력", color="red")
                continue
            rev = code[::-1]
            self.driver.execute_script(
                "var e=document.getElementById('randeomChk');"
                "if(e){e.value=arguments[0];e.dispatchEvent(new Event('input',{bubbles:true}));}", rev)
            self.log(f"🔐 확인코드 '{code}' 감지 → 역순 '{rev}' 자동 입력")
            if self.code_confirm_flag[0]:
                self.driver.execute_script("try{goCompare();}catch(e){}")
                self.log("   [확인] 자동 클릭 → 저장 진행됨", color="red")
            else:
                self.log("   역순 입력 완료 — 이 팝업의 [확인]을 직접 누르세요(저장)")
                stay_on = h   # 팝업에 포커스 남겨 사용자가 바로 확인
        # 포커스 정리: 수동확인이면 팝업에, 아니면 메인 폼으로
        try:
            if stay_on:
                self.driver.switch_to.window(stay_on)
                self._popup_focus = stay_on
            elif self.main_handle:
                self.driver.switch_to.window(self.main_handle)
        except Exception:
            pass

    def precheck_models(self):
        """대기 중 시점에 차종 목록을 가져와 데이터 차종이 있는지 미리 검사."""
        opts = _option_list(self.driver, "model_cd")
        real = [(v, t) for v, t in opts if v]
        for i, row in enumerate(self.rows, start=1):
            ts = row.get("_text_select", set())
            if "model_cd" not in ts:
                continue
            want = str(row.get("model_cd", "") or "").strip()
            if not want:
                continue
            if not real:
                self.log(f"[{i}건] 차종 목록이 아직 안 보입니다(공고/폼 진입 후 자동 재확인): '{want}'")
                continue
            hit = match_option(opts, want)
            if hit:
                self.log(f"[{i}건] 차종 확인 ✓ '{want}' → 목록에 있음 ({hit[0]})")
            else:
                cl = closest_option(opts, want)
                if cl:
                    self.log(f"[{i}건] ❌ 차종 '{want}' 목록에 없음 → 가장 비슷한 "
                             f"'{cl[1]}' ({cl[0]}, 유사도 {cl[2]:.0%}) 사용 예정. 반드시 직접 확인!",
                             color="red")
                else:
                    self.log(f"[{i}건] ❌ 차종 선택 옵션이 없습니다: '{want}'", color="red")

    def reverify(self):
        """현재 폼을 직전 입력 건과 다시 대조 (저장 직전 점검용)."""
        if self.idx == 0:
            self.log("재검증할 직전 입력 건이 없습니다.")
            return
        row = self.rows[self.idx - 1]
        mm = verify_fill(self.driver, row)
        if mm:
            self.log(f"❌ 재검증 불일치 {len(mm)}건:")
            for fid, want, got in mm:
                self.log(f"    · {fid}: 기대 '{want}' / 실제 '{got}'")
            self.set_state(status=f"❌ 재검증 실패 {len(mm)}건")
        else:
            self.log("✅ 재검증 통과 — 폼 값이 엑셀과 일치.")
            self.set_state(status="✅ 재검증 통과")

    def stop(self):
        self.running = False


# ── GUI ───────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EV 보조금 신청 자동입력 — 검증판")
        self.geometry("520x560")
        self.worker = None
        self.rows = []
        self.valid = False
        self.log_q = queue.Queue()
        self.state_q = queue.Queue()
        self.cmd_q = queue.Queue()
        self.auto_flag = [True]
        self.addr_flag = [True]
        self.code_flag = [True]           # 확인코드 자동입력
        self.code_confirm_flag = [False]  # 확인(저장)까지 자동 — 기본 끔

        ttk.Label(self, text="① 크롬을 디버그 포트(9222)로 띄우고 로그인",
                  font=("", 9)).pack(anchor="w", padx=10, pady=(10, 0))
        ttk.Label(self, text="② 엑셀 불러오기(자동 검증) → 시작 → 지자체 선택 후 폼 진입",
                  font=("", 9)).pack(anchor="w", padx=10)

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10, pady=8)
        ttk.Button(bar, text="엑셀 불러오기", command=self.load_excel).pack(side="left")
        self.start_btn = ttk.Button(bar, text="시작", command=self.start, state="disabled")
        self.start_btn.pack(side="left", padx=5)
        ttk.Button(bar, text="정지", command=self.stop).pack(side="left")

        bar2 = ttk.Frame(self)
        bar2.pack(fill="x", padx=10, pady=(0, 4))
        self.fill_btn = ttk.Button(bar2, text="▶ 다음 건 입력", command=self.manual_fill,
                                   state="disabled")
        self.fill_btn.pack(side="left")
        self.reverify_btn = ttk.Button(bar2, text="현재 폼 재검증", command=self.manual_reverify,
                                       state="disabled")
        self.reverify_btn.pack(side="left", padx=5)
        self.auto_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar2, text="빈 폼 뜨면 자동 입력", variable=self.auto_var,
                        command=self.toggle_auto).pack(side="left", padx=8)
        self.addr_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar2, text="주소 자동검색", variable=self.addr_var,
                        command=self.toggle_addr).pack(side="left")

        bar3 = ttk.Frame(self)
        bar3.pack(fill="x", padx=10, pady=(0, 4))
        self.code_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar3, text="확인코드(보안문자) 역순 자동입력", variable=self.code_var,
                        command=self.toggle_code).pack(side="left")
        self.code_confirm_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar3, text="확인까지 자동(=저장 실행)", variable=self.code_confirm_var,
                        command=self.toggle_code).pack(side="left", padx=10)

        info = ttk.LabelFrame(self, text="현재 상태")
        info.pack(fill="x", padx=10, pady=4)
        self.v_status = tk.StringVar(value="대기")
        self.v_local = tk.StringVar(value="(미선택)")
        self.v_prog = tk.StringVar(value="0/0")
        for label, var in [("상태", self.v_status),
                           ("감지된 지자체", self.v_local),
                           ("진행", self.v_prog)]:
            r = ttk.Frame(info)
            r.pack(fill="x", padx=6, pady=2)
            ttk.Label(r, text=label, width=12).pack(side="left")
            ttk.Label(r, textvariable=var, font=("", 10, "bold")).pack(side="left")

        ttk.Label(self, text="로그").pack(anchor="w", padx=10)
        self.log_box = tk.Text(self, height=14, state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_box.tag_config("red", foreground="#cc0000")
        self.log_box.tag_config("green", foreground="#117711")

        self.after(200, self.poll_queues)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
        if not path:
            return
        if detect_format(path) == "vertical":
            self.load_vertical(path)
            return
        try:
            df = (pd.read_csv(path, dtype=str) if path.endswith(".csv")
                  else pd.read_excel(path, dtype=str)).fillna("")
        except Exception as e:
            self._log(f"❌ 엑셀 읽기 실패: {e}")
            return
        rows = df.to_dict("records")
        cols = list(df.columns)

        errors, warnings = validate_rows(rows, cols)
        self._log(f"── 엑셀 검증: {len(rows)}건 ──")
        for w in warnings:
            self._log(f"  ⚠ {w}")
        for e in errors:
            self._log(f"  ❌ {e}")

        if errors:
            self.rows, self.valid = [], False
            self.start_btn.config(state="disabled")
            self._log(f"검증 실패: 오류 {len(errors)}건 → 엑셀 수정 후 다시 불러오세요. (시작 비활성)")
            self.v_status.set(f"엑셀 검증 실패 ({len(errors)}건)")
            self.v_prog.set(f"0/{len(rows)}")
        else:
            self.rows, self.valid = rows, True
            self.start_btn.config(state="normal")
            self._log(f"✅ 검증 통과 ({len(rows)}건"
                      + (f", 경고 {len(warnings)}건" if warnings else "") + "). 시작 가능.")
            self.v_status.set("엑셀 검증 통과 — 시작 가능")
            self.v_prog.set(f"0/{len(rows)}")

    def load_vertical(self, path):
        try:
            res = parse_vertical(path)
        except Exception as e:
            self._log(f"❌ 세로형 엑셀 파싱 실패: {e}")
            return
        mapped = res["mapped"]
        special = res["special"]
        row = dict(mapped)
        ts = set()
        for k, v in special.items():
            row[k] = v
            ts.add(k)
        row["_text_select"] = ts

        self._log("── 세로형 양식 1건 파싱 ──")
        self._log("[자동 입력 대상]")
        for k, v in mapped.items():
            self._log(f"   {k} = {v}")
        for k, v in special.items():
            self._log(f"   {k} = {v}  (페이지 옵션 텍스트매칭)")
        self._log("[🖐 사람이 직접 확인/입력]")
        for lbl, val, reason in res["handoff"]:
            self._log(f"   · {lbl} = '{val}'  ({reason})")
        for n in res["notes"]:
            self._log(f"   {n}")

        self.rows = [row]
        self.valid = True
        self.start_btn.config(state="normal")
        self.v_status.set("세로형 1건 로드 — 시작 가능 (위 handoff 직접 확인)")
        self.v_prog.set("0/1")
        self._log("✅ 로드 완료. [시작] 후 폼에서 자동 입력됩니다.")

    def start(self):
        if not (self.rows and self.valid):
            self._log("시작 불가: 검증 통과한 엑셀이 없습니다.")
            return
        if self.worker and self.worker.is_alive():
            return
        self.worker = Worker(self.rows, self.log_q, self.state_q,
                             self.cmd_q, self.auto_flag, self.addr_flag,
                             self.code_flag, self.code_confirm_flag)
        self.worker.start()
        self.fill_btn.config(state="normal")
        self.reverify_btn.config(state="normal")
        self._log("워커 시작")

    def stop(self):
        if self.worker:
            self.worker.stop()
            self.fill_btn.config(state="disabled")
            self.reverify_btn.config(state="disabled")
            self._log("정지 요청")

    def manual_fill(self):
        if self.worker and self.worker.is_alive():
            self.cmd_q.put("fill")

    def manual_reverify(self):
        if self.worker and self.worker.is_alive():
            self.cmd_q.put("reverify")

    def toggle_auto(self):
        self.auto_flag[0] = self.auto_var.get()
        self._log(f"자동 입력 {'켜짐' if self.auto_flag[0] else '꺼짐'}")

    def toggle_addr(self):
        self.addr_flag[0] = self.addr_var.get()
        self._log(f"주소 자동검색 {'켜짐' if self.addr_flag[0] else '꺼짐'}")

    def toggle_code(self):
        self.code_flag[0] = self.code_var.get()
        self.code_confirm_flag[0] = self.code_confirm_var.get()
        msg = f"확인코드 자동입력 {'켜짐' if self.code_flag[0] else '꺼짐'}"
        if self.code_confirm_flag[0]:
            msg += " / 확인까지 자동(저장 실행됨)"
        self._log(msg)

    def poll_queues(self):
        while not self.log_q.empty():
            item = self.log_q.get()
            if isinstance(item, tuple):
                self._log(item[0], item[1])
            else:
                self._log(item)
        while not self.state_q.empty():
            st = self.state_q.get()
            if "status" in st:
                self.v_status.set(st["status"])
            if "local" in st:
                self.v_local.set(st["local"])
            if "progress" in st:
                self.v_prog.set(st["progress"])
        self.after(200, self.poll_queues)

    def _log(self, msg, color=None):
        self.log_box.config(state="normal")
        if color:
            self.log_box.insert("end", msg + "\n", color)
        else:
            self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
