import os
import re
import time
import json
import difflib
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog
import openpyxl
# pandas 는 표형식(CSV/일반 xlsx) 읽기에만 쓰여 무겁고 선택적 → load_excel 에서 지연 import.
# 실제 워크플로(세로형 양식)는 parser.py(openpyxl)만 사용하므로 exe 에 pandas 불필요.
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

from parser import parse_vertical
import browse_control   # 지자체 선택 → 신청서 폼 진입(과부하 재시도) 자동화

URL_SUFFIX = "sellerApplyform"

# ── 필드 사양 (검증·입력·되읽기 대조에 공통으로 사용) ────────────────
# kind: text / select / radio
# required: 필수 입력 여부
# enum: 허용 코드값 (정적 검증). model_cd 는 라이브 옵션으로 별도 대조.
# norm: 되읽기 대조 시 정규화 방식 ("digits"=숫자만, "lower"=소문자, None=공백정리)
FIELDS = [
    {"id": "req_kind",         "kind": "select", "required": True,  "enum": ["P", "B", "G"]},
    {"id": "contract_day",     "kind": "text",   "required": False, "norm": "digits"},
    {"id": "req_nm",           "kind": "text",   "required": True},   # 개인=성명 / 단체=기관명
    {"id": "ceo",              "kind": "text",   "required": False},  # 단체 전용: 대표자
    {"id": "grp_reqst_se",     "kind": "select", "required": False,   # 단체 전용: 신청구분
     "enum": ["01", "02", "99"]},
    {"id": "birth1",           "kind": "text",   "required": False, "norm": "digits"},
    {"id": "birth2",           "kind": "text",   "required": False, "norm": "digits"},  # 개인=주민번호 / 단체=법인등록번호
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
    {"id": "bms_yn",           "kind": "radio",  "required": False, "enum": ["Y", "N"]},   # 승용 전용
    {"id": "taxi_yn",          "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    # ↓ 전기 화물(car_type=12) 전용 — 승용 폼엔 없음(없으면 'missing' 로그만, 데이터 있을 때만 입력)
    {"id": "farmng_yn",        "kind": "radio",  "required": False, "enum": ["Y", "N"]},   # 농업인 여부
    {"id": "hdry_yn",          "kind": "radio",  "required": False, "enum": ["Y", "N"]},   # 택배여부
    # 경유화물차 보유 미이행자 여부(화물 전용) — Y 면 보유차량 정보 동적행 생성(fill_thrgh_hold)
    {"id": "thrgh_ex_yn",      "kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "exchange_3year_yn","kind": "radio",  "required": False, "enum": ["Y", "N"]},
    {"id": "priority_type",    "kind": "radio",  "required": False,
     "enum": ["10", "20", "30", "40", "50", "00"]},
    {"id": "contact_nm",       "kind": "text",   "required": False},
    {"id": "contact_mobile",   "kind": "text",   "required": False, "norm": "digits"},
    {"id": "seller_mgrid",     "kind": "text",   "required": False},
]
FIELD_BY_ID = {f["id"]: f for f in FIELDS}
# local_nm 지자체확인 / car_kind 전기·수소 가드 / addr 주소검색 /
# exchange_scrap·exchange_3year_cnt 전환지원금 폐차 정보(특수 처리)
KNOWN_COLS = set(FIELD_BY_ID) | {"local_nm", "car_kind", "addr",
                                 "exchange_scrap", "exchange_3year_cnt",
                                 "thrgh_hold", "thrgh_ex_cnt"}

# 전환지원금 폐차 정보: 라이브 폼의 name 기반 동적 필드 → 한글 라벨(로그·검증용)
SCRAP_LABELS = {"bf_owner_nm": "직전소유주", "exchg_delivery_day": "차량최초등록일",
                "own_start_dt": "소유기간시작일", "own_end_dt": "소유기간종료일",
                "fuel": "유종", "exchg_vh_num": "폐차차량번호"}

# 폼에 원하는 값이 없을 때 대체값 (예: 수소엔 개인사업자(B)가 없어 개인(P)로)
SELECT_FALLBACK = {"req_kind": {"B": "P"}}

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
                # 단체(G)는 성별(개인 항목)이 없음 → 필수 제외
                if f["id"] == "req_sex" and g("req_kind") == "G":
                    continue
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
            if st == "missing":
                fb = SELECT_FALLBACK.get(fid, {}).get(str(v))
                if fb and select_value(driver, fid, fb) == "ok":
                    row[fid] = fb   # 검증도 대체값 기준으로
                    st = "ok"
                    if log:
                        log(f"ℹ '{fid}' 값 '{v}' 이 폼에 없어 '{fb}'(으)로 대체 입력", color="red")
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


# ── 전환지원금 폐차 정보 (동적 name 기반 섹션) ──────────────────
# 라이브 폼: exchange_3year_yn=Y → '폐차대수'(exchange_3year_cnt) 입력 →
#   [확인] = create3yearNewCarInfo() 가 '전환지원금 폐차 정보' 행을 생성.
# 생성된 행 입력칸은 id 가 없고 name 으로만 접근(같은 name 의 숨은 템플릿 존재 →
#   화면에 '보이는' 요소만 골라 채운다).
_SCRAP_GEN_JS = r"""
var cnt = String(arguments[0] || '1');
var c = document.getElementById('exchange_3year_cnt');
if(c){ c.value=cnt;
  c.dispatchEvent(new Event('input',{bubbles:true}));
  c.dispatchEvent(new Event('change',{bubbles:true})); }
// create3yearNewCarInfo 가 confirm/alert 를 띄워도 멈추지 않게 잠깐 가로챔
var _cf=window.confirm, _al=window.alert;
window.confirm=function(){return true;}; window.alert=function(){};
try{ create3yearNewCarInfo(); }catch(e){ window.confirm=_cf; window.alert=_al; return 'ERR:'+e; }
window.confirm=_cf; window.alert=_al;
return 'ok';
"""

_SCRAP_VISEL_JS = r"""
function visEl(name){
  var els=document.getElementsByName(name);
  for(var i=0;i<els.length;i++){var e=els[i];var s=getComputedStyle(e);var r=e.getBoundingClientRect();
    if(!(s.display==='none'||s.visibility==='hidden'||(r.width===0&&r.height===0))) return e;}
  return els.length?els[0]:null;
}
"""

_SCRAP_FILL_JS = _SCRAP_VISEL_JS + r"""
var data=arguments[0]; var out={};
function setV(el,v){el.value=v;
  el.dispatchEvent(new Event('input',{bubbles:true}));
  el.dispatchEvent(new Event('change',{bubbles:true}));}
Object.keys(data).forEach(function(name){
  var val=data[name]; if(val===''||val==null) return;
  var el=visEl(name); if(!el){ out[name]='NULL'; return; }
  if((el.tagName||'')==='SELECT'){
    var matched=false;
    for(var i=0;i<el.options.length;i++){ if(el.options[i].value===val){el.selectedIndex=i;matched=true;break;} }
    if(!matched){ for(var j=0;j<el.options.length;j++){ if((el.options[j].text||'').indexOf(val)>=0){el.selectedIndex=j;matched=true;break;} } }
    el.dispatchEvent(new Event('change',{bubbles:true}));
    out[name]=el.value;
  } else { setV(el,val); out[name]=el.value; }
});
return JSON.stringify(out);
"""

_SCRAP_READ_JS = _SCRAP_VISEL_JS + r"""
var names=arguments[0]; var out={};
names.forEach(function(n){ var e=visEl(n); out[n]= e? (e.value||'') : ''; });
return JSON.stringify(out);
"""


def fill_exchange_scrap(driver, row, log=None):
    """전환지원금 폐차 정보 행을 생성하고 채운다. (exchange_3year_yn=Y 전용)"""
    scrap = row.get("exchange_scrap") or {}
    if not scrap:
        return
    cnt = str(row.get("exchange_3year_cnt") or "1")
    try:
        r = driver.execute_script(_SCRAP_GEN_JS, cnt)
    except Exception as e:
        r = "ERR:%s" % e
    if str(r).startswith("ERR"):
        if log:
            log(f"❌ 전환지원금 폐차행 생성 실패({r}) → 폐차대수·확인 직접", color="red")
        return
    # 행 등장 대기
    for _ in range(20):
        try:
            if driver.execute_script("return document.getElementsByName('bf_owner_nm').length>0;"):
                break
        except Exception:
            pass
        time.sleep(0.3)
    # 채우기(동적 재렌더 대비 재시도)
    res = {}
    for _ in range(6):
        try:
            res = json.loads(driver.execute_script(_SCRAP_FILL_JS, scrap))
        except Exception:
            res = {}
        if res.get("bf_owner_nm") not in (None, "NULL"):
            break
        time.sleep(0.5)
    nulls = [SCRAP_LABELS.get(k, k) for k, v in res.items() if v == "NULL"]
    if log:
        if nulls:
            log(f"❌ 전환지원금 폐차정보 일부 미입력: {', '.join(nulls)} → 직접 확인", color="red")
        else:
            log(f"   ✓ 전환지원금 폐차정보 {len(scrap)}칸 입력 (폐차대수 {cnt})", color="green")


def verify_exchange_scrap(driver, row):
    """폐차 정보 되읽기 대조. 반환: 불일치 [(라벨, 기대, 실제)]."""
    scrap = row.get("exchange_scrap") or {}
    if not scrap:
        return []
    try:
        got = json.loads(driver.execute_script(_SCRAP_READ_JS, list(scrap.keys())))
    except Exception:
        return [("전환지원금폐차정보", "(읽기실패)", "")]
    mm = []
    for k, want in scrap.items():
        g = str(got.get(k, "") or "")
        if k in ("exchg_delivery_day", "own_start_dt", "own_end_dt"):
            ok = re.sub(r"\D", "", g) == re.sub(r"\D", "", str(want))
        else:
            ok = g.strip() == str(want).strip()
        if not ok:
            mm.append((SCRAP_LABELS.get(k, k), want, g or "(빈칸)"))
    return mm


# ── 경유화물차 보유 미이행자 정보 (thrgh_ex_yn=Y, 동적 name 기반) ─────────
# 라이브 폼: thrgh_ex_yn=Y → '경유화물차 보유대수'(thrgh_ex_cnt) 입력 →
#   createNewCarInfoHold() 가 보유차량 입력행을 생성. 생성칸은 id 없이 name 으로만:
#   ex_vh_num_hold(차량번호) / ex_vh_id_hold(차대번호) / ex_model_nm_hold(보유모델명)
#   (전환지원금 폐차정보와 동일 구조 → _SCRAP_FILL_JS/_SCRAP_READ_JS 재사용)
HOLD_LABELS = {"ex_vh_num_hold": "보유차량번호", "ex_vh_id_hold": "차대번호",
               "ex_model_nm_hold": "보유모델명"}

_HOLD_GEN_JS = r"""
var cnt = String(arguments[0] || '1');
var c = document.getElementById('thrgh_ex_cnt');
if(c){ c.value=cnt;
  c.dispatchEvent(new Event('input',{bubbles:true}));
  c.dispatchEvent(new Event('change',{bubbles:true})); }
var _cf=window.confirm, _al=window.alert;
window.confirm=function(){return true;}; window.alert=function(){};
try{ createNewCarInfoHold(); }catch(e){ window.confirm=_cf; window.alert=_al; return 'ERR:'+e; }
window.confirm=_cf; window.alert=_al;
return 'ok';
"""


def fill_thrgh_hold(driver, row, log=None):
    """경유화물차 보유(미이행) 차량정보 행을 생성하고 채운다. (thrgh_ex_yn=Y 전용)"""
    hold = row.get("thrgh_hold") or {}
    if not hold:
        return
    cnt = str(row.get("thrgh_ex_cnt") or "1")
    try:
        r = driver.execute_script(_HOLD_GEN_JS, cnt)
    except Exception as e:
        r = "ERR:%s" % e
    if str(r).startswith("ERR"):
        if log:
            log(f"❌ 경유화물차 미이행 보유행 생성 실패({r}) → 보유대수·확인 직접", "red")
        return
    for _ in range(20):
        try:
            if driver.execute_script("return document.getElementsByName('ex_vh_num_hold').length>0;"):
                break
        except Exception:
            pass
        time.sleep(0.3)
    res = {}
    for _ in range(6):
        try:
            res = json.loads(driver.execute_script(_SCRAP_FILL_JS, hold))
        except Exception:
            res = {}
        if res.get("ex_vh_num_hold") not in (None, "NULL"):
            break
        time.sleep(0.5)
    nulls = [HOLD_LABELS.get(k, k) for k, v in res.items() if v == "NULL"]
    if log:
        if nulls:
            log(f"❌ 경유화물차 미이행 보유차량정보 일부 미입력: {', '.join(nulls)} → 직접 확인", "red")
        else:
            log(f"   ✓ 경유화물차 미이행 보유차량정보 {len(hold)}칸 입력 (보유대수 {cnt})", "green")


def verify_thrgh_hold(driver, row):
    """경유화물차 보유(미이행) 차량정보 되읽기 대조. 반환: 불일치 [(라벨, 기대, 실제)]."""
    hold = row.get("thrgh_hold") or {}
    if not hold:
        return []
    try:
        got = json.loads(driver.execute_script(_SCRAP_READ_JS, list(hold.keys())))
    except Exception:
        return [("경유화물차미이행보유정보", "(읽기실패)", "")]
    mm = []
    for k, want in hold.items():
        g = str(got.get(k, "") or "")
        if g.strip() != str(want).strip():
            mm.append((HOLD_LABELS.get(k, k), want, g or "(빈칸)"))
    return mm


# ── 첨부파일 업로드 ────────────────────────────────────────────
# 임시저장 후 뜨는 '지원신청 첨부파일' 페이지의 각 행은
#   <button onclick="popupAttachFile('<gubun>');return false;">첨부파일 등록</button>
# 을 가진다. popupAttachFile 은 이름이 'attachFile' 인 팝업을 열고 editForm 을 POST 한다.
# 첨부칸 개수는 신청 조건에 따라 다르다(화면에 '보이는' 칸만 필수):
#   일반   : A(신청서+동의서) / A2(구매계약서) / A3(등본·등기부) = 3곳
#   전환지원금: 위 3곳 + A7(우선순위증빙) A17(말소사실증명서)
#              A18(자동차등록원부) A19(가족관계증명서) = 7곳
# → 하드코딩 대신 visible_attach_gubuns() 로 '지금 보이는' 칸을 자동 감지해 전부 올린다.
ATTACH_GUBUNS = ["A", "A2", "A3"]   # 감지 실패 시 최소 기본값(폴백)

# A7(우선순위 증빙자료)는 '우선순위 대상'만 첨부, 순수 일반 신청만 제외.
#  우선순위 대상 = 사회계층(소상공인·차상위 등, social_yn=Y) · 생애최초 · 미세먼지 · 택시 · 전환지원금
#  - 일반(우선순위 아님) : 보이는 4칸(A·A2·A3·A7) 중 A7 빼고 3곳
#  - 우선순위(사회계층 등): 보이는 4칸 전부(A7=우선순위 증빙 포함)
#  - 전환지원금          : 보이는 7칸(A·A2·A3·A7·A17·A18·A19) 전부 업로드
#  전환지원금 여부는 전용 첨부(말소/등록원부/가족관계) 존재로, 나머지는 row 값(social_yn 등)으로 판별.
ATTACH_A7 = "A7"
EXCHANGE_ATTACH_MARKERS = {"A17", "A18", "A19"}
# 모든 케이스에서 자동 업로드 제외할 첨부칸(사람이 직접 판단해 올림)
#  A5 = 기타 증빙서류(출고배정표 등) → 자동화 대상에서 항상 제외
ATTACH_ALWAYS_SKIP = {"A5"}

# 첨부 화면에서 각 칸(gubun)이 어떤 서류인지 (로그·안내용)
ATTACH_DOC_NAMES = {
    "A": "신청서+개인정보동의서", "A2": "차량구매계약서", "A3": "주민등록등본/등기부",
    "A4": "지방세 납세증명서", "A5": "기타 증빙서류(출고배정표 등)",
    "A7": "우선순위 증빙자료(취약계층·다자녀·소상공인 등)",
    "A16": "농업인 확인서/농업경영체 등록확인서",
    "A17": "말소사실증명서", "A18": "자동차 등록원부", "A19": "가족관계증명서",
}

_VISIBLE_ATTACH_JS = r"""
var seen={}, out=[];
document.querySelectorAll('button,a,input[type=button]').forEach(function(b){
  var oc=(b.getAttribute('onclick')||'');
  var m=oc.match(/popupAttachFile\('([^']+)'\)/);
  if(!m) return;
  var g=m[1]; if(seen[g]) return;
  var r=b.getBoundingClientRect();
  if(r.width>0||r.height>0){ seen[g]=1; out.push(g); }
});
return out;
"""


def visible_attach_gubuns(driver):
    """지금 첨부 화면에서 '보이는' 첨부칸 gubun 목록(DOM 순서). 없으면 []"""
    try:
        gs = driver.execute_script(_VISIBLE_ATTACH_JS)
        return [str(g) for g in (gs or [])]
    except Exception:
        return []

# 팝업 내부의 업로드/저장 버튼 후보 탐색용 JS
_ATTACH_POPUP_BTN_JS = r"""
const out=[];
document.querySelectorAll('button,a,input[type=button],input[type=submit]').forEach(b=>{
  const t=(b.textContent||b.value||'').trim();
  const oc=(b.getAttribute('onclick')||'');
  if(/업로드|등록|저장|확인|적용|첨부|전송/.test(t) || /upload|save|submit|file|attach/i.test(oc)){
    out.push({t:t.slice(0,20), oc:oc.slice(0,140)});
  }
});
return out;
"""


def _accept_alerts(driver, log=None, max_iter=20):
    """업로드 직후 뜨는 alert('등록완료') 등을 자동 확인(accept).
    드라이버가 unhandledPromptBehavior='ignore' 라 알럿이 떠 있어도
    수동으로 switch_to.alert.accept() 는 동작한다.
    반환: 확인한 알럿 텍스트 리스트."""
    texts = []
    for _ in range(max_iter):
        try:
            al = driver.switch_to.alert
            t = (al.text or "").strip()
            al.accept()
            texts.append(t)
            if log:
                log(f"     알림 자동확인: '{t[:40]}'")
            time.sleep(0.4)
        except Exception:
            # 알럿 없음 / 수락 직후 팝업이 닫혀 창이 사라진 경우 등
            if texts:          # 이미 하나 이상 처리했으면 종료
                break
            time.sleep(0.3)
    return texts


def _find_attach_button(driver, gubun):
    """popupAttachFile('<gubun>') 를 정확히 호출하는 버튼(있으면) 반환.
    'A' 가 'A2' 에 매칭되지 않도록 세미콜론까지 포함해 비교."""
    needle = "popupAttachFile('%s');" % gubun
    els = driver.find_elements(By.XPATH, '//button[contains(@onclick, "%s")]' % needle)
    return els[0] if els else None


def _switch_to_attach_window(driver, gubun):
    """첨부 버튼이 있는 창으로 전환하고 그 핸들을 반환(없으면 None)."""
    for h in driver.window_handles:
        try:
            driver.switch_to.window(h)
        except Exception:
            continue
        if _find_attach_button(driver, gubun):
            return h
    return None


def upload_attachments(driver, pdf_path, gubuns=ATTACH_GUBUNS, submit=False, log=None):
    """pdf_path 한 파일을 gubuns 각 첨부칸에 올린다.
    submit=False: 파일만 선택해두고 팝업의 [업로드/저장]은 사람이 누름(기본, 안전).
    submit=True : 팝업의 업로드/저장 버튼까지 자동 클릭.
    반환: 처리한 칸 수(int)."""
    def _log(m, c=None):
        if log:
            log(m, c)

    if not pdf_path or not os.path.isfile(pdf_path):
        _log(f"⛔ 첨부 PDF 파일이 없습니다: {pdf_path}", "red")
        return 0

    base_handle = _switch_to_attach_window(driver, gubuns[0])
    if not base_handle:
        _log("⛔ 첨부 화면을 못 찾았습니다. 임시저장 후 '지원신청 첨부파일' 화면인지 확인하세요.", "red")
        return 0

    _log(f"📎 첨부 시작: {os.path.basename(pdf_path)} → {len(gubuns)}곳 {gubuns}")
    done = 0
    for g in gubuns:
        driver.switch_to.window(base_handle)
        # 직전 업로드로 부모(첨부 화면)가 새로고침 중일 수 있어 잠깐 재시도
        btn = None
        for _ in range(10):
            btn = _find_attach_button(driver, g)
            if btn:
                break
            time.sleep(0.3)
        if not btn:
            _log(f"  ⚠ '{g}' 첨부 버튼 없음 — 건너뜀")
            continue
        before = set(driver.window_handles)
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.2)
            btn.click()  # 네이티브 클릭 = 사용자 제스처(팝업차단 회피)
        except Exception as e:
            _log(f"  ⚠ '{g}' 버튼 클릭 실패: {e}")
            continue

        # 팝업 대기. 'attachFile' 이름은 재사용되므로 새 핸들이 안 생길 수도 있다
        popup = None
        for _ in range(20):
            time.sleep(0.3)
            new = set(driver.window_handles) - before
            if new:
                popup = new.pop()
                break
        if not popup:
            # 같은 이름 재사용 → 파일 input 이 있는 비(非)메인 창 탐색
            for h in driver.window_handles:
                if h == base_handle:
                    continue
                try:
                    driver.switch_to.window(h)
                    if driver.find_elements(By.CSS_SELECTOR, "input[type=file]"):
                        popup = h
                        break
                except Exception:
                    pass
        if not popup:
            _log(f"  ⚠ '{g}' 첨부 팝업이 안 떴습니다 — 건너뜀")
            continue

        driver.switch_to.window(popup)
        # 파일 input 등장 대기
        finp = None
        for _ in range(20):
            els = driver.find_elements(By.CSS_SELECTOR, "input[type=file]")
            if els:
                finp = els[0]
                break
            time.sleep(0.2)
        if not finp:
            _log(f"  ⚠ '{g}' 팝업에 파일 입력칸이 없습니다 — 건너뜀")
            driver.switch_to.window(base_handle)
            continue

        try:
            # 숨김 input 이어도 send_keys 는 동작. 혹시 막히면 보이게 처리 후 재시도
            try:
                finp.send_keys(pdf_path)
            except Exception:
                driver.execute_script(
                    "arguments[0].style.display='block';"
                    "arguments[0].style.visibility='visible';", finp)
                finp.send_keys(pdf_path)
            _log(f"  📎 '{g}' 파일 선택 완료")
        except Exception as e:
            _log(f"  ⚠ '{g}' 파일 지정 실패: {e}")
            driver.switch_to.window(base_handle)
            continue

        # 팝업의 업로드/저장 버튼 후보 진단(로그)
        try:
            cands = driver.execute_script(_ATTACH_POPUP_BTN_JS) or []
        except Exception:
            cands = []
        if cands:
            names = ", ".join(f"'{c['t']}'" for c in cands if c.get("t"))
            _log(f"     팝업 버튼 후보: {names or cands}")

        if submit:
            clicked = driver.execute_script(r"""
              const els=[...document.querySelectorAll('button,a,input[type=button],input[type=submit]')];
              const pri=el=>{const t=(el.textContent||el.value||'').trim();
                const oc=(el.getAttribute('onclick')||'');
                if(/업로드|전송/.test(t)) return 5;
                if(/저장|등록|확인|적용/.test(t)) return 4;
                if(/upload/i.test(oc)) return 3;
                if(/save|submit/i.test(oc)) return 2;
                if(/file|attach/i.test(oc)) return 1;
                return 0;};
              let best=null,bs=0;
              els.forEach(el=>{const s=pri(el); if(s>bs){bs=s;best=el;}});
              if(best){best.click(); return (best.textContent||best.value||'').trim().slice(0,20);}
              return null;
            """)
            if clicked:
                _log(f"  ⬆ '{g}' 업로드 버튼 자동 클릭: '{clicked}'", "green")
                # 업로드 후 뜨는 '등록완료' 알럿 자동 확인(사람이 안 눌러도 됨)
                _accept_alerts(driver, log=_log)
                time.sleep(0.8)
            else:
                _log(f"  ⚠ '{g}' 업로드 버튼을 못 찾음 → 팝업에서 직접 누르세요", "red")
        else:
            _log(f"  🖐 '{g}' 파일만 선택해 둠 — 팝업의 [업로드/저장]은 직접 누르세요")

        done += 1
        # 팝업이 아직 떠 있으면 정리하고 부모 화면으로 복귀
        try:
            if submit:
                for h in list(driver.window_handles):
                    if h == base_handle:
                        continue
                    try:
                        driver.switch_to.window(h)
                        if driver.find_elements(By.ID, "filename"):
                            driver.close()
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            driver.switch_to.window(base_handle)
        except Exception:
            base_handle = _switch_to_attach_window(driver, gubuns[0]) or base_handle

    _log(f"📎 첨부 처리 완료: {done}/{len(gubuns)}곳"
         + ("" if submit else " (각 팝업의 업로드/저장은 직접 확인)"))
    return done


# ── 임시저장(안전형) ───────────────────────────────────────────
# 라이브 H2 폼 확인 결과:
#   [임시저장] 버튼  : text="임시저장", onclick="goSave();return false;"
#   goSave() → goSave_1('N') → (유효성검사) → confirm("임시저장을 하시겠습니까?")
#            → OK → checkFinish() → 보안 확인코드 팝업창(popupSellerRandomChk)
# 안전 원칙: '임시저장' 텍스트 + onclick goSave() 인 버튼만, 화면에 보이고 유일할 때만 클릭.
#            '신청/제출/지급/우선순위/저장(단독)' 등은 절대 누르지 않는다(오발송 방지).
_SAVE_AVOID = ("신청", "제출", "최종", "지급", "우선순위", "취소", "삭제", "보완", "등록")


def _find_temp_save_buttons(driver):
    """화면에 보이는 '임시저장' 버튼 후보 [(el, text)]. goSave() onclick 인 것만."""
    cands = []
    try:
        els = driver.find_elements(
            By.XPATH,
            "//button|//a|//input[@type='button']|//input[@type='submit']")
    except Exception:
        return cands
    for el in els:
        try:
            t = (el.text or el.get_attribute("value") or "").strip()
        except Exception:
            continue
        if "임시저장" not in t or len(t) >= 20:
            continue
        if any(a in t for a in _SAVE_AVOID):
            continue
        oc = (el.get_attribute("onclick") or "")
        if "goSave(" not in oc:        # goSavePay/goPriSave 등 배제
            continue
        try:
            if el.is_displayed():
                cands.append((el, t))
        except Exception:
            pass
    return cands


# confirm/alert 를 가로채 둔다(한 번 설치하면 페이지가 살아있는 동안 계속 유지).
#  - confirm("임시저장을 하시겠습니까?") → 네이티브 대화상자 없이 자동 OK(true) 반환
#  - alert(...)  → 화면 대화상자 대신 배열(window.__cm_msgs)로 수집
# 왜 'persistent' 인가: "임시저장 완료" 알림은 보안코드 처리·저장이 다 끝난 한참 뒤에
# 뜬다. 잠깐만 가로채고 풀면 그 완료 알림이 네이티브 창으로 떠 멈춰버린다(드라이버까지 블록).
# 그래서 풀지 않고 유지하고, 워커 루프가 __cm_msgs 를 비우며 로그로 남긴다.
# checkFinish() 가 여는 보안코드 '창'(window.open)은 영향 없음 → 워커가 따로 처리.
_ALERT_HOOK_JS = r"""
if(!window.__cm_hooked){
  window.__cm_hooked = true;
  window.__cm_msgs = [];
  window.confirm = function(m){ window.__cm_msgs.push('[confirm]'+String(m)); return true; };
  window.alert   = function(m){ window.__cm_msgs.push('[alert]'+String(m)); };
}
"""
_ALERT_DRAIN_JS = r"""
if(!window.__cm_msgs){ return []; }
var m = window.__cm_msgs; window.__cm_msgs = []; return m;
"""
# alert 메시지가 '저장 진행 안 됨'을 뜻하는지 판단할 키워드(검증/차단 메시지)
_SAVE_BLOCK_HINTS = ("입력", "선택", "확인해", "0원", "중복", "최대", "권한", "없습니다", "다시")


def click_temp_save(driver, log=None):
    """'임시저장' 버튼을 안전하게 클릭. confirm("임시저장을 하시겠습니까?")은 자동 OK.
    보안 확인코드 팝업창은 건드리지 않는다(워커의 handle_random_popup 가 처리).
    반환: 'ok'(저장 진행) | 'blocked'(검증 등으로 저장 안 됨) | 'notfound' | 'ambiguous' | 'error'."""
    def _log(m, c=None):
        if log:
            log(m, c)

    cands = _find_temp_save_buttons(driver)
    if not cands:
        _log("⛔ '임시저장' 버튼을 못 찾았습니다 → 직접 누르세요", "red")
        return "notfound"
    if len(cands) > 1:
        names = ", ".join(f"'{t}'" for _, t in cands)
        _log(f"⛔ '임시저장' 후보가 여러 개({names}) → 안전상 자동 클릭 보류, 직접 누르세요", "red")
        return "ambiguous"

    btn, label = cands[0]
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        driver.execute_script(_ALERT_HOOK_JS)  # confirm/alert 가로채기(유지 — 풀지 않음)
        time.sleep(0.2)
        btn.click()                            # 네이티브 클릭. confirm 자동 OK 라 안 멈춤.
        _log(f"💾 '{label}' 클릭")
    except Exception as e:
        _log(f"⛔ '임시저장' 클릭 실패: {e}", "red")
        return "error"

    # checkFinish() 의 비동기 검사(alert(RETURNMSG)) 까지 잠깐 대기 후 메시지 수집.
    # 'hook' 은 풀지 않는다 → 한참 뒤의 "임시저장 완료" 알림은 워커 루프가 비워서 처리.
    time.sleep(1.2)
    try:
        msgs = driver.execute_script(_ALERT_DRAIN_JS) or []
    except Exception:
        msgs = []   # 페이지가 이미 넘어갔으면 메시지 못 읽음

    has_confirm = any(str(m).startswith("[confirm]") for m in msgs)
    alerts = [str(m)[len("[alert]"):] for m in msgs if str(m).startswith("[alert]")]
    block_alerts = [a for a in alerts if any(k in a for k in _SAVE_BLOCK_HINTS)]

    if block_alerts:
        joined = " / ".join(block_alerts)
        _log(f"⛔ 임시저장 차단(검증 알럿): '{joined}' → 저장 안 됨, 폼 확인 필요", "red")
        return "blocked"
    if has_confirm:
        _log("   확인창 자동 OK — 보안코드 팝업/첨부화면으로 진행", "green")
        return "ok"
    if alerts:
        _log(f"   알림: '{' / '.join(alerts)}'")
        return "ok"
    # confirm 도 alert 도 없었음: 페이지가 이미 넘어갔거나(저장 진행) 흐름이 다름
    _log("   확인창 없이 진행됨(저장 단계로 넘어갔을 수 있음) — 보안팝업/첨부화면 확인")
    return "ok"


# ── 워커 ──────────────────────────────────────────────────────
class Worker(threading.Thread):
    def __init__(self, rows, log_q, state_q, cmd_q, auto_flag, addr_flag,
                 code_flag, code_confirm_flag, attach_pattern, attach_submit_flag,
                 autosave_flag, autoattach_flag):
        super().__init__(daemon=True)
        self.rows = rows
        self.log_q = log_q
        self.state_q = state_q
        self.cmd_q = cmd_q
        self.auto_flag = auto_flag
        self.addr_flag = addr_flag
        self.code_flag = code_flag            # 확인코드 자동입력 on/off
        self.code_confirm_flag = code_confirm_flag  # 확인(저장)까지 자동 on/off
        self.attach_pattern = attach_pattern  # [str] 첨부 PDF 경로 패턴({name} 치환)
        self.attach_submit_flag = attach_submit_flag  # [bool] 팝업 업로드/저장까지 자동
        self.autosave_flag = autosave_flag    # [bool] 검증 통과 시 임시저장 자동 클릭
        self.autoattach_flag = autoattach_flag  # [bool] 첨부화면 감지 시 자동 업로드
        self.attach_handled = False           # 첨부 자동 1회 처리 디바운스
        self.idx = 0
        self.running = True
        self.driver = None
        self.blank_handled = False
        self.main_handle = None
        self._known_handles = set()    # 이미 본 창들(새 창만 검사 → 포커스 도둑질 방지)
        self._popup_focus = None       # 사용자가 [확인] 누를 보안팝업(포커스 유지)
        self._save_armed = False       # 임시저장을 한 번이라도 쓰면 알림 가로채기 유지

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

        # 3-2) 전환지원금 폐차 정보 (exchange_3year_yn=Y 이고 데이터 있을 때)
        #   fill_one 에서 exchange_3year_yn 라디오는 이미 Y 로 눌림 → 폐차대수/행 생성/입력
        if str(row.get("exchange_3year_yn", "") or "").strip() == "Y" and row.get("exchange_scrap"):
            self.log("전환지원금(폐차 후 전기차) — 폐차 정보 입력…")
            fill_exchange_scrap(self.driver, row, log=self.log)

        # 3-3) 경유화물차 보유 미이행 정보 (thrgh_ex_yn=Y 이고 보유차량 데이터 있을 때)
        #   fill_one 에서 thrgh_ex_yn 라디오는 이미 Y 로 눌림 → 보유대수/행 생성/입력
        if str(row.get("thrgh_ex_yn", "") or "").strip() == "Y" and row.get("thrgh_hold"):
            self.log("경유화물차 보유 미이행 — 보유차량 정보 입력…")
            fill_thrgh_hold(self.driver, row, log=self.log)

        # 4) 되읽기 대조 (일반 필드 + 전환지원금 폐차 + 경유화물차 보유 미이행)
        time.sleep(0.3)
        mm = verify_fill(self.driver, row)
        mm += verify_exchange_scrap(self.driver, row)
        mm += verify_thrgh_hold(self.driver, row)
        self.idx += 1
        # 차종(model_cd)은 '가장 비슷한 것'으로 들어가 거의 불일치 → 임시저장 차단에서 제외.
        # (사람이 빨간 경고로 직접 확인하기로 한 항목. 단, 화면엔 그대로 표시함)
        blocking = [m for m in mm if m[0] != "model_cd"]
        model_mm = [m for m in mm if m[0] == "model_cd"]
        if mm:
            self.log(f"❌ {self.idx}번째 검증 — 불일치 {len(mm)}건:")
            for fid, want, got in mm:
                tag = "  ← 차종(자동저장 차단 제외, 직접 확인!)" if fid == "model_cd" else ""
                self.log(f"    · {fid}: 기대 '{want}' / 실제 '{got}'{tag}",
                         color=("red" if fid == "model_cd" else None))
        if blocking:
            self.set_state(status=f"❌ 검증 실패 {len(blocking)}건 — 저장 말고 직접 확인!",
                           progress=f"{self.idx}/{len(self.rows)}")
            return
        # 여기부터: 차종 외 모든 항목 일치(통과). 차종만 불일치면 경고 후 진행.
        if model_mm:
            self.log("⚠ 차종 외 항목은 모두 일치 — 차종은 직접 확인하기로 하고 진행", color="red")
            self.set_state(status="⚠ 차종 외 통과 — (차종 직접 확인)",
                           progress=f"{self.idx}/{len(self.rows)}")
        else:
            self.log(f"✅ {self.idx}번째 입력·검증 통과. 검토 후 저장/제출은 직접.")
            self.set_state(status="✅ 검증 통과 — 검토 후 다음 폼으로",
                           progress=f"{self.idx}/{len(self.rows)}")
        # 임시저장 자동(옵션): 차종 외 불일치가 없을 때만. (차종은 위에서 경고함)
        if self.autosave_flag[0]:
            time.sleep(0.4)
            self.log("자동 임시저장 진행(차종 외 검증 통과 건)…")
            self.do_temp_save()

    def run(self):
        # 연결 검증
        try:
            opts = webdriver.ChromeOptions()
            opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            # 확인창(confirm/alert)을 드라이버가 자동으로 닫지 않게 함.
            # (폴링 중 사용자의 "이동하시겠습니까?" 창이 자동 취소되던 문제 방지)
            opts.set_capability("unhandledPromptBehavior", "ignore")
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
            # 가로챈 알림(confirm OK / "임시저장 완료" 등) 비우며 로그 — 네이티브 창 방지
            self._pump_alerts()

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

            # 보안팝업 포커스가 없을 때만: 저장 완료 등 네이티브 alert 자동 수락
            self._accept_native_alert()

            try:
                while True:
                    cmd = self.cmd_q.get_nowait()
                    if isinstance(cmd, tuple):
                        if cmd and cmd[0] == "enter":
                            self.do_enter(*cmd[1:])
                        continue
                    if cmd == "fill":
                        self.do_fill("수동 입력")
                    elif cmd == "reverify":
                        self.reverify()
                    elif cmd == "attach":
                        self.do_attach()
                    elif cmd == "tempsave":
                        self.do_temp_save()
            except queue.Empty:
                pass

            local = get_value(self.driver, "local_nm")
            req_nm = get_value(self.driver, "req_nm")
            self.set_state(local=local or "(미선택)")

            # 첨부 화면 감지(지자체 감지처럼): 감지되면 상태표시 + 옵션 시 자동 업로드
            on_attach = self.on_attach_screen()
            if not on_attach:
                self.attach_handled = False
            else:
                if not self.attach_handled:
                    self.set_state(status="📎 첨부화면 감지됨"
                                   + (" — 자동 업로드 진행" if self.autoattach_flag[0]
                                      else " — [📎 첨부] 또는 자동옵션"))
                if self.autoattach_flag[0] and not self.attach_handled:
                    self.attach_handled = True   # 1회만(중복 업로드 방지)
                    self.log("첨부화면 자동 감지 → 첨부 업로드 진행")
                    self.do_attach()

            blank_form = self.on_form() and local != "" and req_nm == "" and not on_attach
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

    def do_attach(self):
        """임시저장 후 첨부 화면에서 현재(직전 입력) 신청자의 PDF를
        '지금 보이는' 첨부칸에 올린다. A7(우선순위 증빙)은 전환지원금·택시 신청자만 포함."""
        # 직전 입력 건의 신청자 row / 이름
        ridx = self.idx - 1 if self.idx > 0 else 0
        row = self.rows[ridx] if 0 <= ridx < len(self.rows) else {}
        name = str(row.get("req_nm", "") or "").strip()
        pattern = (self.attach_pattern[0] or "").strip()
        if not pattern:
            self.log("⛔ 첨부 PDF 경로(패턴)가 비어 있습니다.", color="red")
            self.set_state(status="⛔ 첨부 경로 없음")
            return
        pdf_path = pattern.replace("{name}", name)
        self.log(f"📎 첨부 대상자: '{name or '(미상)'}' / 파일: {pdf_path}")
        if not os.path.isfile(pdf_path):
            self.log(f"⛔ 파일을 못 찾음: {pdf_path} — 파일명/경로를 확인하세요.", color="red")
            self.set_state(status="⛔ 첨부 파일 없음")
            return
        # 화면에 보이는 첨부칸 자동 감지 (전환지원금이면 7곳으로 늘어남)
        detected = visible_attach_gubuns(self.driver) or list(ATTACH_GUBUNS)
        # A7(우선순위 증빙자료)은 '우선순위 대상'만 첨부.
        #  폼 안내상 우선순위 항목 = 사회계층 여부 · 미세먼지 개선효과 · 생애최초 (+ 택시 · 전환지원금)
        #  → 소상공인 등 사회계층(social_yn=Y)도 우선순위 대상이므로 A7 포함.
        is_exchange = any(g in detected for g in EXCHANGE_ATTACH_MARKERS)
        is_taxi = str(row.get("taxi_yn", "") or "").strip() == "Y"
        is_social = str(row.get("social_yn", "") or "").strip() == "Y"
        is_first_buy = str(row.get("first_buy_yn", "") or "").strip() == "Y"
        is_improve_fd = str(row.get("improve_fd_yn", "") or "").strip() == "Y"
        is_priority = is_social or is_first_buy or is_improve_fd
        include_a7 = is_exchange or is_taxi or is_priority
        if include_a7:
            gubuns = list(detected)          # A7 포함 전부
            skipped = []
        else:
            gubuns = [g for g in detected if g != ATTACH_A7]   # 일반: A7 제외
            skipped = [g for g in detected if g == ATTACH_A7]
        # 모든 케이스에서 제외할 칸(A5 기타 증빙 등) — 자동 업로드 대상에서 항상 뺀다
        always_skip = [g for g in gubuns if g in ATTACH_ALWAYS_SKIP]
        if always_skip:
            gubuns = [g for g in gubuns if g not in ATTACH_ALWAYS_SKIP]
            skipped = skipped + always_skip
        # 판정 사유(로그용)
        reasons = []
        if is_exchange:
            reasons.append("전환지원금")
        if is_taxi:
            reasons.append("택시")
        if is_social:
            social_txt = str(row.get("social_kind", "") or "").strip() or "사회계층"
            reasons.append(social_txt)
        if is_first_buy:
            reasons.append("생애최초")
        if is_improve_fd:
            reasons.append("미세먼지")
        kind = "·".join(reasons) if reasons else "일반"
        docs = ", ".join("%s(%s)" % (g, ATTACH_DOC_NAMES.get(g, "?")) for g in gubuns)
        self.log(f"📎 [{kind}] 감지 {len(detected)}곳 → {len(gubuns)}곳 업로드: {docs}")
        if skipped:
            sk = ", ".join("%s(%s)" % (g, ATTACH_DOC_NAMES.get(g, "?")) for g in skipped)
            self.log(f"   ⏭ 자동 제외: {sk} "
                     f"(A7=우선순위 신청자만 · A5=기타는 항상 제외) — 필요시 직접 업로드")
        submit = bool(self.attach_submit_flag[0])
        try:
            n = upload_attachments(self.driver, pdf_path, gubuns=gubuns,
                                   submit=submit, log=self.log)
        except Exception as e:
            self.log(f"첨부 오류: {e}", color="red")
            self.set_state(status="첨부 오류 (로그 확인)")
            return
        finally:
            # 작업창으로 복귀 시도
            try:
                if self.main_handle and self.main_handle in self.driver.window_handles:
                    self.driver.switch_to.window(self.main_handle)
            except Exception:
                pass
        if n:
            tail = "업로드까지 자동" if submit else "파일선택 — 업로드/저장은 직접"
            self.set_state(status=f"📎 첨부 {n}곳 ({tail})")
        else:
            self.set_state(status="📎 첨부 처리 0곳 — 로그 확인")

    def _accept_native_alert(self):
        """JS 가로채기로 못 막는 네이티브 alert 를 드라이버로 직접 수락.
        대표 사례: 저장 완료 후 randomsave(iframe)에서 뜨는 '임시저장 완료' alert.
        임시저장을 한 번이라도 쓴 뒤(_save_armed)에만, 보안팝업 포커스가 없을 때만 동작."""
        if not self._save_armed:
            return
        # 현재 창이 죽었으면(닫힌 팝업 등) 메인 폼으로 복귀 후 검사
        try:
            self.driver.current_window_handle
        except Exception:
            try:
                if self.main_handle and self.main_handle in self.driver.window_handles:
                    self.driver.switch_to.window(self.main_handle)
            except Exception:
                return
        try:
            al = self.driver.switch_to.alert
            t = (al.text or "").strip()
            al.accept()
        except Exception:
            return  # 떠 있는 alert 없음
        if "완료" in t:
            self.log(f"   ✅ 완료 알림 자동확인: '{t}'", color="green")
            self.set_state(status="✅ 임시저장 완료 — 첨부화면 확인")
        elif any(k in t for k in _SAVE_BLOCK_HINTS):
            self.log(f"   ⛔ 알림(저장 안 됨 가능): '{t}'", color="red")
        else:
            self.log(f"   알림 자동확인: '{t}'")

    def _pump_alerts(self):
        """가로챈 confirm/alert 메시지(window.__cm_msgs)를 비우며 로그로 남긴다.
        현재 창 기준으로만 동작(창 전환 안 함 → 포커스 도둑질 없음).
        hook 이 안 걸린 창이면 그냥 빈 결과."""
        try:
            if self._save_armed:
                # 저장 후 페이지가 다시 그려져도 가로채기가 살아있도록 재설치(멱등)
                self.driver.execute_script(_ALERT_HOOK_JS)
            msgs = self.driver.execute_script(_ALERT_DRAIN_JS) or []
        except Exception:
            return
        for m in msgs:
            s = str(m)
            if s.startswith("[confirm]"):
                self.log(f"   확인창 자동 OK: '{s[len('[confirm]'):]}'", color="green")
                continue
            txt = s[len("[alert]"):] if s.startswith("[alert]") else s
            if "완료" in txt:
                self.log(f"   ✅ 완료 알림 자동확인: '{txt}'", color="green")
            elif any(k in txt for k in _SAVE_BLOCK_HINTS):
                self.log(f"   ⛔ 알림(저장 안 됨 가능): '{txt}'", color="red")
            else:
                self.log(f"   알림 자동확인: '{txt}'")

    def on_attach_screen(self):
        """'지원신청 첨부파일' 화면인지 감지.
        신호: 같은 sellerApplyform URL 인데 ① popupAttachFile('A') 버튼이 보이고
              ② 임시저장 버튼이 사라짐(입력 폼 단계가 끝남)."""
        try:
            if not self.on_form():
                return False
            b = _find_attach_button(self.driver, ATTACH_GUBUNS[0])
            if not (b and b.is_displayed()):
                return False
            if _find_temp_save_buttons(self.driver):
                return False   # 임시저장 버튼이 아직 있으면 입력 폼 단계
            return True
        except Exception:
            return False

    def do_enter(self, car_kind, car_type, local_name):
        """지자체 선택 → 신청서 작성 폼 진입(과부하 시 재시도). 저장/제출은 안 함.
        browse_control.enter_form 에 위임하고, 팝업 처리 후 창 핸들/포커스를 재정렬한다."""
        self.set_state(status="🚪 폼 진입 중… (과부하면 자동 재시도)")
        try:
            ok = browse_control.enter_form(
                self.driver, car_type=car_type, car_kind=car_kind,
                local_name=(local_name or None), log=self.log)
        except Exception as e:
            self.log(f"폼 진입 오류: {e}", color="red")
            self.set_state(status="폼 진입 오류 (로그 확인)")
            return
        finally:
            # 팝업 열고 닫은 뒤: 메인(폼) 창 재확정 + 알려진 창 목록 갱신
            try:
                if self.on_form():
                    self.main_handle = self.driver.current_window_handle
                self._known_handles = set(self.driver.window_handles)
            except Exception:
                pass
        if ok:
            loc = get_value(self.driver, "local_nm")
            self.log(f"✅ 폼 진입 완료 — 지자체 '{loc or '(미표시)'}'. [▶ 다음 건 입력] 또는 자동입력 대기.")
            self.set_state(status=f"✅ 폼 진입 완료 (지자체 {loc or '?'})",
                           local=loc or "(미선택)")
        else:
            self.set_state(status="⚠ 폼 진입 미완료 — 로그 확인 후 재시도/직접 진입")

    def do_temp_save(self):
        """'임시저장' 버튼을 안전하게 클릭(confirm 자동 수락). 보안코드 팝업은 기존 루프가 처리."""
        if not self.on_form():
            self.log("⚠ 임시저장: 신청서 작성 폼이 아닙니다 → 보류", color="red")
            self.set_state(status="⚠ 임시저장 보류 — 폼 아님")
            return
        self._save_armed = True   # 이후 루프가 알림 가로채기를 유지
        try:
            st = click_temp_save(self.driver, log=self.log)
        except Exception as e:
            self.log(f"임시저장 오류: {e}", color="red")
            self.set_state(status="임시저장 오류 (로그 확인)")
            return
        finally:
            try:
                if self.main_handle and self.main_handle in self.driver.window_handles:
                    self.driver.switch_to.window(self.main_handle)
            except Exception:
                pass
        if st == "ok":
            self.set_state(status="💾 임시저장 클릭 — 보안코드/첨부화면 확인")
        else:
            self.set_state(status="💾 임시저장 보류 — 직접 누르세요")

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
        self.geometry("560x720")
        self.worker = None
        self.rows = []
        self.valid = False
        self.log_q = queue.Queue()
        self.state_q = queue.Queue()
        self.cmd_q = queue.Queue()
        self.auto_flag = [True]
        self.addr_flag = [True]
        self.code_flag = [True]           # 확인코드 자동입력 — 기본 켬
        self.code_confirm_flag = [True]   # 확인(저장)까지 자동 — 기본 켬(완전 자동)
        self.autosave_flag = [True]       # 검증 통과 시 임시저장 자동 — 기본 켬
        self.autoattach_flag = [True]     # 첨부화면 감지 시 자동 업로드 — 기본 켬
        _desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        # 첨부 PDF 경로 패턴: {name} 은 신청자명으로 치환됨
        self.attach_pattern = [os.path.join(_desktop, "세종 수소 {name}.pdf")]
        self.attach_submit_flag = [True]   # 팝업 [업로드/저장]까지 자동 — 기본 켬

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

        # ── 지자체 선택 → 신청서 폼 진입(과부하 재시도) ──
        # 표시명 → (car_kind, car_type). 전기 11 승용 / 12 화물 / 13 승합, 수소는 단일 폼.
        self.CARTYPES = {
            "전기 승용": ("전기", "11"), "전기 화물": ("전기", "12"),
            "전기 승합": ("전기", "13"), "수소 승용": ("수소", "11"),
        }
        bar_enter = ttk.Frame(self)
        bar_enter.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Label(bar_enter, text="차종").pack(side="left")
        self.cartype_var = tk.StringVar(value="전기 승용")
        ttk.Combobox(bar_enter, textvariable=self.cartype_var, width=8, state="readonly",
                     values=list(self.CARTYPES)).pack(side="left", padx=(2, 8))
        ttk.Label(bar_enter, text="지자체").pack(side="left")
        self.local_var = tk.StringVar(value="")
        ttk.Entry(bar_enter, textvariable=self.local_var, width=12).pack(side="left", padx=(2, 8))
        self.enter_btn = ttk.Button(bar_enter, text="🚪 지자체 선택→폼 진입",
                                    command=self.manual_enter, state="disabled")
        self.enter_btn.pack(side="left")

        bar3 = ttk.Frame(self)
        bar3.pack(fill="x", padx=10, pady=(0, 4))
        self.code_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar3, text="확인코드(보안문자) 역순 자동입력", variable=self.code_var,
                        command=self.toggle_code).pack(side="left")
        self.code_confirm_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar3, text="확인까지 자동(=저장 실행)", variable=self.code_confirm_var,
                        command=self.toggle_code).pack(side="left", padx=10)

        # ── 임시저장 ──
        bar4 = ttk.Frame(self)
        bar4.pack(fill="x", padx=10, pady=(0, 4))
        self.tempsave_btn = ttk.Button(bar4, text="💾 임시저장", command=self.manual_tempsave,
                                       state="disabled")
        self.tempsave_btn.pack(side="left")
        self.autosave_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar4, text="검증 통과 시 임시저장 자동", variable=self.autosave_var,
                        command=self.toggle_autosave).pack(side="left", padx=8)
        ttk.Label(bar4, text="('임시저장' 버튼+확인창만 자동 · 보안코드 [확인]은 위 옵션)",
                  font=("", 8)).pack(side="left")

        # ── 첨부파일(임시저장 후 화면) ──
        att = ttk.LabelFrame(self, text="첨부파일 (임시저장 후 화면에서)")
        att.pack(fill="x", padx=10, pady=4)
        arow = ttk.Frame(att)
        arow.pack(fill="x", padx=6, pady=(4, 2))
        ttk.Label(arow, text="PDF 경로", width=8).pack(side="left")
        self.attach_var = tk.StringVar(value=self.attach_pattern[0])
        # tk.Entry(클래식) — 테두리 색을 줄 수 있음(ttk.Entry는 macOS에서 테두리색 안 먹힘)
        self.attach_entry = tk.Entry(arow, textvariable=self.attach_var,
                                     highlightthickness=2)
        self.attach_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.attach_entry.bind("<KeyRelease>", lambda e: self._refresh_pdf_border())
        ttk.Button(arow, text="찾기", width=6,
                   command=self.pick_attach).pack(side="left")
        arow2 = ttk.Frame(att)
        arow2.pack(fill="x", padx=6, pady=(0, 4))
        ttk.Label(arow2, text="(파일명의 신청자 자리에 {name} 사용 가능 · 화면에 보이는 첨부칸 자동감지: 일반 3곳/전환지원금 7곳)",
                  font=("", 8)).pack(side="left")
        arow3 = ttk.Frame(att)
        arow3.pack(fill="x", padx=6, pady=(0, 5))
        self.attach_btn = ttk.Button(arow3, text="📎 첨부 올리기(보이는 칸 전부)", command=self.manual_attach,
                                     state="disabled")
        self.attach_btn.pack(side="left")
        self.attach_submit_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(arow3, text="팝업 [업로드/저장]까지 자동", variable=self.attach_submit_var,
                        command=self.toggle_attach).pack(side="left", padx=8)
        self.autoattach_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(arow3, text="첨부화면 감지 시 자동 업로드", variable=self.autoattach_var,
                        command=self.toggle_autoattach).pack(side="left", padx=8)

        info = ttk.LabelFrame(self, text="현재 상태")
        info.pack(fill="x", padx=10, pady=4)
        # 워커 상태등: 작동중=초록, 정지=빨강
        wrow = ttk.Frame(info)
        wrow.pack(fill="x", padx=6, pady=2)
        ttk.Label(wrow, text="워커", width=12).pack(side="left")
        self.worker_dot = tk.Label(wrow, text="● 정지 (시작 안 함)",
                                   fg="#cc0000", font=("", 10, "bold"))
        self.worker_dot.pack(side="left")
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

        self._refresh_pdf_border()
        self._refresh_worker_dot()
        self.after(200, self.poll_queues)

    def _refresh_pdf_border(self):
        """PDF 경로 미설정/없는 파일이면 테두리 빨강, 정상이면 초록."""
        path = self.attach_var.get().strip()
        # {name} 템플릿이면 신청자명 들어오기 전이라 파일존재 확인 불가 → 경로만 있으면 OK로 본다
        if not path:
            ok = False
        elif "{name}" in path:
            ok = True
        else:
            ok = os.path.isfile(path)
        color = "#117711" if ok else "#cc0000"
        try:
            self.attach_entry.config(highlightbackground=color, highlightcolor=color)
        except Exception:
            pass

    def _refresh_worker_dot(self):
        """워커 작동 여부에 따라 상태등 색/문구 갱신."""
        alive = bool(self.worker and self.worker.is_alive())
        if alive:
            self.worker_dot.config(text="● 작동중", fg="#117711")
        else:
            self.worker_dot.config(text="● 정지 (시작 안 함)", fg="#cc0000")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
        if not path:
            return
        if detect_format(path) == "vertical":
            self.load_vertical(path)
            return
        try:
            import pandas as pd  # 표형식 전용(선택적 의존성)
        except ImportError:
            self._log("❌ 표형식(CSV/일반 xlsx)은 pandas가 필요합니다. "
                      "세로형 양식(신청서_이름.xlsx)을 사용하세요.")
            return
        try:
            df = (pd.read_csv(path, dtype=str) if path.endswith(".csv")
                  else pd.read_excel(path, dtype=str)).fillna("")
        except Exception as e:
            self._log(f"❌ 엑셀 읽기 실패: {e}")
            return
        rows = df.to_dict("records")
        cols = list(df.columns)

        # 표형식에서도 차종·사회계층은 한글 텍스트매칭(_text_select)으로 처리
        # (코드 없이 "디올뉴넥쏘" 등 한글명 → 페이지 옵션 대조 + 빨간경고 유지),
        # '구분'(car_kind) 열이 없으면 차종명으로 수소차 추정(전기/수소 안전가드 유지)
        for row in rows:
            ts = set()
            for fid in ("model_cd", "social_kind"):
                if str(row.get(fid, "") or "").strip():
                    ts.add(fid)
            row["_text_select"] = ts
            if not str(row.get("car_kind", "") or "").strip():
                m = re.sub(r"\s+", "", str(row.get("model_cd", "") or ""))
                if "넥쏘" in m or "수소" in m:
                    row["car_kind"] = "수소"

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
            if rows:
                self._prefill_enter_from_row(rows[0])
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
        self._prefill_enter_from_row(row)
        self.start_btn.config(state="normal")
        self.v_status.set("세로형 1건 로드 — 시작 가능 (위 handoff 직접 확인)")
        self.v_prog.set("0/1")
        self._log("✅ 로드 완료. [시작] 후 폼에서 자동 입력됩니다.")

    def _prefill_enter_from_row(self, row):
        """로드된 건의 지자체/전기·수소를 '폼 진입' 입력칸에 미리 채운다.
        (승용/화물/승합 세부는 엑셀로 알 수 없어 승용 기본 — 콤보박스로 조정)."""
        try:
            loc = str(row.get("local_nm", "") or "").strip()
            if loc:
                self.local_var.set(loc)
            kind = str(row.get("car_kind", "") or "").strip()
            if kind == "수소":
                self.cartype_var.set("수소 승용")
            elif kind == "전기":
                self.cartype_var.set("전기 승용")
        except Exception:
            pass

    def start(self):
        if not (self.rows and self.valid):
            self._log("시작 불가: 검증 통과한 엑셀이 없습니다.")
            return
        if self.worker and self.worker.is_alive():
            return
        self.attach_pattern[0] = self.attach_var.get().strip()
        self.worker = Worker(self.rows, self.log_q, self.state_q,
                             self.cmd_q, self.auto_flag, self.addr_flag,
                             self.code_flag, self.code_confirm_flag,
                             self.attach_pattern, self.attach_submit_flag,
                             self.autosave_flag, self.autoattach_flag)
        self.worker.start()
        self.fill_btn.config(state="normal")
        self.reverify_btn.config(state="normal")
        self.attach_btn.config(state="normal")
        self.tempsave_btn.config(state="normal")
        self.enter_btn.config(state="normal")
        self._log("워커 시작")

    def stop(self):
        if self.worker:
            self.worker.stop()
            self.fill_btn.config(state="disabled")
            self.reverify_btn.config(state="disabled")
            self.attach_btn.config(state="disabled")
            self.tempsave_btn.config(state="disabled")
            self.enter_btn.config(state="disabled")
            self._log("정지 요청")

    def manual_fill(self):
        if self.worker and self.worker.is_alive():
            self.cmd_q.put("fill")

    def manual_enter(self):
        """지자체 선택 → 신청서 폼 진입 요청. 지자체명이 비면 폼 진입만."""
        if not (self.worker and self.worker.is_alive()):
            self._log("폼 진입하려면 먼저 [시작] 하세요.")
            return
        kind, ctype = self.CARTYPES.get(self.cartype_var.get(), ("전기", "11"))
        local = self.local_var.get().strip()
        self.cmd_q.put(("enter", kind, ctype, local))
        self._log(f"🚪 폼 진입 요청: {self.cartype_var.get()}"
                  + (f" · 지자체 '{local}'" if local else " · (지자체 미지정 → 폼 진입만)"))

    def manual_reverify(self):
        if self.worker and self.worker.is_alive():
            self.cmd_q.put("reverify")

    def manual_attach(self):
        if self.worker and self.worker.is_alive():
            self.attach_pattern[0] = self.attach_var.get().strip()
            self.cmd_q.put("attach")
        else:
            self._log("첨부하려면 먼저 [시작] 하세요.")

    def manual_tempsave(self):
        if self.worker and self.worker.is_alive():
            self.cmd_q.put("tempsave")
        else:
            self._log("임시저장하려면 먼저 [시작] 하세요.")

    def toggle_autosave(self):
        self.autosave_flag[0] = self.autosave_var.get()
        self._log("검증 통과 시 임시저장 자동 "
                  + ("켜짐 (검증 통과 건만 임시저장 클릭)" if self.autosave_flag[0] else "꺼짐"))

    def toggle_autoattach(self):
        self.autoattach_flag[0] = self.autoattach_var.get()
        self._log("첨부화면 감지 시 자동 업로드 "
                  + ("켜짐 (첨부화면 뜨면 A/A2/A3 자동)" if self.autoattach_flag[0] else "꺼짐"))

    def pick_attach(self):
        path = filedialog.askopenfilename(
            title="첨부할 PDF 선택", filetypes=[("PDF", "*.pdf"), ("모든 파일", "*.*")])
        if path:
            self.attach_var.set(path)
            self.attach_pattern[0] = path
            self._refresh_pdf_border()

    def toggle_attach(self):
        self.attach_submit_flag[0] = self.attach_submit_var.get()
        self._log("팝업 업로드/저장까지 자동 "
                  + ("켜짐 (자동 업로드 실행됨)" if self.attach_submit_flag[0] else "꺼짐"))

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
        self._refresh_worker_dot()
        self._refresh_pdf_border()
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
