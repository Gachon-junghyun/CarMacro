"""
browse_control.py — ev.or.kr 신청서 폼 '진입' 자동화 (지자체 선택 → 신청서 작성 폼).

app.py 가 이미 디버그 포트 9222 로 attach 해 둔 **Selenium driver 를 그대로 받아서**:
  1) (과부하 시 자주 '멈추는') 신청서 작성 폼을 '진짜 페이지 이동 + 넉넉한 타임아웃 재시도' 로 뚫고,
  2) '지자체 선택' 팝업을 열어 원하는 지자체(예: 대전광역시)를 setLocal 로 선택한다.
저장/제출은 하지 않는다(사람 몫). 여기선 '입력 직전, 폼 앞까지 데려다주기' 만 한다.

── 왜 이렇게 하나 (2026-07-07 진단) ─────────────────────────────
  · 서버는 정상. 폼 HTML 은 HTTP/1.1 로 즉시 200 응답(과부하로 죽은 게 아님).
  · 그런데 브라우저에서 sellerApplyform 최초 로딩이 pnp4web 보안모듈 핸드셰이크 + 백엔드
    처리 때문에 한 번에 ~13초 걸리고, 가끔 응답이 끊겨(ERR_EMPTY_RESPONSE) '멈춘 듯' 보인다.
  · 인페이지 fetch 는 보안모듈이 막으므로 소용없다 → '실제 내비게이션' 을 20초+ 타임아웃으로 재시도해야 잡힌다.
  · 한 번에 요청 1개(사람이 F5 누르는 수준). 병렬 폭주/대기열 우회 없음. 성공하는 즉시 멈춘다.
"""
import re
import time

from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException

EV_BASE = "https://ev.or.kr/ev_ps/ps/seller/"   # 전기: car_type=11 승용 / 12 화물 / 13 승합
H2_BASE = "https://ev.or.kr/ev_ps/h2/seller/"   # 수소: car_type 없음(단일 폼)
FORM_SUFFIX = "sellerApplyform"


def _emit(log, msg, color=None):
    """Worker.log(msg, color) 와 호환. 색 인자를 못 받는 콜백이면 메시지만."""
    if not log:
        return
    try:
        log(msg, color)
    except TypeError:
        log(msg)


def _urls(car_kind, car_type):
    """(선택페이지 URL, 신청서폼 URL) 반환. 수소는 car_type 미사용."""
    base = H2_BASE if str(car_kind).strip() == "수소" else EV_BASE
    ct = str(car_type or "").strip()
    q = f"?car_type={ct}" if (ct and str(car_kind).strip() != "수소") else ""
    return base + "sellerApplyWrite" + q, base + "sellerApplyform" + q


def _set_timeout(driver, sec):
    try:
        driver.set_page_load_timeout(sec)
    except Exception:
        pass


def _landed(driver):
    """신청서 작성 폼에 실제로 안착했는지 — URL + 폼 필드 존재로 이중 확인
    (chrome-error 페이지를 성공으로 오인하지 않도록)."""
    try:
        if FORM_SUFFIX not in driver.current_url:
            return False
        return bool(driver.execute_script(
            "return !!(document.getElementById('local_nm')"
            "||document.getElementById('req_nm')"
            "||document.getElementById('req_kind'));"))
    except Exception:
        return False


def _punch_to_form(driver, sel_url, form_url, max_min, gap, nav_timeout, log):
    """폼이 뜰 때까지 '실제 내비게이션' 단일 재시도. 성공 시 True."""
    end = time.time() + max_min * 60
    n = 0
    while time.time() < end:
        n += 1
        # origin/세션 복구용으로 가벼운 선택 페이지 먼저(짧은 타임아웃).
        _set_timeout(driver, 12)
        try:
            driver.get(sel_url)
        except Exception:
            _stop_loading(driver)
        # 본 타깃: 신청서 폼(넉넉한 타임아웃 — 성공 자체가 ~13초).
        _set_timeout(driver, nav_timeout)
        t = time.time()
        status = "loaded"
        try:
            driver.get(form_url)
        except TimeoutException:
            status = "timeout"
            _stop_loading(driver)
        except WebDriverException as e:
            status = "err:" + str(e).splitlines()[0][:40]
            _stop_loading(driver)
        dt = round(time.time() - t, 1)
        ok = _landed(driver)
        _emit(log, f"   #{n} {status} {dt}s 진입={ok}")
        if ok:
            return True
        time.sleep(gap)
    return False


def _stop_loading(driver):
    try:
        driver.execute_script("window.stop();")
    except Exception:
        pass


def _wait_new_window(driver, before, timeout):
    end = time.time() + timeout
    while time.time() < end:
        try:
            new = set(driver.window_handles) - before
        except Exception:
            new = set()
        if new:
            return next(iter(new))
        time.sleep(0.2)
    return None


def _find_local_button(driver):
    """폼에서 '지자체 선택' 버튼 요소를 찾는다.
    주의: 상단 단계 트래커에도 '지자체선택' 텍스트 버튼이 있으나 팝업을 안 연다 →
    onclick 에 popupLocal 이 있는 '진짜' 버튼을 1순위로 잡는다."""
    try:
        els = driver.find_elements(By.XPATH, "//button|//a|//input")
    except Exception:
        return None
    # 1순위: onclick 에 popupLocal
    for el in els:
        try:
            if "popupLocal" in (el.get_attribute("onclick") or "") and el.is_displayed():
                return el
        except Exception:
            continue
    # 2순위: 텍스트가 정확히 '지자체 선택' 이고 클릭 가능(단계 트래커와 구분 위해 onclick 존재 요구)
    for el in els:
        try:
            txt = (el.text or el.get_attribute("value") or "").strip().replace(" ", "")
            if txt == "지자체선택" and (el.get_attribute("onclick") or "") and el.is_displayed():
                return el
        except Exception:
            continue
    return None


def _find_local_code(driver, name):
    """지자체 팝업에서 목표 지자체 행의 setLocal 코드값을 찾는다.
    행 버튼 onclick 예: setLocal('3000','대전광역시')."""
    ocs = driver.execute_script(
        "return [...document.querySelectorAll('[onclick]')]"
        ".map(e=>e.getAttribute('onclick')||'')"
        ".filter(o=>o.indexOf('setLocal')>=0);") or []
    exact = loose = None
    for oc in ocs:
        m = re.search(r"setLocal\('([0-9]+)'\s*,\s*'([^']+)'\)", oc)
        if not m:
            continue
        cd, nm = m.group(1), m.group(2)
        if nm == name:
            exact = cd
        elif name and (name in nm or nm in name):
            loose = loose or cd
    return exact or loose


def _wait_local(driver, name, timeout):
    """setLocal 후 폼이 다시 로드되며 지자체가 반영될 때까지 대기."""
    end = time.time() + timeout
    while time.time() < end:
        try:
            if FORM_SUFFIX in driver.current_url:
                v = driver.execute_script(
                    "var e=document.getElementById('local_nm');"
                    "return e?(e.value||e.textContent||''):'';") or ""
                v = v.strip()
                if v == name or (name and name in v):
                    return True
        except Exception:
            pass
        time.sleep(0.6)
    return False


def select_local(driver, local_name, log=None, timeout=25):
    """신청서 폼에서 '지자체 선택' 팝업을 열어 목표 지자체를 선택한다. 성공 시 True."""
    try:
        before = set(driver.window_handles)
        main = driver.current_window_handle
    except Exception as e:
        _emit(log, f"   창 상태 확인 실패: {e}", "red")
        return False

    # 지자체 선택 팝업 열기 — window.open 팝업차단을 피하려 '실제 버튼'을 Selenium 으로 클릭
    # (execute_script 로 popupLocal() 을 직접 부르면 사용자 제스처가 아니라 차단될 수 있음).
    btn = _find_local_button(driver)
    if btn is not None:
        try:
            btn.click()
        except Exception:
            try:
                driver.execute_script("arguments[0].click();", btn)
            except Exception as e:
                _emit(log, f"   지자체 선택 버튼 클릭 실패: {e}", "red")
                return False
    else:
        # 버튼을 못 찾으면 최후수단으로 함수 직접 호출
        try:
            driver.execute_script("popupLocal();")
        except Exception as e:
            _emit(log, f"   지자체 팝업 열기 실패(버튼 없음): {e}", "red")
            return False

    popup = _wait_new_window(driver, before, 10)
    if not popup:
        _emit(log, "   지자체 선택 팝업이 뜨지 않음", "red")
        return False

    driver.switch_to.window(popup)
    # 팝업 목록 로딩 대기
    end = time.time() + 10
    code = None
    while time.time() < end:
        code = _find_local_code(driver, local_name)
        if code:
            break
        time.sleep(0.3)

    if not code:
        _emit(log, f"   팝업 목록에서 '{local_name}' 을(를) 못 찾음", "red")
        try:
            driver.close()
            driver.switch_to.window(main)
        except Exception:
            pass
        return False

    # setLocal 을 팝업 컨텍스트에서 실행 → opener(메인 폼)에 반영 + 팝업 닫힘
    try:
        driver.execute_script("setLocal(arguments[0], arguments[1]);", code, local_name)
    except Exception as e:
        _emit(log, f"   지자체 setLocal 실패: {e}", "red")
        try:
            driver.close()
            driver.switch_to.window(main)
        except Exception:
            pass
        return False
    _emit(log, f"   지자체 선택: {local_name}({code})")

    time.sleep(0.5)
    # 메인 폼으로 복귀
    try:
        if main in driver.window_handles:
            driver.switch_to.window(main)
        else:  # 팝업 닫힘으로 핸들이 하나만 남았을 수 있음
            driver.switch_to.window(driver.window_handles[0])
    except Exception:
        pass

    if _wait_local(driver, local_name, timeout):
        return True
    _emit(log, f"   ⚠ 지자체 반영 확인 지연 — 화면에서 '{local_name}' 표시 직접 확인", "red")
    return False


def enter_form(driver, car_type="11", car_kind="전기", local_name=None,
               max_min=6.0, nav_timeout=22, gap=1.2, log=None):
    """
    신청서 작성 폼까지 진입(필요 시 지자체 선택 포함). 성공 시 True.
      car_type : 전기 11 승용 / 12 화물 / 13 승합 (수소는 무시)
      car_kind : '전기' 또는 '수소'
      local_name: 선택할 지자체명(예: '대전광역시'). None 이면 폼 진입만.
    저장/제출은 하지 않는다.
    """
    sel_url, form_url = _urls(car_kind, car_type)
    _emit(log, f"🚪 폼 진입 시작: {car_kind} car_type={car_type}"
               + (f" · 지자체 '{local_name}'" if local_name else ""))

    if _landed(driver):
        _emit(log, "   이미 신청서 폼 — 재진입 생략")
    else:
        if not _punch_to_form(driver, sel_url, form_url, max_min, gap, nav_timeout, log):
            _emit(log, "   ❌ 폼 진입 실패(상한 내 서버 지연 지속). 잠시 후 다시 시도.", "red")
            _set_timeout(driver, 300)
            return False
        _emit(log, "   ✅ 신청서 작성 폼 진입 성공")

    ok = True
    if local_name:
        ok = select_local(driver, local_name, log=log)
        if ok:
            _emit(log, f"   ✅ 지자체 '{local_name}' 선택 완료 — 입력 준비")

    _set_timeout(driver, 300)   # 이후 app.py 동작에 영향 없도록 기본값 복원
    return ok
