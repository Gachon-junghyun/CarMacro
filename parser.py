# -*- coding: utf-8 -*-
"""세로형(한 파일=한 신청서) 양식 파서.

라벨 셀을 찾아 같은 행 오른쪽의 값을 읽는다. 안전하게 매핑되는 값만
site 필드 id 로 변환하고, 추측이 위험한 항목은 handoff(사람이 직접) 로 분리한다.
"""
import re
import openpyxl


def _norm(s):
    return re.sub(r"\s+", "", str(s or "")).strip()


def _is_empty_val(v):
    s = str(v or "").strip()
    return s == "" or s == "."


def _date_norm(v):
    """2026.6.17 / 2026-6-17 / 2026년 6월 17일 / 20260617 → 2026-06-17"""
    s = str(v or "").strip()
    parts = [p for p in re.split(r"\D+", s) if p]
    if len(parts) == 3:
        y, m, d = parts
        if len(y) == 4:
            return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
    digits = re.sub(r"\D", "", s)
    if len(digits) == 8:
        return f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]}"
    return s


def _birth6(v):
    """1978-06-11 → 780611 (YYMMDD)"""
    digits = re.sub(r"\D", "", str(v or ""))
    if len(digits) == 8:       # YYYYMMDD
        return digits[2:8]
    if len(digits) == 6:
        return digits
    return ""


# 값 변환 사전 (정규화된 한글 → site 코드)
MOBILE_RE = re.compile(r"01[016789]-?\d{3,4}-?\d{4}")
REQ_KIND = {"개인": "P", "개인사업자": "B", "단체": "G"}
SEX = {"남": "M", "남자": "M", "여": "F", "여자": "F"}
PRIORITY = {"우선": "10", "우선순위": "10", "법인기관": "20", "법인·기관": "20",
            "중소기업": "30", "택배": "50", "일반": "00"}
YN = {"Y": "Y", "N": "N", "예": "Y", "아니오": "N", "O": "Y"}


def parse_vertical(path):
    """반환: dict(mapped, handoff, notes, raw)
    - mapped:  site 필드 id -> 값 (자동 입력 대상)
    - special: 텍스트 매칭이 필요한 항목 (model_cd, social_kind 한글명)
    - handoff: [(라벨, 값, 사유)] 사람이 직접 확인/입력
    - notes:   파싱 경고
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # 셀 그리드: (row, col) -> 문자열
    grid = {}
    maxr = ws.max_row
    maxc = ws.max_column
    for r in range(1, maxr + 1):
        for c in range(1, maxc + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                grid[(r, c)] = str(v).strip()

    def value_right(r, c):
        """같은 행에서 (r,c) 오른쪽 첫 비어있지 않은 셀 값."""
        for cc in range(c + 1, maxc + 1):
            if (r, cc) in grid:
                return grid[(r, cc)], cc
        return "", None

    # 라벨 위치 모으기: 정규화 라벨 -> [(r,c), ...]
    label_pos = {}
    for (r, c), v in grid.items():
        label_pos.setdefault(_norm(v), []).append((r, c))

    mapped, special, handoff, notes = {}, {}, [], []
    used = set()  # 값으로 소비한 셀 좌표 (중복 방지)

    def get(label):
        """라벨에 매칭되는 첫 값과 좌표. 없으면 ('', None)."""
        for key, positions in label_pos.items():
            if key == _norm(label):
                for (r, c) in positions:
                    val, vc = value_right(r, c)
                    if vc is not None:
                        return val, (r, vc)
        return "", None

    def get_all(label):
        """동일 라벨이 여러 번일 때 모든 (값, 좌표)."""
        res = []
        for (r, c) in label_pos.get(_norm(label), []):
            val, vc = value_right(r, c)
            if vc is not None:
                res.append((val, (r, vc)))
        return res

    def mobile_in_row_of(label):
        """라벨이 있는 행에서 휴대폰 패턴 값을 찾음(라벨 없는 신청자 휴대폰용)."""
        for (r, c) in label_pos.get(_norm(label), []):
            for cc in range(1, maxc + 1):
                val = grid.get((r, cc), "")
                m = MOBILE_RE.search(val.replace(" ", ""))
                if m:
                    return m.group(0)
        return ""

    # ── 안전 매핑 (텍스트/확정 코드) ──────────────────────────
    # 전기/수소 구분 (열린 폼과 대조용 — 사고 방지)
    v, _ = get("구분")
    if v:
        nv = _norm(v)
        if "수소" in nv:
            mapped["car_kind"] = "수소"
        elif "전기" in nv:
            mapped["car_kind"] = "전기"

    v, _ = get("계약일자")
    if v:
        mapped["contract_day"] = _date_norm(v)

    v, _ = get("신청유형")
    if v:
        code = REQ_KIND.get(_norm(v))
        if code:
            mapped["req_kind"] = code
        else:
            handoff.append(("신청유형", v, "req_kind 코드 매칭 실패"))

    v, _ = get("성명")
    if v:
        mapped["req_nm"] = v

    v, _ = get("생년월일")
    if v:
        d8 = _date_norm(v)   # 1978-06-11 / 1982.3.22 → YYYY-MM-DD
        if re.match(r"^\d{4}-\d{2}-\d{2}$", d8):
            mapped["birth1"] = d8
        elif _birth6(v):
            mapped["birth1"] = _birth6(v)   # 6자리만 있으면 그대로
        else:
            handoff.append(("생년월일", v, "형식 불명 → 직접 확인"))

    v, _ = get("성별")
    if v:
        code = SEX.get(_norm(v))
        if code:
            mapped["req_sex"] = code
        else:
            handoff.append(("성별", v, "M/F 매칭 실패"))

    # 사업자 정보
    v, _ = get("사업자번호")
    if v:
        mapped["busi_no"] = v
    v, _ = get("사업자명")
    if v:
        mapped["pri_busi_nm"] = v

    # 신청차종 → 라이브 옵션 텍스트 매칭 필요
    v, _ = get("신청차종")
    if v:
        special["model_cd"] = v   # 한글 차종명

    v, _ = get("신청대수")
    if v:
        mapped["req_cnt"] = re.sub(r"\D", "", v) or v

    v, _ = get("출고예정일자")
    if v:
        mapped["delivery_sch_day"] = _date_norm(v)

    # 주소
    v, _ = get("주소")
    if v:
        mapped["addr"] = v
    v, _ = get("이하주소")
    if v:
        mapped["addr_detail"] = v
    # 우편번호는 데이터에 없음 → 주소(도로명)로 검색 팝업 자동화(결과 1개일 때 자동 선택)
    notes.append("※ 우편번호는 '주소'(도로명)로 주소검색 팝업을 자동 실행합니다. "
                 "결과가 정확히 1개면 자동 선택, 여러 개/0개면 팝업에서 직접 선택.")

    # 연락처: 휴대폰은 '신청자 휴대폰'(mobile), 전화번호 '.'은 빈값
    # 신청자 휴대폰 = E11 의 010-... (라벨 없이 값만). 담당자 휴대폰과 구분 주의.
    # 전화/이메일은 필수(*) 칸이라 데이터의 '.' 도 그대로 입력(빈칸만 건너뜀)
    v, _ = get("전화번호")
    if str(v).strip():
        mapped["phone"] = str(v).strip()
    v, _ = get("이메일")
    if str(v).strip():
        mapped["email"] = str(v).strip()
    # 신청자 휴대폰: 전화번호/이메일 행에 라벨 없이 들어있는 010-... 값
    mob = mobile_in_row_of("이메일") or mobile_in_row_of("전화번호")
    if mob:
        mapped["mobile"] = mob

    # 사회계층여부 (라벨 2개: 여부 Y/N + 유형 한글)
    socs = get_all("사회계층여부")
    social_yn = ""
    social_kind_txt = ""
    for val, _pos in socs:
        nv = _norm(val)
        if nv in YN:
            social_yn = YN[nv]
        elif val.strip():
            social_kind_txt = val.strip()
    if social_yn:
        mapped["social_yn"] = social_yn
    if social_yn == "Y" and social_kind_txt:
        special["social_kind"] = social_kind_txt   # '소상공인' 등 텍스트 매칭

    # 우선순위
    v, _ = get("우선순위 배정.집행 선택")
    if not v:
        v, _ = get("우선순위")
    if v:
        code = PRIORITY.get(_norm(v))
        if code:
            mapped["priority_type"] = code
        else:
            handoff.append(("우선순위", v, "priority_type 매칭 실패"))

    # 담당자성명 → contact_nm, 담당자 휴대폰 → contact_mobile (라벨 '휴대폰')
    v, _ = get("담당자성명")
    if v:
        mapped["contact_nm"] = v
    v, _ = get("휴대폰")
    if v and MOBILE_RE.search(v.replace(" ", "")):
        mapped["contact_mobile"] = v
    # 계약번호 → 제조수입사 관리(계약)번호 (seller_mgrid)
    v, _ = get("계약번호")
    if v:
        mapped["seller_mgrid"] = v

    # ── 위험/특수: 사람이 직접 ───────────────────────────────
    for lbl, reason in [
        ("경유화물차 폐차 미이행여부", "대응 필드 불명확 → 직접 선택"),
        ("차량번호", "교체(폐차)차량 섹션 → 직접"),
        ("차대번호", "교체(폐차)차량 섹션 → 직접"),
        ("보유모델명", "교체(폐차)차량 섹션 → 직접"),
    ]:
        v, _ = get(lbl)
        if v:
            handoff.append((lbl, v, reason))

    # 암호 역순
    handoff.append(("암호 역순 입력", "제출 직전 단계", "캡차성 단계 → 사람이 직접 (자동화 안 함)"))

    return {"mapped": mapped, "special": special, "handoff": handoff, "notes": notes}


if __name__ == "__main__":
    import os
    import sys
    default = os.path.join(os.path.dirname(__file__), "examples", "real_sample_ev.xlsx")
    path = sys.argv[1] if len(sys.argv) > 1 else default
    res = parse_vertical(path)
    print("=== 자동 입력(mapped) ===")
    for k, v in res["mapped"].items():
        print(f"  {k:18} = {v}")
    print("\n=== 텍스트 매칭 필요(special) ===")
    for k, v in res["special"].items():
        print(f"  {k:18} = {v}")
    print("\n=== 사람이 직접(handoff) ===")
    for lbl, val, reason in res["handoff"]:
        print(f"  · {lbl} = '{val}'  ({reason})")
    print("\n=== 메모 ===")
    for n in res["notes"]:
        print("  ", n)
