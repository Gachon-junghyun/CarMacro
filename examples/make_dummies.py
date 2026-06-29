# -*- coding: utf-8 -*-
"""전기/수소 세로형 더미 엑셀 2개 생성 (구분 칸 포함)."""
import os
import openpyxl


def build(path, gubun, car_kind, model_name, name, busi_no, busi_nm,
          birth, mobile, contact_nm, contact_mobile, contract_no):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "신청서"
    cells = {
        "B2": "구분", "C2": gubun,                       # 전기 / 수소 (안전 구분)
        "B3": "계약일자", "C3": "2026-04-23",
        "B4": "신청유형", "C4": "개인사업자",
        "B5": "성명", "C5": name, "D5": "생년월일", "E5": birth, "F5": "성별", "G5": "남",
        "B6": "개인사업자", "D6": "사업자번호", "E6": busi_no, "F6": "사업자명", "G6": busi_nm,
        "B7": "신청차종", "C7": model_name,
        "D7": "신청대수", "E7": "1", "F7": "출고예정일자", "G7": "2026.6.17",
        "F8": "보조금금액", "G8": "1840만원",
        "B9": "주소", "C9": "세종특별자치시 시청대로 78",
        "B10": "이하주소", "C10": "306동 401호",
        "B11": "전화번호", "C11": ".", "E11": mobile, "F11": "이메일", "G11": ".",
        "B12": "사회계층여부", "C12": "Y", "D12": "사회계층여부", "E12": "소상공인",
        "B13": "경유화물차 폐차 미이행여부", "C13": "Y",
        "B14": "차량번호", "C14": "83누4742", "D14": "차대번호", "E14": "KPACA4AE1HP293610",
        "F14": "보유모델명", "G14": "코란도스포츠",
        "B15": "우선순위 배정.집행 선택", "C15": "우선",
        "B16": "제조사연락처", "F16": "전화번호", "G16": "공란 가능",
        "B17": "담당자성명", "C17": contact_nm, "D17": "휴대폰", "E17": contact_mobile,
        "F17": "계약번호", "G17": contract_no,
        "B19": "6o3BHt1ZRW", "C19": "WRZ1tHB3o6",
    }
    for addr, val in cells.items():
        ws[addr] = val
    wb.save(path)
    print("created:", path, "| 구분:", gubun, "| 차종:", model_name)


base = os.path.dirname(os.path.abspath(__file__))

# 전기차: 캐스퍼(목록에 정확히 있음) — 정상 매칭 확인용
build(os.path.join(base, "real_sample_ev.xlsx"), "전기", "전기",
      "캐스퍼 일렉트릭 기본형 15인치", "윤정수", "143-02-30294", "씨엔티솔루션",
      "1978-06-11", "010-9420-1380", "박수경", "010-7747-3849", "I6226US000010")

# 수소차: 디올뉴넥쏘(수소 목록에 있음)
build(os.path.join(base, "real_sample_h2.xlsx"), "수소", "수소",
      "디올뉴넥쏘", "김수소", "211-03-40521", "에이치투모빌리티",
      "1982-03-22", "010-2200-3300", "이담당", "010-4400-5500", "H2226US000077")
