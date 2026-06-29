# -*- coding: utf-8 -*-
"""더미 신청서 엑셀 생성 — ev.or.kr 실제 코드값에 맞춤."""
import pandas as pd

# 열 머리글 = app.py 필드 id
cols = ["local_nm", "req_kind", "contract_day", "req_nm", "birth1", "birth2", "req_sex",
        "model_cd", "req_cnt", "delivery_sch_day",
        "zipno", "addr", "addr_detail",
        "phone", "mobile", "email",
        "improve_fd_yn", "first_buy_yn", "social_yn", "school_bus_yn",
        "in_facility_yn", "disaster_scrap_yn", "bms_yn", "priority_type"]

# req_kind: P=개인, B=법인사업자, G=단체
# req_sex: M / F
# model_cd: 실제 차종 코드 (예시)
# priority_type: 10/20/30/40/50, 00=해당없음
# *_yn: Y / N
rows = [
    {"local_nm": "세종특별자치시", "req_kind": "P", "contract_day": "2026-06-01", "req_nm": "홍길동",
     "birth1": "900101", "birth2": "900101-1234567", "req_sex": "M",
     "model_cd": "CASPER_S15", "req_cnt": "1", "delivery_sch_day": "2026-07-15",
     "zipno": "30151", "addr": "세종특별자치시 한누리대로 2130", "addr_detail": "101동 1001호",
     "phone": "", "mobile": "010-1111-2222", "email": "hong@example.com",
     "improve_fd_yn": "N", "first_buy_yn": "Y", "social_yn": "N", "school_bus_yn": "N",
     "in_facility_yn": "N", "disaster_scrap_yn": "N", "bms_yn": "N", "priority_type": "00"},
    {"local_nm": "세종특별자치시", "req_kind": "P", "contract_day": "2026-06-03", "req_nm": "김영희",
     "birth1": "851212", "birth2": "851212-2234567", "req_sex": "F",
     "model_cd": "KONA_2WD_S17", "req_cnt": "1", "delivery_sch_day": "2026-07-20",
     "zipno": "30151", "addr": "세종특별자치시 도움5로 20", "addr_detail": "202호",
     "phone": "", "mobile": "010-3333-4444", "email": "kim@example.com",
     "improve_fd_yn": "N", "first_buy_yn": "N", "social_yn": "Y", "school_bus_yn": "N",
     "in_facility_yn": "N", "disaster_scrap_yn": "N", "bms_yn": "N", "priority_type": "00"},
    {"local_nm": "세종특별자치시", "req_kind": "P", "contract_day": "2026-06-05", "req_nm": "이철수",
     "birth1": "780808", "birth2": "780808-1234567", "req_sex": "M",
     "model_cd": "IONIQ9_ACT_2WD", "req_cnt": "1", "delivery_sch_day": "2026-08-01",
     "zipno": "30151", "addr": "세종특별자치시 갈매로 388", "addr_detail": "303호",
     "phone": "", "mobile": "010-5555-6666", "email": "lee@example.com",
     "improve_fd_yn": "N", "first_buy_yn": "Y", "social_yn": "N", "school_bus_yn": "N",
     "in_facility_yn": "N", "disaster_scrap_yn": "N", "bms_yn": "N", "priority_type": "00"},
]

df = pd.DataFrame(rows, columns=cols)
out = r"C:\Users\fivep\OneDrive\Desktop\CarMacro\sample_applications.xlsx"
df.to_excel(out, index=False)
print("created:", out, "rows:", len(df))
