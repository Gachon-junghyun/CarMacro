# -*- coding: utf-8 -*-
"""validate_rows 단위 테스트 — 정상/오류 케이스가 제대로 걸러지는지 확인."""
from app import validate_rows, FIELD_BY_ID

good = [{
    "local_nm": "세종특별자치시", "req_kind": "P", "contract_day": "2026-06-01",
    "req_nm": "홍길동", "birth1": "900101", "birth2": "900101-1234567",
    "req_sex": "M", "model_cd": "CASPER_S15", "req_cnt": "1",
    "delivery_sch_day": "2026-07-15", "mobile": "010-1111-2222",
    "email": "hong@example.com", "priority_type": "00",
}]
e, w = validate_rows(good, list(good[0].keys()))
assert e == [], f"정상행에 오류가 잡힘: {e}"
print("정상행 통과 OK (경고:", len(w), ")")

bad = [{
    "req_kind": "X",            # enum 위반
    "req_nm": "",              # 필수 비어있음
    "req_sex": "남",            # enum 위반
    "model_cd": "FOO",         # (정적검증은 통과, 라이브에서 잡힘)
    "req_cnt": "한대",          # 숫자 아님
    "mobile": "01012",         # 형식 이상
    # birth1/birth2 둘 다 없음
    "priority_type": "99",     # enum 위반
}]
e2, w2 = validate_rows(bad, list(bad[0].keys()))
print(f"\n오류행에서 잡은 오류 {len(e2)}건:")
for x in e2:
    print("  -", x)
# 핵심 오류들이 잡혔는지 확인
joined = " ".join(e2)
for must in ["req_kind", "req_nm", "req_sex", "생년월일", "req_cnt", "휴대폰", "priority_type"]:
    assert must in joined, f"'{must}' 오류가 안 잡힘"
print("\n오류행 모든 핵심오류 검출 OK")
print("\n전체 테스트 통과 [PASS]")
