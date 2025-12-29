"""
임시 테스트 데이터 추가 스크립트
"""
from src.database import init_database, upsert_company

# DB 초기화
init_database()

# 테스트 데이터 추가
upsert_company(
    sap_code="B907",
    sap_name="AMS",
    rule_table_name="rule_B907",
    warranty_mileage=50000,
    warranty_period=3 * 365  # 3년을 일 단위로 변환 (1095일)
)

print("테스트 데이터가 추가되었습니다!")
print("SAP 코드: B907")
print("SAP 기업명: AMS")
print("규칙 테이블: rule_B907")
print("보증 주행거리: 50000 km")
print("보증 기간: 3년 (1095일)")

