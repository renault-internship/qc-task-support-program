"""
SAP 테이블에 협력사 데이터 추가 스크립트
사용법: python insert_sap_data.py
"""
import sqlite3
from pathlib import Path
from src.database import init_database, upsert_company

# 데이터베이스 경로
DB_PATH = Path("data/TestDB.sqlite")

# 데이터베이스 초기화
init_database()

# 추가할 데이터
suppliers = [
    {"sap_name": "Hanrim Intech", "sap_code": "I806", "renault_code": "247744"},
    {"sap_name": "SMC Co.", "sap_code": "C508", "renault_code": "250034"},
    {"sap_name": "Dk Austech", "sap_code": "C202", "renault_code": "247751"},
    {"sap_name": "Hung-A Forming", "sap_code": "B801", "renault_code": "247752"},
    {"sap_name": "HBPO", "sap_code": "B933", "renault_code": "260864"},
    {"sap_name": "Dongwon Tech.", "sap_code": "I201", "renault_code": "247750"},
    {"sap_name": "AMS", "sap_code": "B907", "renault_code": "247736"},
]

# 디폴트값
default_warranty_mileage = 60000
default_warranty_period = 3 * 365  # 3년을 일 단위로 변환 (1095일)

print("SAP 테이블에 협력사 데이터 추가 중...")
print("=" * 50)

for supplier in suppliers:
    sap_code = supplier["sap_code"]
    sap_name = supplier["sap_name"]
    renault_code = supplier["renault_code"]
    
    # rule_table_name 자동 생성
    rule_table_name = f"rule_{sap_code}"
    
    try:
        # 기존 데이터 업데이트 또는 새로 추가
        upsert_company(
            sap_code=sap_code,
            sap_name=sap_name,
            warranty_mileage=default_warranty_mileage,
            warranty_period=default_warranty_period,
            rule_table_name=rule_table_name,
        )
        
        # renault_code는 별도로 업데이트 필요 (upsert_company에 renault_code 파라미터가 없으므로)
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()
        
        # renault_code 컬럼이 있는지 확인하고 업데이트
        cursor.execute("PRAGMA table_info(sap)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if "renault_code" in columns:
            cursor.execute("""
                UPDATE sap 
                SET renault_code = ? 
                WHERE sap_code = ?
            """, (renault_code, sap_code))
            conn.commit()
        
        conn.close()
        
        print(f"✓ {sap_name} ({sap_code}) - RENAULT CODE: {renault_code}")
    except Exception as e:
        print(f"✗ {sap_name} ({sap_code}) 추가 실패: {str(e)}")

print("=" * 50)
print("데이터 추가 완료!")

