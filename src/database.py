"""
SQLite 데이터베이스 관리 모듈
SAP 기업정보 저장 및 조회
"""
import sqlite3
from pathlib import Path
from typing import Optional, Dict, Any, List

# 데이터베이스 파일 경로
DB_PATH = Path("data/TestDB.sqlite")


def init_database():
    """데이터베이스 초기화 (테이블은 이미 존재하므로 연결만 확인)"""
    # data 폴더가 없으면 생성
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    # 연결 테스트
    conn = sqlite3.connect(str(DB_PATH))
    conn.close()


def get_company_info(sap_code_or_name: str) -> Optional[Dict[str, Any]]:
    """
    SAP 기업정보 조회 (sap_code 또는 sap_name으로 조회 가능)
    
    Args:
        sap_code_or_name: SAP 코드 또는 SAP 기업명
        
    Returns:
        기업정보 딕셔너리 (기존 코드 호환을 위해 필드명 변환)
    """
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # sap_code 또는 sap_name으로 조회
    cursor.execute("""
        SELECT * FROM sap WHERE sap_code = ? OR sap_name = ?
    """, (sap_code_or_name, sap_code_or_name))
    
    row = cursor.fetchone()
    conn.close()
    
    if row:
        data = dict(row)
        # warranty_period가 일 단위인지 년 단위인지 확인 필요
        # 일단 일 단위로 가정하고 365로 나눔 (필요시 수정)
        warranty_period_days = data.get("warranty_period")
        warranty_years = warranty_period_days / 365.0 if warranty_period_days else 2
        
        # 기존 코드 호환을 위한 필드명 매핑
        result = {
            "sap_code": data.get("sap_code"),
            "sap_name": data.get("sap_name"),
            "company_name": data.get("sap_name"),  # 호환성
            "mileage_threshold": data.get("warranty_mileage", 50000),
            "warranty_years": warranty_years,
            "warranty_mileage": data.get("warranty_mileage"),
            "warranty_period": data.get("warranty_period"),
            "rule_table_name": data.get("rule_table_name"),
            "sheet_index": 0,  # 기본값 (sap 테이블에 없음)
            "header_row": 3,   # 기본값
            "data_start_row": 4,  # 기본값
        }
        return result
    return None


def get_all_companies() -> List[str]:
    """모든 SAP 기업명 목록 조회 (sap_name 반환)"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    cursor.execute("SELECT sap_name FROM sap ORDER BY sap_name")
    rows = cursor.fetchall()
    conn.close()
    
    return [row[0] for row in rows] if rows else []


def get_all_sap_codes() -> List[str]:
    """모든 SAP 코드 목록 조회"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    cursor.execute("SELECT sap_code FROM sap ORDER BY sap_code")
    rows = cursor.fetchall()
    conn.close()
    
    return [row[0] for row in rows] if rows else []


def upsert_company(sap_code: str, sap_name: str = None, warranty_mileage: int = None, 
                   warranty_period: int = None, rule_table_name: str = None):
    """
    SAP 기업정보 저장/업데이트
    
    Args:
        sap_code: SAP 코드 (PK)
        sap_name: SAP 기업명
        warranty_mileage: 보증 주행거리
        warranty_period: 보증 기간 (일 단위)
        rule_table_name: 규칙 테이블명
    """
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    # 기존 데이터 확인
    existing = get_company_info(sap_code)
    
    if existing:
        # 업데이트
        updates = []
        values = []
        
        if sap_name is not None:
            updates.append("sap_name = ?")
            values.append(sap_name)
        if warranty_mileage is not None:
            updates.append("warranty_mileage = ?")
            values.append(warranty_mileage)
        if warranty_period is not None:
            updates.append("warranty_period = ?")
            values.append(warranty_period)
        if rule_table_name is not None:
            updates.append("rule_table_name = ?")
            values.append(rule_table_name)
        
        if updates:
            updates.append("updated_at = DATETIME('now', 'localtime')")
            values.append(sap_code)
            
            cursor.execute(f"""
                UPDATE sap 
                SET {", ".join(updates)}
                WHERE sap_code = ?
            """, values)
    else:
        # 삽입
        cursor.execute("""
            INSERT INTO sap (sap_code, sap_name, warranty_mileage, warranty_period, rule_table_name)
            VALUES (?, ?, ?, ?, ?)
        """, (sap_code, sap_name, warranty_mileage, warranty_period, rule_table_name))
    
    conn.commit()
    conn.close()

