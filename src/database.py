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


def get_rules_from_table(rule_table_name: str) -> List[Dict[str, Any]]:
    """
    rule_table_name에 해당하는 테이블에서 모든 규칙 조회
    
    Args:
        rule_table_name: 규칙 테이블명 (예: "rule_B907")
        
    Returns:
        규칙 리스트 (priority 순서로 정렬)
    """
    if not rule_table_name:
        return []
    
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    try:
        # 동적 테이블명 사용 (주의: SQL injection 방지를 위해 테이블명 검증 필요)
        # 테이블명이 rule_로 시작하는지 확인
        if not rule_table_name.startswith("rule_"):
            return []
        
        cursor.execute(f"""
            SELECT * FROM "{rule_table_name}" 
            ORDER BY priority ASC, rule_id ASC
        """)
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows] if rows else []
    except sqlite3.OperationalError:
        # 테이블이 없으면 빈 리스트 반환
        conn.close()
        return []


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


def add_rule_to_table(
    rule_table_name: str,
    priority: int,
    status: str,
    repair_region: str,
    vehicle_classification: str,
    liability_ratio: float,
    amount_cap_type: str,
    project_code: str = "",
    part_name: str = "",
    part_no: str = "",
    exclude_project_code: str = "",
    warranty_mileage_override: int = None,
    warranty_period_override: int = None,
    amount_cap_value: int = None,
    note: str = "",
    valid_from: str = "",
    valid_to: str = "",
    engine_form: str = "",
) -> int:
    """
    rule 테이블에 규칙 추가
    
    Args:
        rule_table_name: 규칙 테이블명 (예: "rule_B907")
        priority: 우선순위
        status: 상태 (예: "ACTIVE", "INACTIVE")
        repair_region: 수리 지역
        vehicle_classification: 차량 분류
        liability_ratio: 구상율
        amount_cap_type: 금액 상한 타입
        project_code: 프로젝트 코드 (선택, 비워두면 모든 프로젝트)
        part_name: 부품명 (선택, 비워두면 모든 부품)
        part_no: 부품 번호 (선택)
        exclude_project_code: 제외 프로젝트 코드 (선택)
        warranty_mileage_override: 보증 주행거리 오버라이드 (선택)
        warranty_period_override: 보증 기간 오버라이드 (선택)
        amount_cap_value: 금액 상한 값 (선택)
        note: 비고 (선택)
        valid_from: 유효 시작일 (선택)
        valid_to: 유효 종료일 (선택)
        engine_form: 엔진 형태 (선택)
        
    Returns:
        추가된 rule_id
    """
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")
    
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    try:
        # 필수 필드 검증
        if not repair_region:
            raise ValueError("수리 지역은 필수입니다.")
        if not vehicle_classification:
            raise ValueError("차량 분류는 필수입니다.")
        if liability_ratio is None:
            raise ValueError("구상율은 필수입니다.")
        if not amount_cap_type:
            raise ValueError("금액 상한 타입은 필수입니다.")
        
        # INSERT 쿼리 실행
        cursor.execute(f"""
            INSERT INTO "{rule_table_name}" (
                priority, status, repair_region, project_code, exclude_project_code,
                vehicle_classification, part_no, part_name, engine_form,
                warranty_mileage_override, warranty_period_override,
                liability_ratio, amount_cap_type, amount_cap_value,
                note, valid_from, valid_to,
                created_at, updated_at
            ) VALUES (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?,
                ?, ?, ?,
                ?, ?, ?,
                DATETIME('now', 'localtime'), DATETIME('now', 'localtime')
            )
        """, (
            priority, status, repair_region, project_code or "", exclude_project_code or "",
            vehicle_classification, part_no or "", part_name or "", engine_form or "",
            warranty_mileage_override, warranty_period_override,
            liability_ratio, amount_cap_type, amount_cap_value,
            note or "", valid_from or "", valid_to or "",
        ))
        
        rule_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        return rule_id
    except sqlite3.OperationalError as e:
        conn.close()
        raise ValueError(f"Rule 추가 실패: {str(e)}")

