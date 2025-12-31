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
            "remark": data.get("remark", ""),  # remark 추가
            "renault_code": data.get("renault_code", ""),  # renault_code 추가
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


def get_all_companies_with_code() -> List[Dict[str, str]]:
    """모든 SAP 기업 정보 조회 (sap_code와 sap_name 반환)"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("SELECT sap_code, sap_name FROM sap ORDER BY sap_name")
    rows = cursor.fetchall()
    conn.close()
    
    return [{"sap_code": row["sap_code"], "sap_name": row["sap_name"]} for row in rows] if rows else []


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


def create_rule_table(rule_table_name: str, cursor=None) -> bool:
    """
    룰 테이블 생성
    
    Args:
        rule_table_name: 규칙 테이블명 (예: "rule_B907")
        cursor: 기존 커서 (None이면 새 연결 생성)
        
    Returns:
        생성 성공 여부
    """
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")
    
    use_existing_cursor = cursor is not None
    if not use_existing_cursor:
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()
    
    try:
        # 테이블이 이미 존재하는지 확인
        cursor.execute("""
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name=?
        """, (rule_table_name,))
        
        if cursor.fetchone():
            # 테이블이 이미 존재함
            if not use_existing_cursor:
                conn.close()
            return True
        
        # 룰 테이블 생성
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS "{rule_table_name}" (
                rule_id INTEGER PRIMARY KEY AUTOINCREMENT,
                priority INTEGER NOT NULL DEFAULT -1,
                status TEXT NOT NULL DEFAULT 'ACTIVE' CHECK (status IN ('ACTIVE','INACTIVE')),
                repair_region TEXT NOT NULL CHECK (repair_region IN ('DOMESTIC','OVERSEAS','ALL')),
                project_code TEXT NOT NULL DEFAULT 'ALL',
                exclude_project_code TEXT,
                vehicle_classification TEXT NOT NULL DEFAULT 'ALL',
                part_no TEXT NOT NULL DEFAULT 'ALL',
                part_name TEXT NOT NULL DEFAULT 'ALL',
                engine_form TEXT NOT NULL DEFAULT 'ALL',
                warranty_mileage_override INTEGER,
                warranty_period_override INTEGER,
                liability_ratio REAL NOT NULL,
                amount_cap_type TEXT NOT NULL DEFAULT 'NONE' CHECK (amount_cap_type IN ('LABOR','OUTSOURCE_LABOR','BOTH_LABOR','NONE')),
                amount_cap_value INTEGER,
                valid_from TEXT CHECK (valid_from IS NULL OR date(valid_from) IS NOT NULL),
                valid_to TEXT CHECK (valid_to IS NULL OR date(valid_to) IS NOT NULL),
                created_at TEXT DEFAULT (DATETIME('now', 'localtime')),
                updated_at TEXT DEFAULT (DATETIME('now', 'localtime'))
            )
        """)
        
        if not use_existing_cursor:
            conn.commit()
            conn.close()
        return True
    except sqlite3.Error as e:
        if not use_existing_cursor:
            conn.close()
        raise ValueError(f"룰 테이블 생성 실패: {str(e)}")


def upsert_company(sap_code: str, sap_name: str = None, warranty_mileage: int = None, 
                   warranty_period: int = None, rule_table_name: str = None, renault_code: str = None):
    """
    SAP 기업정보 저장/업데이트
    
    Args:
        sap_code: SAP 코드 (PK)
        sap_name: SAP 기업명
        warranty_mileage: 보증 주행거리
        warranty_period: 보증 기간 (일 단위)
        rule_table_name: 규칙 테이블명
        renault_code: 르노 코드
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
        if renault_code is not None:
            updates.append("renault_code = ?")
            values.append(renault_code)
        
        if updates:
            updates.append("updated_at = DATETIME('now', 'localtime')")
            values.append(sap_code)
            
            cursor.execute(f"""
                UPDATE sap 
                SET {", ".join(updates)}
                WHERE sap_code = ?
            """, values)
        
        # 기존 협력사 업데이트 시에도 룰 테이블이 없으면 생성
        final_rule_table_name = rule_table_name or existing.get("rule_table_name")
        if final_rule_table_name:
            try:
                create_rule_table(final_rule_table_name, cursor)
            except Exception as e:
                # 룰 테이블 생성 실패해도 업데이트는 진행
                pass
    else:
        # 삽입 (새 협력사 추가)
        cursor.execute("""
            INSERT INTO sap (sap_code, sap_name, warranty_mileage, warranty_period, rule_table_name, renault_code)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (sap_code, sap_name, warranty_mileage, warranty_period, rule_table_name, renault_code))
        
        # 새 협력사 추가 시 룰 테이블도 생성 (같은 트랜잭션에서)
        if rule_table_name:
            try:
                create_rule_table(rule_table_name, cursor)
            except Exception as e:
                # 룰 테이블 생성 실패 시 롤백
                conn.rollback()
                conn.close()
                raise ValueError(f"협력사 추가 실패: {str(e)}")
    
    conn.commit()
    conn.close()


def update_company_remark(sap_code: str, remark: str) -> bool:
    """
    SAP 기업의 remark 업데이트
    
    Args:
        sap_code: SAP 코드
        remark: remark 내용
        
    Returns:
        성공 여부
    """
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
            UPDATE sap 
            SET remark = ?, updated_at = DATETIME('now', 'localtime')
            WHERE sap_code = ?
        """, (remark, sap_code))
        
        conn.commit()
        conn.close()
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        conn.close()
        raise ValueError(f"Remark 업데이트 실패: {str(e)}")


def add_rule_to_table(
    rule_table_name: str,
    status: str,
    repair_region: str,
    vehicle_classification: str,
    liability_ratio: float,
    amount_cap_type: str,
    project_code: str = "ALL",
    part_name: str = "ALL",
    part_no: str = "ALL",
    engine_form: str = "ALL",
    exclude_project_code: str = None,
    warranty_mileage_override: int = None,
    warranty_period_override: int = None,
    amount_cap_value: int = None,
    valid_from: str = None,
    valid_to: str = None,
    priority: int = None,  # None이면 -1로 설정되어 트리거가 자동으로 채움
) -> int:
    """
    rule 테이블에 규칙 추가
    
    Args:
        rule_table_name: 규칙 테이블명 (예: "rule_B907")
        status: 상태 (예: "ACTIVE", "INACTIVE") - DEFAULT 'ACTIVE'
        repair_region: 수리 지역 ('DOMESTIC','OVERSEAS','ALL')
        vehicle_classification: 차량 분류 - DEFAULT 'ALL'
        liability_ratio: 구상율 (필수)
        amount_cap_type: 금액 상한 타입 ('LABOR','OUTSOURCE_LABOR','BOTH_LABOR','NONE') - DEFAULT 'NONE'
        project_code: 프로젝트 코드 - DEFAULT 'ALL'
        part_name: 부품명 - DEFAULT 'ALL'
        part_no: 부품 번호 - DEFAULT 'ALL'
        engine_form: 엔진 형태 - DEFAULT 'ALL'
        exclude_project_code: 제외 프로젝트 코드 (NULL 허용)
        warranty_mileage_override: 보증 주행거리 오버라이드 (NULL 허용)
        warranty_period_override: 보증 기간 오버라이드 (NULL 허용, 일 단위)
        amount_cap_value: 금액 상한 값 (NULL 허용)
        valid_from: 유효 시작일 (NULL 허용, YYYY-MM-DD 형식)
        valid_to: 유효 종료일 (NULL 허용, YYYY-MM-DD 형식)
        priority: 우선순위 (None이면 -1로 설정되어 트리거가 자동으로 채움)
        
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
        if repair_region not in ["DOMESTIC", "OVERSEAS", "ALL"]:
            raise ValueError("수리 지역은 DOMESTIC, OVERSEAS, ALL 중 하나여야 합니다.")
        
        if not vehicle_classification:
            vehicle_classification = "ALL"
        
        if liability_ratio is None:
            raise ValueError("구상율은 필수입니다.")
        
        if not amount_cap_type:
            amount_cap_type = "NONE"
        if amount_cap_type not in ["LABOR", "OUTSOURCE_LABOR", "BOTH_LABOR", "NONE"]:
            raise ValueError("금액 상한 타입은 LABOR, OUTSOURCE_LABOR, BOTH_LABOR, NONE 중 하나여야 합니다.")
        
        if not project_code:
            project_code = "ALL"
        if not part_name:
            part_name = "ALL"
        if not part_no:
            part_no = "ALL"
        if not engine_form:
            engine_form = "ALL"
        
        if not status:
            status = "ACTIVE"
        if status not in ["ACTIVE", "INACTIVE"]:
            raise ValueError("상태는 ACTIVE 또는 INACTIVE여야 합니다.")
        
        # Priority: None이면 현재 테이블의 최대 우선순위 + 1로 설정
        if priority is None:
            cursor.execute(f'SELECT MAX(priority) FROM "{rule_table_name}"')
            max_priority = cursor.fetchone()[0]
            if max_priority is None:
                priority = 1  # 첫 번째 규칙
            else:
                priority = max_priority + 1
        
        # 날짜 형식 검증
        if valid_from and valid_from.strip():
            try:
                from datetime import datetime
                datetime.strptime(valid_from.strip(), "%Y-%m-%d")
            except ValueError:
                raise ValueError("유효 시작일은 YYYY-MM-DD 형식이어야 합니다.")
        
        if valid_to and valid_to.strip():
            try:
                from datetime import datetime
                datetime.strptime(valid_to.strip(), "%Y-%m-%d")
            except ValueError:
                raise ValueError("유효 종료일은 YYYY-MM-DD 형식이어야 합니다.")
        
        # INSERT 쿼리 실행 (note 컬럼 제거됨)
        cursor.execute(f"""
            INSERT INTO "{rule_table_name}" (
                priority, status, repair_region, project_code, exclude_project_code,
                vehicle_classification, part_no, part_name, engine_form,
                warranty_mileage_override, warranty_period_override,
                liability_ratio, amount_cap_type, amount_cap_value,
                valid_from, valid_to,
                created_at, updated_at
            ) VALUES (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?,
                ?, ?, ?,
                ?, ?,
                DATETIME('now', 'localtime'), DATETIME('now', 'localtime')
            )
        """, (
            priority, status, repair_region, project_code, exclude_project_code,
            vehicle_classification, part_no, part_name, engine_form,
            warranty_mileage_override, warranty_period_override,
            liability_ratio, amount_cap_type, amount_cap_value,
            valid_from, valid_to,
        ))
        
        rule_id = cursor.lastrowid
        conn.commit()
        conn.close()
        
        return rule_id
    except sqlite3.OperationalError as e:
        conn.close()
        raise ValueError(f"Rule 추가 실패: {str(e)}")


def update_rule_in_table(
    rule_table_name: str,
    rule_id: int,
    priority: int = None,
    status: str = None,
    repair_region: str = None,
    vehicle_classification: str = None,
    liability_ratio: float = None,
    amount_cap_type: str = None,
    project_code: str = None,
    part_name: str = None,
    part_no: str = None,
    exclude_project_code: str = None,
    warranty_mileage_override: int = None,
    warranty_period_override: int = None,
    amount_cap_value: int = None,
    valid_from: str = None,
    valid_to: str = None,
    engine_form: str = None,
) -> bool:
    """
    rule 테이블의 규칙 수정
    
    Returns:
        성공 여부
    """
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")
    
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    try:
        updates = []
        values = []
        
        if priority is not None:
            updates.append("priority = ?")
            values.append(priority)
        if status is not None:
            updates.append("status = ?")
            values.append(status)
        if repair_region is not None:
            updates.append("repair_region = ?")
            values.append(repair_region)
        if vehicle_classification is not None:
            updates.append("vehicle_classification = ?")
            values.append(vehicle_classification)
        if liability_ratio is not None:
            updates.append("liability_ratio = ?")
            values.append(liability_ratio)
        if amount_cap_type is not None:
            updates.append("amount_cap_type = ?")
            values.append(amount_cap_type)
        if project_code is not None:
            updates.append("project_code = ?")
            values.append(project_code)
        if part_name is not None:
            updates.append("part_name = ?")
            values.append(part_name)
        if part_no is not None:
            updates.append("part_no = ?")
            values.append(part_no)
        if exclude_project_code is not None:
            updates.append("exclude_project_code = ?")
            values.append(exclude_project_code)
        if warranty_mileage_override is not None:
            updates.append("warranty_mileage_override = ?")
            values.append(warranty_mileage_override)
        if warranty_period_override is not None:
            updates.append("warranty_period_override = ?")
            values.append(warranty_period_override)
        if amount_cap_value is not None:
            updates.append("amount_cap_value = ?")
            values.append(amount_cap_value)
        if valid_from is not None:
            updates.append("valid_from = ?")
            values.append(valid_from)
        if valid_to is not None:
            updates.append("valid_to = ?")
            values.append(valid_to)
        if engine_form is not None:
            updates.append("engine_form = ?")
            values.append(engine_form)
        
        if not updates:
            conn.close()
            return False
        
        updates.append("updated_at = DATETIME('now', 'localtime')")
        values.append(rule_id)
        
        cursor.execute(f"""
            UPDATE "{rule_table_name}"
            SET {", ".join(updates)}
            WHERE rule_id = ?
        """, values)
        
        conn.commit()
        conn.close()
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        conn.close()
        raise ValueError(f"Rule 수정 실패: {str(e)}")


def update_rule_priorities(rule_table_name: str, rule_ids_in_order: List[int]) -> bool:
    """
    드래그 앤 드롭으로 변경된 순서에 따라 priority 재할당
    
    Args:
        rule_table_name: 규칙 테이블명 (예: "rule_B907")
        rule_ids_in_order: 새로운 순서대로 정렬된 rule_id 리스트
        
    Returns:
        성공 여부
    """
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")
    
    if not rule_ids_in_order:
        return True
    
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    try:
        # 순서대로 priority 재할당 (1부터 시작)
        for new_priority, rule_id in enumerate(rule_ids_in_order, start=1):
            cursor.execute(f'''
                UPDATE "{rule_table_name}" 
                SET priority = ?, updated_at = DATETIME('now', 'localtime')
                WHERE rule_id = ?
            ''', (new_priority, rule_id))
        
        conn.commit()
        conn.close()
        return True
    except sqlite3.Error as e:
        conn.rollback()
        conn.close()
        raise ValueError(f"우선순위 업데이트 실패: {str(e)}")


def delete_rule_from_table(rule_table_name: str, rule_id: int) -> bool:
    """
    rule 테이블에서 규칙 삭제
    
    Returns:
        성공 여부
    """
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")
    
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    try:
        cursor.execute(f"""
            DELETE FROM "{rule_table_name}"
            WHERE rule_id = ?
        """, (rule_id,))
        
        conn.commit()
        conn.close()
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        conn.close()
        raise ValueError(f"Rule 삭제 실패: {str(e)}")

