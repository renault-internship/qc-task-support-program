"""
SQLite 데이터베이스 관리 모듈
SAP 기업정보 저장 및 조회

요구사항 반영(중요):
- sap 테이블에서 warranty_mileage / warranty_period 컬럼 제거(미사용/없다고 가정)
- 보증정보는 warranty 테이블(전역 1행, id=1)로 분리하여 관리
- get_company_info는 warranty(전역값) 우선으로 보증값을 반환하고, 없으면 기본값(60000/3) 사용

추가:
- common_project_liability 테이블 추가 (project_code -> 기본 liability_ratio)

✅ 이번 변경(딱 이것만):
- rule_{sap_code} 테이블에 note 컬럼 지원
  1) 새로 생성되는 rule 테이블 스키마에 note 컬럼 포함
  2) rule insert/update 시 note 컬럼이 있으면 같이 저장 (없어도 에러 안 나게)
"""

from __future__ import annotations

import sys
import sqlite3
from pathlib import Path
from typing import Optional, Dict, Any, List

# 데이터베이스 파일 경로
# 실행 파일 또는 스크립트 위치 기준 경로 설정
if getattr(sys, 'frozen', False):
    # PyInstaller로 패키징된 경우
    # --add-data로 추가된 파일은 _internal 폴더에 있음
    if hasattr(sys, '_MEIPASS'):
        # _internal 폴더 경로 (--add-data로 추가된 파일 위치)
        base_path = Path(sys._MEIPASS)
    else:
        # 실행 파일과 같은 폴더
        base_path = Path(sys.executable).parent
else:
    # 개발 환경
    base_path = Path(__file__).parent.parent

DB_PATH = base_path / "data" / "TestDB.sqlite"

# 디폴트 보증값(전역 warranty 행이 없을 때 fallback)
DEFAULT_WARRANTY_MILEAGE = 60000
DEFAULT_WARRANTY_PERIOD_YEARS = 3


def init_database() -> None:
    """
    데이터베이스 초기화
    - data 폴더 생성
    - sap / warranty(전역 1행) / common_project_liability 테이블이 없으면 생성
    - warranty는 id=1 한 줄이 항상 존재하도록 보장
    """
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    # sap: 회사 기본정보만 (보증 컬럼 없음)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sap (
            sap_code        TEXT PRIMARY KEY,
            sap_name        TEXT,
            rule_table_name TEXT,
            renault_code    TEXT,
            remark          TEXT DEFAULT '',
            created_at      TEXT DEFAULT (DATETIME('now', 'localtime')),
            updated_at      TEXT DEFAULT (DATETIME('now', 'localtime'))
        )
    """)

    # warranty: 전역 1행 (id=1)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS warranty (
            id               INTEGER PRIMARY KEY CHECK (id = 1),
            warranty_mileage INTEGER NOT NULL,
            warranty_period  INTEGER NOT NULL,  -- years
            created_at       TEXT DEFAULT (DATETIME('now', 'localtime')),
            updated_at       TEXT DEFAULT (DATETIME('now', 'localtime'))
        )
    """)

    # ✅ common_project_liability: 프로젝트별 기본(공통) 구상률 (1프로젝트 1행)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS common_project_liability (
            project_code    TEXT PRIMARY KEY,   -- L43, H45, LFD, HZG, ALL 등
            liability_ratio REAL NOT NULL       -- 0~1 스케일 권장(너가 이미 그렇게 쓰는 걸로)
        )
    """)

    # 전역 warranty 1행 보장
    cursor.execute("""
        INSERT OR IGNORE INTO warranty (id, warranty_mileage, warranty_period)
        VALUES (1, ?, ?)
    """, (DEFAULT_WARRANTY_MILEAGE, DEFAULT_WARRANTY_PERIOD_YEARS))

    conn.commit()
    conn.close()


def _get_global_warranty() -> tuple[int, int]:
    """
    전역 warranty(1행) 읽기
    없으면 DEFAULT 반환
    """
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("""
        SELECT warranty_mileage, warranty_period
        FROM warranty
        WHERE id = 1
        LIMIT 1
    """)
    row = cur.fetchone()
    conn.close()

    if not row:
        return DEFAULT_WARRANTY_MILEAGE, DEFAULT_WARRANTY_PERIOD_YEARS

    mileage = row["warranty_mileage"]
    period = row["warranty_period"]

    try:
        mileage_i = int(mileage)
    except Exception:
        mileage_i = DEFAULT_WARRANTY_MILEAGE

    try:
        period_i = int(period)
    except Exception:
        period_i = DEFAULT_WARRANTY_PERIOD_YEARS

    return mileage_i, period_i


def get_global_warranty() -> tuple[int, int]:
    """
    전역 warranty(1행) 읽기 (public 함수)
    없으면 DEFAULT 반환
    """
    return _get_global_warranty()


def update_global_warranty(warranty_mileage: int, warranty_period_years: int) -> None:
    """
    전역 warranty(1행) 업데이트
    """
    wm = int(warranty_mileage)
    wp = int(warranty_period_years)

    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO warranty (id, warranty_mileage, warranty_period, updated_at)
        VALUES (1, ?, ?, DATETIME('now','localtime'))
        ON CONFLICT(id) DO UPDATE SET
            warranty_mileage = excluded.warranty_mileage,
            warranty_period  = excluded.warranty_period,
            updated_at       = excluded.updated_at
    """, (wm, wp))

    conn.commit()
    conn.close()


# =========================================================
# common_project_liability (프로젝트 기본/공통 구상률) CRUD
# =========================================================

def get_common_project_liability(project_code: str) -> Optional[float]:
    """
    project_code의 기본(공통) 구상률 조회
    없으면 None
    """
    if not project_code:
        return None

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("""
        SELECT liability_ratio
        FROM common_project_liability
        WHERE project_code = ?
        LIMIT 1
    """, (project_code,))

    row = cur.fetchone()
    conn.close()

    if not row:
        return None

    try:
        return float(row["liability_ratio"])
    except Exception:
        return None


def upsert_common_project_liability(project_code: str, liability_ratio: float) -> None:
    """
    project_code의 기본(공통) 구상률 저장/수정
    """
    if not project_code:
        raise ValueError("project_code is required")

    lr = float(liability_ratio)

    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO common_project_liability (project_code, liability_ratio)
        VALUES (?, ?)
        ON CONFLICT(project_code) DO UPDATE SET
            liability_ratio = excluded.liability_ratio
    """, (project_code, lr))

    conn.commit()
    conn.close()


def get_all_common_project_liabilities() -> List[Dict[str, Any]]:
    """
    프로젝트 기본(공통) 구상률 전체 조회 (표/드롭다운 용)
    """
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("""
        SELECT project_code, liability_ratio
        FROM common_project_liability
        ORDER BY project_code
    """)
    rows = cur.fetchall()
    conn.close()

    return [dict(r) for r in rows] if rows else []


def delete_common_project_liability(project_code: str) -> bool:
    """
    프로젝트 기본(공통) 구상률 삭제
    """
    if not project_code:
        return False

    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM common_project_liability
        WHERE project_code = ?
    """, (project_code,))

    conn.commit()
    ok = cur.rowcount > 0
    conn.close()
    return ok


def get_default_liability_ratio_for_project(project_code: str) -> Optional[float]:
    """
    rule_* 매칭 실패 시 사용할 프로젝트 기본(공통) 구상률
    - project_code가 있으면 그걸 먼저 찾고
    - 없으면 ALL을 fallback
    """
    if not project_code:
        return get_common_project_liability("ALL")

    v = get_common_project_liability(project_code)
    if v is not None:
        return v

    return get_common_project_liability("ALL")


# =========================================================
# SAP 기업정보
# =========================================================

def get_company_info(sap_code_or_name: str) -> Optional[Dict[str, Any]]:
    """
    SAP 기업정보 조회 (sap_code 또는 sap_name으로 조회 가능)

    Returns:
        기업정보 딕셔너리 (기존 코드 호환을 위해 필드명 유지)
        - mileage_threshold
        - warranty_years
        - warranty_mileage
        - warranty_period
    """
    # 전역 warranty 읽기
    warranty_mileage, warranty_years = _get_global_warranty()

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM sap
        WHERE sap_code = ? OR sap_name = ?
        LIMIT 1
    """, (sap_code_or_name, sap_code_or_name))

    row = cursor.fetchone()
    conn.close()

    if not row:
        return None

    data = dict(row)

    return {
        "sap_code": data.get("sap_code"),
        "sap_name": data.get("sap_name"),
        "company_name": data.get("sap_name"),  # 호환성
        "mileage_threshold": warranty_mileage,
        "warranty_years": warranty_years,
        "warranty_mileage": warranty_mileage,
        "warranty_period": warranty_years,  # years
        "rule_table_name": data.get("rule_table_name"),
        "remark": data.get("remark", ""),
        "renault_code": data.get("renault_code", ""),
        # GUI 호환용(기존 코드에서 기대하면 유지)
        "sheet_index": 0,
        "header_row": 3,
        "data_start_row": 4,
    }


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

    return [{"sap_code": r["sap_code"], "sap_name": r["sap_name"]} for r in rows] if rows else []


# =========================================================
# rule 테이블
# =========================================================

def _table_has_column(cursor: sqlite3.Cursor, table_name: str, column_name: str) -> bool:
    """
    테이블에 특정 컬럼이 있는지 확인
    - note 컬럼 유무에 따라 INSERT/UPDATE 구문을 바꿔서, 이미 만들어진 테이블이 달라도 에러 안 나게 함.
    """
    try:
        cursor.execute(f'PRAGMA table_info("{table_name}")')
        cols = [r[1] for r in cursor.fetchall()]
        return column_name in cols
    except Exception:
        return False


def get_rules_from_table(rule_table_name: str) -> List[Dict[str, Any]]:
    """
    rule_table_name에 해당하는 테이블에서 모든 규칙 조회
    - priority 오름차순 정렬
    """
    if not rule_table_name:
        return []
    if not rule_table_name.startswith("rule_"):
        return []

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    try:
        cursor.execute(f"""
            SELECT * FROM "{rule_table_name}"
            ORDER BY priority ASC, rule_id ASC
        """)
        rows = cursor.fetchall()
        conn.close()
        return [dict(r) for r in rows] if rows else []
    except sqlite3.OperationalError:
        conn.close()
        return []


def create_rule_table(rule_table_name: str, cursor=None) -> bool:
    """
    룰 테이블 생성

    ✅ note 컬럼 포함 (신규 생성되는 rule_* 테이블에 note가 빠지지 않게)
    """
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")

    use_existing_cursor = cursor is not None
    if not use_existing_cursor:
        conn = sqlite3.connect(str(DB_PATH))
        cursor = conn.cursor()

    try:
        cursor.execute("""
            SELECT name FROM sqlite_master
            WHERE type='table' AND name=?
        """, (rule_table_name,))

        if cursor.fetchone():
            if not use_existing_cursor:
                conn.close()
            return True

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
                liability_ratio REAL,
                amount_cap_type TEXT NOT NULL DEFAULT 'NONE'
                    CHECK (amount_cap_type IN ('LABOR','OUTSOURCE_LABOR','BOTH_LABOR','NONE')),
                amount_cap_value INTEGER,

                -- ✅ note 컬럼 추가(신규 테이블 생성 시)
                note TEXT NOT NULL DEFAULT '',

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


def upsert_company(
    sap_code: str,
    sap_name: str = None,
    rule_table_name: str = None,
    renault_code: str = None,
) -> None:
    """
    SAP 기업정보 저장/업데이트 (보증 제외)

    - sap: sap_code, sap_name, rule_table_name, remark, renault_code 등 "기본 정보"
    - warranty는 전역 1행이라 여기서 건드리지 않음
    """
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cursor.execute("SELECT 1 FROM sap WHERE sap_code = ?", (sap_code,))
        exists = cursor.fetchone() is not None

        if exists:
            updates = []
            values = []

            if sap_name is not None:
                updates.append("sap_name = ?")
                values.append(sap_name)
            if rule_table_name is not None:
                updates.append("rule_table_name = ?")
                values.append(rule_table_name)
            if renault_code is not None:
                updates.append("renault_code = ?")
                values.append(renault_code)

            if updates:
                updates.append("updated_at = DATETIME('now', 'localtime')")
                values.append(sap_code)
                cursor.execute(
                    f"UPDATE sap SET {', '.join(updates)} WHERE sap_code = ?",
                    values,
                )

            # 룰 테이블 생성(원래 로직 유지)
            if rule_table_name:
                try:
                    create_rule_table(rule_table_name, cursor)
                except Exception:
                    pass

        else:
            cursor.execute(
                """
                INSERT INTO sap (sap_code, sap_name, rule_table_name, renault_code)
                VALUES (?, ?, ?, ?)
                """,
                (sap_code, sap_name, rule_table_name, renault_code),
            )

            if rule_table_name:
                create_rule_table(rule_table_name, cursor)

        conn.commit()

    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def update_company_remark(sap_code: str, remark: str) -> bool:
    """SAP 기업의 remark 업데이트"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cursor.execute("""
            UPDATE sap
            SET remark = ?, updated_at = DATETIME('now', 'localtime')
            WHERE sap_code = ?
        """, (remark, sap_code))

        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        raise ValueError(f"Remark 업데이트 실패: {str(e)}")
    finally:
        conn.close()


def add_rule_to_table(
    rule_table_name: str,
    status: str,
    repair_region: str,
    vehicle_classification: str,
    amount_cap_type: str,
    liability_ratio: float = None,
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
    priority: int = None,
    note: str = "",  # ✅ note 추가
) -> int:
    """rule 테이블에 규칙 추가"""
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        if not repair_region:
            raise ValueError("수리 지역은 필수입니다.")
        if repair_region not in ["DOMESTIC", "OVERSEAS", "ALL"]:
            raise ValueError("수리 지역은 DOMESTIC, OVERSEAS, ALL 중 하나여야 합니다.")

        if not vehicle_classification:
            vehicle_classification = "ALL"

        if liability_ratio is None:
            if amount_cap_type in ["LABOR", "OUTSOURCE_LABOR", "BOTH_LABOR"] and amount_cap_value is not None:
                pass
            else:
                raise ValueError("구상율은 필수입니다. (LABOR 최댓값 규칙이 아닌 경우)")

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

        if priority is None:
            cursor.execute(f'SELECT MAX(priority) FROM "{rule_table_name}"')
            max_priority = cursor.fetchone()[0]
            priority = 1 if max_priority is None else (max_priority + 1)

        if valid_from and valid_from.strip():
            from datetime import datetime
            datetime.strptime(valid_from.strip(), "%Y-%m-%d")
        if valid_to and valid_to.strip():
            from datetime import datetime
            datetime.strptime(valid_to.strip(), "%Y-%m-%d")

        has_note = _table_has_column(cursor, rule_table_name, "note")

        if has_note:
            cursor.execute(f"""
                INSERT INTO "{rule_table_name}" (
                    priority, status, repair_region, project_code, exclude_project_code,
                    vehicle_classification, part_no, part_name, engine_form,
                    warranty_mileage_override, warranty_period_override,
                    liability_ratio, amount_cap_type, amount_cap_value,
                    note,
                    valid_from, valid_to,
                    created_at, updated_at
                ) VALUES (
                    ?, ?, ?, ?, ?,
                    ?, ?, ?, ?,
                    ?, ?,
                    ?, ?, ?,
                    ?,
                    ?, ?,
                    DATETIME('now', 'localtime'), DATETIME('now', 'localtime')
                )
            """, (
                priority, status, repair_region, project_code, exclude_project_code,
                vehicle_classification, part_no, part_name, engine_form,
                warranty_mileage_override, warranty_period_override,
                liability_ratio, amount_cap_type, amount_cap_value,
                (note or ""),
                valid_from, valid_to,
            ))
        else:
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
        return rule_id
    except sqlite3.OperationalError as e:
        raise ValueError(f"Rule 추가 실패: {str(e)}")
    finally:
        conn.close()


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
    note: str = None,  # ✅ note 수정 지원
) -> bool:
    """rule 테이블의 규칙 수정"""
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

        # ✅ note는 테이블에 있을 때만 업데이트 (없으면 무시)
        if note is not None and _table_has_column(cursor, rule_table_name, "note"):
            updates.append("note = ?")
            values.append(note)

        if not updates:
            return False

        updates.append("updated_at = DATETIME('now', 'localtime')")
        values.append(rule_id)

        cursor.execute(f"""
            UPDATE "{rule_table_name}"
            SET {", ".join(updates)}
            WHERE rule_id = ?
        """, values)

        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        raise ValueError(f"Rule 수정 실패: {str(e)}")
    finally:
        conn.close()


def update_rule_priorities(rule_table_name: str, rule_ids_in_order: List[int]) -> bool:
    """드래그 앤 드롭으로 변경된 순서에 따라 priority 재할당"""
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")

    if not rule_ids_in_order:
        return True

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        for new_priority, rule_id in enumerate(rule_ids_in_order, start=1):
            cursor.execute(f'''
                UPDATE "{rule_table_name}"
                SET priority = ?, updated_at = DATETIME('now', 'localtime')
                WHERE rule_id = ?
            ''', (new_priority, rule_id))

        conn.commit()
        return True
    except sqlite3.Error as e:
        conn.rollback()
        raise ValueError(f"우선순위 업데이트 실패: {str(e)}")
    finally:
        conn.close()


def delete_rule_from_table(rule_table_name: str, rule_id: int) -> bool:
    """rule 테이블에서 규칙 삭제"""
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
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        raise ValueError(f"Rule 삭제 실패: {str(e)}")
    finally:
        conn.close()
