"""
SQLite 데이터베이스 관리 모듈
SAP 기업정보 저장 및 조회

요구사항 반영(중요):
- sap 테이블에서 warranty_mileage / warranty_period 컬럼 제거(미사용/없다고 가정)
- 보증정보는 warranty 테이블(전역 1행, id=1)로 분리하여 관리
- get_company_info는 warranty(전역값) 우선으로 보증값을 반환하고, 없으면 기본값(60000/3) 사용

추가:
- common_project_liability 테이블 추가 (project_code -> 기본 liability_ratio)

변경:
- rule_{sap_code} 테이블에 note 컬럼 지원
  1) 새로 생성되는 rule 테이블 스키마에 note 컬럼 포함
  2) rule insert/update 시 note 컬럼이 있으면 같이 저장 (없어도 에러 안 나게)
  3) 기존 rule 테이블에도 note 컬럼 없으면 자동 추가(마이그레이션)

변경(차계 프로젝트 맵핑 DB화):
- vehicle_project_map 테이블 (id, vehicle_prefix, project_code, created_at, updated_at)
- get_project_code_from_vehicle_db() 제공 (전처리에서 사용)
- init_database()에서는 절대 seed 데이터 삽입하지 않음
"""

from __future__ import annotations

import sys
import sqlite3
import re
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple


# =========================================================
# DB 경로
# =========================================================

if getattr(sys, "frozen", False):
    if hasattr(sys, "_MEIPASS"):
        base_path = Path(sys._MEIPASS)
    else:
        base_path = Path(sys.executable).parent
else:
    base_path = Path(__file__).parent.parent

DB_PATH = base_path / "data" / "TestDB.sqlite"

DEFAULT_WARRANTY_MILEAGE = 60000
DEFAULT_WARRANTY_PERIOD_YEARS = 3


# =========================================================
# 공통 유틸
# =========================================================

def _table_has_column(cursor: sqlite3.Cursor, table_name: str, column_name: str) -> bool:
    try:
        cursor.execute(f'PRAGMA table_info("{table_name}")')
        cols = [r[1] for r in cursor.fetchall()]
        return column_name in cols
    except Exception:
        return False


def _table_exists(cursor: sqlite3.Cursor, table_name: str) -> bool:
    cursor.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
        (table_name,),
    )
    return cursor.fetchone() is not None


def _index_exists(cursor: sqlite3.Cursor, index_name: str) -> bool:
    cursor.execute(
        "SELECT 1 FROM sqlite_master WHERE type='index' AND name=? LIMIT 1",
        (index_name,),
    )
    return cursor.fetchone() is not None


# =========================================================
# rule 테이블 note 컬럼 마이그레이션
# =========================================================

def ensure_note_column_for_all_rule_tables() -> None:
    """
    기존에 생성된 rule_* 테이블에 note 컬럼이 없으면 추가한다.
    note: TEXT NOT NULL DEFAULT ''
    """
    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT name
            FROM sqlite_master
            WHERE type='table' AND name LIKE 'rule_%'
        """)
        tables = [r[0] for r in cur.fetchall()]

        for t in tables:
            if not _table_has_column(cur, t, "note"):
                cur.execute(f'ALTER TABLE "{t}" ADD COLUMN note TEXT NOT NULL DEFAULT ""')

        conn.commit()
    finally:
        conn.close()


# =========================================================
# vehicle_project_map 마이그레이션 (vehicle_key -> vehicle_prefix)
# =========================================================

def ensure_vehicle_project_map_schema() -> None:
    """
    vehicle_project_map 스키마를 요구사항으로 강제 정렬.

    - 기존 테이블이 없으면 아무 것도 안 함 (init_database가 생성)
    - 기존 테이블이 있으면:
      * vehicle_key / vehicle_prefix 어떤 형태든 읽어서
      * 새 스키마(id, vehicle_prefix, project_code, created_at, updated_at)로 복제
      * old -> __old로 rename, new -> 원래 이름으로 rename, old drop
    - seed 삽입 없음
    """
    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()
    try:
        if not _table_exists(cur, "vehicle_project_map"):
            return

        # 0) 기존 인덱스 정리(테이블 rename 전에 먼저)
        cur.execute("DROP INDEX IF EXISTS ux_vpm_vehicle_key")
        cur.execute("DROP INDEX IF EXISTS idx_vpm_vehicle_key")
        cur.execute("DROP INDEX IF EXISTS ux_vpm_vehicle_prefix")
        cur.execute("DROP INDEX IF EXISTS idx_vpm_vehicle_prefix")
        cur.execute("DROP INDEX IF EXISTS idx_vpm_project_code")

        # 1) 현재 컬럼 목록
        cur.execute('PRAGMA table_info("vehicle_project_map")')
        cols = [r[1] for r in cur.fetchall()]

        has_vehicle_key = "vehicle_key" in cols
        has_vehicle_prefix = "vehicle_prefix" in cols

        # 2) 데이터 읽기: (vehicle_prefix, project_code)
        if has_vehicle_prefix:
            cur.execute("SELECT vehicle_prefix, project_code FROM vehicle_project_map")
        elif has_vehicle_key:
            cur.execute("SELECT vehicle_key AS vehicle_prefix, project_code FROM vehicle_project_map")
        else:
            cur.execute("SELECT NULL AS vehicle_prefix, NULL AS project_code WHERE 0")

        rows = cur.fetchall()

        # 3) 새 테이블 생성(요구사항 스키마)
        #    - SQLite DEFAULT 함수 문제 피하려고 DEFAULT 안 넣음(필요하면 UPDATE로 채움)
        cur.execute("""
            CREATE TABLE vehicle_project_map__new (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                vehicle_prefix TEXT NOT NULL,
                project_code   TEXT NOT NULL,
                created_at     TEXT,
                updated_at     TEXT
            )
        """)

        # 4) 데이터 이관
        for vp, pc in rows:
            if vp is None or pc is None:
                continue
            vp2 = str(vp).strip().upper()
            pc2 = str(pc).strip().upper()
            if not vp2 or not pc2:
                continue
            cur.execute("""
                INSERT INTO vehicle_project_map__new (vehicle_prefix, project_code, created_at, updated_at)
                VALUES (?, ?, DATETIME('now','localtime'), DATETIME('now','localtime'))
            """, (vp2, pc2))

        # 5) 새 테이블 내 중복 vehicle_prefix 정리 (같은 prefix 여러개면 마지막 것만 남김)
        cur.execute("""
            DELETE FROM vehicle_project_map__new
            WHERE id NOT IN (
                SELECT MAX(id)
                FROM vehicle_project_map__new
                GROUP BY vehicle_prefix
            )
        """)

        # 6) 기존 테이블 rename + 새 테이블을 원래 이름으로
        cur.execute("ALTER TABLE vehicle_project_map RENAME TO vehicle_project_map__old")
        cur.execute("ALTER TABLE vehicle_project_map__new RENAME TO vehicle_project_map")

        # 7) 인덱스 재생성
        cur.execute("""
            CREATE UNIQUE INDEX ux_vpm_vehicle_prefix
            ON vehicle_project_map(vehicle_prefix)
        """)
        cur.execute("""
            CREATE INDEX idx_vpm_vehicle_prefix
            ON vehicle_project_map(vehicle_prefix)
        """)
        cur.execute("""
            CREATE INDEX idx_vpm_project_code
            ON vehicle_project_map(project_code)
        """)

        # 8) old 테이블 제거
        cur.execute("DROP TABLE IF EXISTS vehicle_project_map__old")

        conn.commit()
    finally:
        conn.close()



# =========================================================
# DB 초기화
# =========================================================

def init_database() -> None:
    """
    데이터베이스 초기화
    - data 폴더 생성
    - sap / warranty(전역 1행) / common_project_liability / vehicle_project_map 테이블이 없으면 생성
    - warranty는 id=1 한 줄이 항상 존재하도록 보장
    - ✅ vehicle_project_map seed 삽입 금지
    - 기존 rule_* 테이블 note 컬럼 마이그레이션
    - vehicle_project_map vehicle_key -> vehicle_prefix 마이그레이션
    """
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

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

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS warranty (
            id               INTEGER PRIMARY KEY CHECK (id = 1),
            warranty_mileage INTEGER NOT NULL,
            warranty_period  INTEGER NOT NULL,  -- years
            created_at       TEXT DEFAULT (DATETIME('now', 'localtime')),
            updated_at       TEXT DEFAULT (DATETIME('now', 'localtime'))
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS common_project_liability (
            project_code    TEXT PRIMARY KEY,   -- L43, H45, LFD, HZG, ALL 등
            liability_ratio REAL NOT NULL       -- 0~1 스케일
        )
    """)

    # ✅ 요구사항 스키마로 생성(없을 때만)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS vehicle_project_map (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            vehicle_prefix TEXT NOT NULL,
            project_code   TEXT NOT NULL,
            created_at     TEXT DEFAULT (DATETIME('now', 'localtime')),
            updated_at     TEXT DEFAULT (DATETIME('now', 'localtime'))
        )
    """)

    cursor.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS ux_vpm_vehicle_prefix
        ON vehicle_project_map(vehicle_prefix)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_vpm_vehicle_prefix
        ON vehicle_project_map(vehicle_prefix)
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_vpm_project_code
        ON vehicle_project_map(project_code)
    """)

    # warranty 1행 보장
    cursor.execute("""
        INSERT OR IGNORE INTO warranty (id, warranty_mileage, warranty_period)
        VALUES (1, ?, ?)
    """, (DEFAULT_WARRANTY_MILEAGE, DEFAULT_WARRANTY_PERIOD_YEARS))

    # ✅ vehicle_project_map seed 절대 넣지 않음

    conn.commit()
    conn.close()

    # 마이그레이션들(테이블이 이미 있던 경우 포함)
    ensure_note_column_for_all_rule_tables()
    ensure_vehicle_project_map_schema()


# =========================================================
# warranty (전역)
# =========================================================

def _get_global_warranty() -> Tuple[int, int]:
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


def get_global_warranty() -> Tuple[int, int]:
    return _get_global_warranty()


def update_global_warranty(warranty_mileage: int, warranty_period_years: int) -> None:
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
# vehicle_project_map (차계/문자열 -> 프로젝트 코드)
# =========================================================


def get_project_code_from_vehicle_db(vehicle: Any) -> Optional[str]:
    """
    vehicle_project_map에서 project_code 조회

    조회 규칙:
    - vehicle 값에서 '첫 영문자 + 뒤 숫자들'을 뽑아 정규화 키 생성
      예) "G417" / "G-417" / "g 417" -> "G417"
          "G" -> "G"
    - DB에는 "G417" 같은 상세키도, "G" 같은 prefix도 있을 수 있음
    - 가장 구체적인 매칭(길이가 긴 vehicle_prefix) 우선
    """
    v = (str(vehicle) if vehicle is not None else "").strip().upper()
    if not v:
        return None

    # 0) vehicle 자체가 프로젝트 코드면 그대로 반환(우선)
    if v in ("LFD", "HZG", "LJL", "AR1", "AR2"):
        return v

    # 1) 정규화: 첫 알파벳 + 숫자(있으면)만 붙여서 키 생성
    m = re.search(r"([A-Z])\s*[-_ ]*\s*(\d+)?", v)
    if not m:
        return None

    letter = m.group(1)
    digits = m.group(2) or ""
    key_full = f"{letter}{digits}"  # "G417" or "G"

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute(
            """
            SELECT project_code
            FROM vehicle_project_map
            WHERE ? LIKE UPPER(vehicle_prefix) || '%'
            ORDER BY LENGTH(vehicle_prefix) DESC, id DESC
            LIMIT 1
            """,
            (key_full,),
        )
        row = cur.fetchone()
        if row and row["project_code"]:
            return str(row["project_code"]).strip().upper()
        return None
    finally:
        conn.close()



def upsert_vehicle_project_map(vehicle_prefix: str, project_code: str) -> int:
    """
    upsert:
    - vehicle_prefix(유니크) 기준으로 project_code 갱신
    - 없으면 INSERT
    """
    vp = (vehicle_prefix or "").strip().upper()
    pc = (project_code or "").strip().upper()
    if not vp:
        raise ValueError("vehicle_prefix is required")
    if not pc:
        raise ValueError("project_code is required")

    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO vehicle_project_map(vehicle_prefix, project_code, updated_at)
            VALUES (?, ?, DATETIME('now','localtime'))
            ON CONFLICT(vehicle_prefix) DO UPDATE SET
                project_code = excluded.project_code,
                updated_at   = excluded.updated_at
        """, (vp, pc))
        conn.commit()

        cur.execute("SELECT id FROM vehicle_project_map WHERE UPPER(vehicle_prefix)=? LIMIT 1", (vp,))
        row = cur.fetchone()
        return int(row[0]) if row else -1
    finally:
        conn.close()


def get_all_vehicle_project_maps() -> List[Dict[str, Any]]:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT id, vehicle_prefix, project_code, created_at, updated_at
            FROM vehicle_project_map
            ORDER BY id ASC
        """)
        rows = cur.fetchall()
        return [dict(r) for r in rows] if rows else []
    finally:
        conn.close()


def delete_vehicle_project_map(id_: int) -> bool:
    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM vehicle_project_map WHERE id = ?", (int(id_),))
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


# =========================================================
# common_project_liability (프로젝트 기본/공통 구상률) CRUD
# =========================================================

def get_common_project_liability(project_code: str) -> Optional[float]:
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
        "company_name": data.get("sap_name"),
        "mileage_threshold": warranty_mileage,
        "warranty_years": warranty_years,
        "warranty_mileage": warranty_mileage,
        "warranty_period": warranty_years,
        "rule_table_name": data.get("rule_table_name"),
        "remark": data.get("remark", ""),
        "renault_code": data.get("renault_code", ""),
        "sheet_index": 0,
        "header_row": 3,
        "data_start_row": 4,
    }


def get_all_companies() -> List[str]:
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    cursor.execute("SELECT sap_name FROM sap ORDER BY sap_name")
    rows = cursor.fetchall()
    conn.close()

    return [row[0] for row in rows] if rows else []


def get_all_companies_with_code() -> List[Dict[str, str]]:
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

def get_rules_from_table(rule_table_name: str) -> List[Dict[str, Any]]:
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


def update_company(
    old_sap_code: str,
    new_sap_code: str = None,
    sap_name: str = None,
    renault_code: str = None,
) -> bool:
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cursor.execute("SELECT rule_table_name FROM sap WHERE sap_code = ?", (old_sap_code,))
        row = cursor.fetchone()
        if not row:
            raise ValueError(f"협력사를 찾을 수 없습니다: {old_sap_code}")

        old_rule_table_name = row[0]

        if new_sap_code and new_sap_code != old_sap_code:
            new_rule_table_name = f"rule_{new_sap_code}"

            if old_rule_table_name:
                cursor.execute("""
                    SELECT name FROM sqlite_master
                    WHERE type='table' AND name = ?
                """, (old_rule_table_name,))
                if cursor.fetchone():
                    cursor.execute(f'ALTER TABLE "{old_rule_table_name}" RENAME TO "{new_rule_table_name}"')

            updates = []
            values = []

            updates.append("sap_code = ?")
            values.append(new_sap_code)

            updates.append("rule_table_name = ?")
            values.append(new_rule_table_name if old_rule_table_name else None)

            if sap_name is not None:
                updates.append("sap_name = ?")
                values.append(sap_name)

            if renault_code is not None:
                updates.append("renault_code = ?")
                values.append(renault_code)

            updates.append("updated_at = DATETIME('now', 'localtime')")
            values.append(old_sap_code)

            cursor.execute(
                f"UPDATE sap SET {', '.join(updates)} WHERE sap_code = ?",
                values
            )
        else:
            updates = []
            values = []

            if sap_name is not None:
                updates.append("sap_name = ?")
                values.append(sap_name)

            if renault_code is not None:
                updates.append("renault_code = ?")
                values.append(renault_code)

            if updates:
                updates.append("updated_at = DATETIME('now', 'localtime')")
                values.append(old_sap_code)

                cursor.execute(
                    f"UPDATE sap SET {', '.join(updates)} WHERE sap_code = ?",
                    values
                )

        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        conn.rollback()
        raise ValueError(f"협력사 업데이트 실패: {str(e)}")
    finally:
        conn.close()


def delete_company(sap_code: str) -> bool:
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cursor.execute("SELECT rule_table_name FROM sap WHERE sap_code = ?", (sap_code,))
        row = cursor.fetchone()
        if not row:
            return False

        rule_table_name = row[0]

        if rule_table_name:
            cursor.execute(f'DROP TABLE IF EXISTS "{rule_table_name}"')

        cursor.execute("DELETE FROM sap WHERE sap_code = ?", (sap_code,))

        conn.commit()
        return cursor.rowcount > 0
    except sqlite3.OperationalError as e:
        conn.rollback()
        raise ValueError(f"협력사 삭제 실패: {str(e)}")
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
    note: str = "",
) -> int:
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
            has_warranty_override = warranty_mileage_override is not None or warranty_period_override is not None
            has_amount_cap = amount_cap_type in ["LABOR", "OUTSOURCE_LABOR", "BOTH_LABOR"] and amount_cap_value is not None
            if not (has_warranty_override or has_amount_cap):
                raise ValueError("구상율은 필수입니다. (워런티 오버라이드 또는 공임비 상한 규칙이 아닌 경우)")

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
    note: str = None,
) -> bool:
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
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")

    if not rule_ids_in_order:
        return True

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        for new_priority, rule_id in enumerate(rule_ids_in_order, start=1):
            cursor.execute(f"""
                UPDATE "{rule_table_name}"
                SET priority = ?, updated_at = DATETIME('now', 'localtime')
                WHERE rule_id = ?
            """, (new_priority, rule_id))

        conn.commit()
        return True
    except sqlite3.Error as e:
        conn.rollback()
        raise ValueError(f"우선순위 업데이트 실패: {str(e)}")
    finally:
        conn.close()


def delete_rule_from_table(rule_table_name: str, rule_id: int) -> bool:
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
