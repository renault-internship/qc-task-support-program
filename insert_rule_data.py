import re
import sqlite3
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

DB_PATH = Path("data/TestDB.sqlite")

# 1) 기존 레거시(6필드)
LegacyRuleRow = Tuple[str, str, str, float, Optional[str], Optional[str]]

# 2) CSV/신규(15필드)
# (repair_region, project_code, exclude_project_code, vehicle_classification,
#  part_no, part_name, engine_form, mileage_cap, period_cap,
#  liability_ratio, cap_type, cap_value, note, valid_from, valid_to)
FullRuleRow = Tuple[
    str, str, Optional[str], str,
    str, str, str, Optional[int], Optional[int],
    Optional[float], str, Optional[int], Optional[str],
    Optional[str], Optional[str]
]


def safe_table_name(name: str) -> str:
    if not re.fullmatch(r"rule_[A-Za-z0-9_]+", name):
        raise ValueError(f"Invalid table name: {name}")
    return name


def _strip_smart_quotes(s: str) -> str:
    return s.replace("‘", "").replace("’", "").replace('"', "").strip()


def normalize_nullable_date(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return None
    return s


def normalize_nullable_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return None
    return s


def normalize_text(v: Any, default: str) -> str:
    if v is None:
        return default
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return default
    if s.lower() == "all":
        return "ALL"
    return s


def normalize_nullable_int(v: Any) -> Optional[int]:
    if v is None:
        return None
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def normalize_cap_type(v: Any) -> str:
    # DB CHECK: ('LABOR','OUTSOURCE_LABOR','BOTH_LABOR','NONE')
    if v is None:
        return "NONE"
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return "NONE"
    s_up = s.upper()

    # 흔한 표현들 흡수
    if s_up in {"NONE", "NO", "N/A", "NA"}:
        return "NONE"
    if s_up in {"LABOR"}:
        return "LABOR"
    if s_up in {"OUTSOURCE_LABOR", "OUTSOURCELABOR", "OUTSOURCE"}:
        return "OUTSOURCE_LABOR"
    if s_up in {"BOTH_LABOR", "BOTH", "LABOR+OUTSOURCE", "BOTH_LABOUR"}:
        return "BOTH_LABOR"

    # 모르는 값이면 DB 제약에 걸리니 NONE으로 강제
    return "NONE"


def normalize_liability_ratio(v: Any) -> Optional[float]:
    if v is None:
        return None
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return None

    # "50%" 같은 케이스
    s = s.replace("%", "").strip()

    try:
        lr = float(s)
    except Exception:
        return None

    # 50 -> 0.5
    if lr > 1.0:
        lr = lr / 100.0

    return lr


def validate_row(row: tuple) -> FullRuleRow:
    """
    6필드 레거시도 받고, 15필드 신규도 받는다.
    반환은 항상 FullRuleRow(15필드)로 정규화.
    """
    if len(row) == 6:
        repair_region, project_code, part_name, liability_ratio, valid_from, valid_to = row

        rr = normalize_text(repair_region, "ALL")
        pc = normalize_text(project_code, "ALL")
        epc = None
        vc = "ALL"
        part_no = "ALL"
        pn = normalize_text(part_name, "ALL")
        eng = "ALL"
        mileage_cap = None
        period_cap = None
        lr = normalize_liability_ratio(liability_ratio)
        if lr is None:
            raise ValueError(f"liability_ratio is required. row={row}")

        cap_type = "NONE"
        cap_value = None
        note = None
        vf = normalize_nullable_date(valid_from)
        vt = normalize_nullable_date(valid_to)

        return (
            rr, pc, epc, vc,
            part_no, pn, eng, mileage_cap, period_cap,
            lr, cap_type, cap_value, note,
            vf, vt
        )

    if len(row) == 15:
        (
            repair_region, project_code, exclude_project_code, vehicle_classification,
            part_no, part_name, engine_form, mileage_cap, period_cap,
            liability_ratio, cap_type, cap_value, note, valid_from, valid_to
        ) = row

        rr = normalize_text(repair_region, "ALL")
        pc = normalize_text(project_code, "ALL")
        epc = normalize_nullable_text(exclude_project_code)  # nullable
        vc = normalize_text(vehicle_classification, "ALL")
        pno = normalize_text(part_no, "ALL")
        pname = normalize_text(part_name, "ALL")
        eng = normalize_text(engine_form, "ALL")

        mcap = normalize_nullable_int(mileage_cap)
        pcap = normalize_nullable_int(period_cap)

        lr = normalize_liability_ratio(liability_ratio)
        # amount_cap_type에 따라서 lr이 None일 수도 있지만,
        # 네 규칙상 대부분 lr은 들어오니 여기선 None이면 에러로 둔다.
        if lr is None:
            raise ValueError(f"liability_ratio is required. row={row}")

        ct = normalize_cap_type(cap_type)
        cv = normalize_nullable_int(cap_value)

        nt = normalize_nullable_text(note)
        vf = normalize_nullable_date(valid_from)
        vt = normalize_nullable_date(valid_to)

        return (
            rr, pc, epc, vc,
            pno, pname, eng, mcap, pcap,
            lr, ct, cv, nt,
            vf, vt
        )

    raise ValueError(
        f"Row must have 6 fields(legacy) or 15 fields(full). "
        f"Got len={len(row)} row={row}"
    )


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    cur = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    )
    return cur.fetchone() is not None


def get_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
    cur = conn.execute(f"PRAGMA table_info({table_name})")
    return [r[1] for r in cur.fetchall()]


def ensure_unique_index(conn: sqlite3.Connection, table_name: str) -> None:
    """
    rerun 중복 방지용 NULL-safe unique index
    - IFNULL로 NULL을 센티넬로 치환해서 NULL 때문에 중복 허용되는 문제 방지
    """
    idx_name = f"ux_{table_name}_key"
    conn.execute(f'DROP INDEX IF EXISTS "{idx_name}"')

    conn.execute(
        f"""
        CREATE UNIQUE INDEX "{idx_name}"
        ON "{table_name}"(
            repair_region,
            project_code,
            IFNULL(exclude_project_code, ''),
            vehicle_classification,
            part_no,
            part_name,
            engine_form,
            amount_cap_type,
            IFNULL(amount_cap_value, -1),
            IFNULL(warranty_mileage_override, -1),
            IFNULL(warranty_period_override, -1),
            IFNULL(liability_ratio, -1),
            IFNULL(note, ''),
            IFNULL(valid_from, '0000-00-00'),
            IFNULL(valid_to,   '9999-12-31')
        )
        """
    )


def bulk_insert_rules(rules_to_insert: Dict[str, List[tuple]]):
    if not DB_PATH.exists():
        raise FileNotFoundError(f"DB not found: {DB_PATH.resolve()}")

    with sqlite3.connect(str(DB_PATH)) as conn:
        conn.row_factory = sqlite3.Row

        print(f"[DB] {DB_PATH.resolve()}")
        conn.execute("BEGIN")
        try:
            total_attempted = 0
            total_inserted = 0

            for table_name, rows in rules_to_insert.items():
                if not rows:
                    print(f"[SKIP] {table_name}: rows empty")
                    continue

                table_name = safe_table_name(table_name)

                if not table_exists(conn, table_name):
                    print(f"[SKIP] {table_name}: table not found")
                    continue

                cols = set(get_columns(conn, table_name))

                required = {
                    "repair_region",
                    "project_code",
                    "exclude_project_code",
                    "vehicle_classification",
                    "part_no",
                    "part_name",
                    "engine_form",
                    "warranty_mileage_override",
                    "warranty_period_override",
                    "liability_ratio",
                    "amount_cap_type",
                    "amount_cap_value",
                    "note",
                    "valid_from",
                    "valid_to",
                }
                missing = required - cols
                if missing:
                    print(f"[SKIP] {table_name}: missing columns {sorted(missing)}")
                    continue

                normalized_rows: List[FullRuleRow] = [validate_row(r) for r in rows]

                # ✅ NULL-safe 유니크 인덱스 보장
                ensure_unique_index(conn, table_name)

                before = conn.execute(
                    f'SELECT COUNT(*) AS cnt FROM "{table_name}"'
                ).fetchone()["cnt"]

                # ✅ 중복이면 무시 (unique index가 NULL-safe)
                conn.executemany(
                    f"""
                    INSERT OR IGNORE INTO "{table_name}"
                    (
                      repair_region,
                      project_code,
                      exclude_project_code,
                      vehicle_classification,
                      part_no,
                      part_name,
                      engine_form,
                      warranty_mileage_override,
                      warranty_period_override,
                      liability_ratio,
                      amount_cap_type,
                      amount_cap_value,
                      note,
                      valid_from,
                      valid_to
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    normalized_rows,
                )

                after = conn.execute(
                    f'SELECT COUNT(*) AS cnt FROM "{table_name}"'
                ).fetchone()["cnt"]

                attempted_here = len(normalized_rows)
                inserted_here = after - before

                total_attempted += attempted_here
                total_inserted += inserted_here

                print(f"[OK] {table_name}: attempted {attempted_here}, inserted {inserted_here} (now total rows={after})")

            conn.commit()
            print(f"[DONE] attempted rows = {total_attempted}, actually inserted = {total_inserted}")

        except Exception as e:
            conn.rollback()
            print("[ROLLBACK] error:", repr(e))
            raise


if __name__ == "__main__":
    rules_to_insert = {
        # 레거시 6필드도 그대로 가능
        "rule_Z551": [
            # 너 원본에 ("ALL","ALL","LJL","ALL",0.309) 이건 len=5라서 원래도 에러였음.
            # 아래는 예시로 정상 6필드 형태만 유지해야 함.
            ("ALL", "LJL", "ALL", 0.60, None, None),
        ],

        # 신규 15필드도 가능 (note까지 포함)
        # (repair_region, project_code, exclude_project_code, vehicle_classification,
        #  part_no, part_name, engine_form, mileage_cap, period_cap,
        #  liability_ratio, cap_type, cap_value, note, valid_from, valid_to)
        "rule_A201": [
            ("ALL", "L38", None, "ALL", "ALL", "ALL", "ALL", None, None, 0.60, "NONE", None, "All Items (HVAC etc.)", "2022-06-13", None),
        ],
    }

    bulk_insert_rules(rules_to_insert)
