import re
import sqlite3
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

DB_PATH = Path("data/TestDB.sqlite")

RuleRow = Tuple[str, str, str, float, Optional[str], Optional[str]]


def safe_table_name(name: str) -> str:
    if not re.fullmatch(r"rule_[A-Za-z0-9_]+", name):
        raise ValueError(f"Invalid table name: {name}")
    return name


def normalize_nullable_date(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = str(v).strip()
    if s == "":
        return None
    s = s.replace("‘", "").replace("’", "").replace('"', "").strip()
    return s or None


def normalize_text(v: Any, default: str) -> str:
    if v is None:
        return default
    s = str(v).strip()
    if s == "":
        return default
    s = s.replace("‘", "").replace("’", "").replace('"', "").strip()
    if s.lower() == "all":
        return "ALL"
    return s


def validate_row(row: tuple) -> RuleRow:
    if len(row) != 6:
        raise ValueError(
            f"Row must have 6 fields: "
            f"(repair_region, project_code, part_name, liability_ratio, valid_from, valid_to). "
            f"Got len={len(row)} row={row}"
        )

    repair_region, project_code, part_name, liability_ratio, valid_from, valid_to = row

    rr = normalize_text(repair_region, "ALL")
    pc = normalize_text(project_code, "ALL")
    pn = normalize_text(part_name, "ALL")

    if liability_ratio is None or str(liability_ratio).strip() == "":
        raise ValueError(f"liability_ratio is required. row={row}")

    lr = float(str(liability_ratio).strip().replace("%", ""))
    if lr > 1.0:
        lr = lr / 100.0  # 50 -> 0.5

    vf = normalize_nullable_date(valid_from)
    vt = normalize_nullable_date(valid_to)

    return (rr, pc, pn, lr, vf, vt)


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
    NULL-safe unique index:
    - SQLite UNIQUE index allows multiple rows when any indexed column is NULL.
    - So we normalize NULL(valid_from/valid_to) to sentinel values via IFNULL().
    """
    idx_name = f"ux_{table_name}_key"

    # 기존에 잘못 만들어진 인덱스가 있을 수 있으니 제거 후 재생성
    conn.execute(f"DROP INDEX IF EXISTS {idx_name}")

    conn.execute(
        f"""
        CREATE UNIQUE INDEX {idx_name}
        ON {table_name}(
            repair_region,
            project_code,
            part_name,
            amount_cap_type,
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

                cols = get_columns(conn, table_name)

                required = {
                    "repair_region",
                    "project_code",
                    "part_name",
                    "liability_ratio",
                    "amount_cap_type",
                    "valid_from",
                    "valid_to",
                }
                missing = required - set(cols)
                if missing:
                    print(f"[SKIP] {table_name}: missing columns {sorted(missing)}")
                    continue

                normalized_rows: List[RuleRow] = [validate_row(r) for r in rows]

                # ✅ NULL-safe 유니크 인덱스 보장
                ensure_unique_index(conn, table_name)

                before = conn.execute(f"SELECT COUNT(*) AS cnt FROM {table_name}").fetchone()["cnt"]

                # ✅ 중복이면 무시(인덱스가 NULL-safe라서 None/None도 중복 차단됨)
                conn.executemany(
                    f"""
                    INSERT OR IGNORE INTO {table_name}
                    (
                      repair_region,
                      project_code,
                      part_name,
                      liability_ratio,
                      amount_cap_type,
                      valid_from,
                      valid_to
                    )
                    VALUES (?, ?, ?, ?, 'NONE', ?, ?)
                    """,
                    normalized_rows,
                )

                after = conn.execute(f"SELECT COUNT(*) AS cnt FROM {table_name}").fetchone()["cnt"]

                attempted_here = len(normalized_rows)
                inserted_here = after - before

                total_attempted += attempted_here
                total_inserted += inserted_here

                print(
                    f"[OK] {table_name}: attempted {attempted_here}, inserted {inserted_here} (now total rows={after})"
                )

            conn.commit()
            print(f"[DONE] attempted rows = {total_attempted}, actually inserted = {total_inserted}")

        except Exception as e:
            conn.rollback()
            print("[ROLLBACK] error:", repr(e))
            raise


if __name__ == "__main__":
    rules_to_insert = {
        "rule_Z543": [
            ("ALL", "L38", "Alternator", 0.25, None, None),
            ("ALL", "LJL", "ALL", 0.60, None, None),
        ],
        "rule_Z460": [
            ("ALL", "H45", "GLASS RUN-DOOR WDW", 0.16, None, None),
            ("ALL", "HZG", "All Items(GLASS RUN-DOOR WDW)", 0.50, None, None),
        ],
        "rule_Z456": [
            ("ALL", "LJL", "CONT ASSY-USM", 0.175, "2021-07-08", None),
            ("ALL", "LFD", "CONT ASSY-USM", 0.175, "2021-07-08", None),
            ("ALL", "HZG", "CONT ASSY-USM", 0.175, "2021-07-08", None),
            ("ALL", "LJL", "UNDERHOOD SWITCHING MODULE", 0.175, "2021-07-08", None),
            ("ALL", "LFD", "UNDERHOOD SWITCHING MODULE", 0.175, "2021-07-08", None),
            ("ALL", "HZG", "UNDERHOOD SWITCHING MODULE", 0.175, "2021-07-08", None),
        ],
        "rule_Z432": [
            ("ALL", "HZG", "Radar Camera", 0.3231, "2020-09-11", None),
        ],
        "rule_Z389": [
            ("ALL", "ALL", "Radio Car", 0.35, "2021-03-23", None),
        ],
        "rule_Z386": [
            ("ALL", "ALL", "ALL", 0.50, "2020-03-01", "2021-02-29"),
            ("ALL", "ALL", "ALL", 0.1154, "2021-03-01", None),
        ],
        "rule_Z383": [
            ("ALL", "HZG", "OUTSIDE MIRROR", 0.50, None, None),
            ("ALL", "HZG", "MIRROR ASSY-DOOR", 0.50, None, None),
        ],
        "rule_Z369": [
            ("ALL", "LJL", "control assy-FR RADAR", 0.50, "2021-03-01", None),
        ],
        "rule_Z308": [
            ("ALL", "L38", "LOCK-TRUNK LID", 0.10, None, None),
            ("ALL", "L43", "LOCK-TRUNK LID", 0.10, None, None),
            ("ALL", "L47", "LOCK-TRUNK LID", 0.10, None, None),
        ],
        "rule_Z286": [
            ("ALL", "L38", "ALL", 0.42, None, None),
            ("ALL", "L43", "ALL", 0.42, None, None),
        ],
        "rule_Z262": [
            ("ALL", "L38", "STRIKER-DOOR LOCK", 0.46, None, None),
            ("ALL", "L43", "STRIKER-DOOR LOCK", 0.46, None, None),
            ("ALL", "L47", "STRIKER-DOOR LOCK", 0.46, None, None),
        ],
        # ✅ rule_Z260 중복 키 문제 해결: 하나로 합침
        "rule_Z260": [
            ("ALL", "L38", "ALL", 0.53, None, None),
            ("ALL", "L43", "ALL", 0.53, None, None),
            ("ALL", "L47", "ALL", 0.53, None, None),
            ("ALL", "L38", "COMPR-AIR COND", 0.43, None, None),
        ],
        "rule_Z246": [
            ("ALL", "L43", "82008_35879", 0.378, None, None),
            ("ALL", "L43", "82008_98810", 0.378, None, None),
            ("ALL", "L43", "92600_3748R", 0.378, None, None),
            ("ALL", "L43", "92600_7598R", 0.378, None, None),
            ("ALL", "L47", "82008_35879", 0.378, None, None),
            ("ALL", "L47", "82008_98810", 0.378, None, None),
            ("ALL", "L47", "92600_3748R", 0.378, None, None),
            ("ALL", "L47", "92600_7598R", 0.378, None, None),
        ],
        "rule_Z239": [
            ("ALL", "L38", "CABLE-SHIFT CTRL", 0.1061, None, None),
            ("ALL", "L38", "DEVICE ASSY-TRANSMISSION CONTR", 0.0741, None, None),
            ("ALL", "L38", "KNOB-SHIFT LEVER", 0.0741, None, None),
            ("ALL", "L38", "PN STOPPER ASSY-SOLENOID", 0.0741, None, None),
            ("ALL", "L38", "UNIT-GEAR CONT", 0.0741, None, None),
            ("ALL", "L43", "CABLE-SHIFT CTRL", 0.1061, None, None),
            ("ALL", "L43", "DEVICE ASSY-TRANSMISSION CONTR", 0.0741, None, None),
            ("ALL", "L43", "KNOB-SHIFT LEVER", 0.0741, None, None),
            ("ALL", "L43", "PN STOPPER ASSY-SOLENOID", 0.0741, None, None),
            ("ALL", "L43", "UNIT-GEAR CONT", 0.0741, None, None),
            ("ALL", "L47", "CABLE-SHIFT CTRL", 0.1061, None, None),
            ("ALL", "L47", "DEVICE ASSY-TRANSMISSION CONTR", 0.0741, None, None),
            ("ALL", "L47", "KNOB-SHIFT LEVER", 0.0741, None, None),
            ("ALL", "L47", "PN STOPPER ASSY-SOLENOID", 0.0741, None, None),
            ("ALL", "L47", "UNIT-GEAR CONT", 0.0741, None, None),
            ("ALL", "H45", "CABLE-SHIFT CTRL", 0.1061, None, None),
            ("ALL", "H45", "DEVICE ASSY-TRANSMISSION CONTR", 0.0741, None, None),
            ("ALL", "H45", "KNOB-SHIFT LEVER", 0.0741, None, None),
            ("ALL", "H45", "PN STOPPER ASSY-SOLENOID", 0.0741, None, None),
            ("ALL", "H45", "UNIT-GEAR CONT", 0.0741, None, None),
        ],
        "rule_Z111": [
            ("DOMESTIC", "H45", "Arm assy", 0.50, None, None),
            ("DOMESTIC", "H45", "Driver assy", 0.50, None, None),
            ("OVERSEAS", "H45", "Arm assy", 0.58, None, None),
            ("OVERSEAS", "H45", "Driver assy", 0.58, None, None),
        ],
        "rule_A201": [
            ("ALL", "L38", "All", 0.60, "2022-06-13", None),
            ("ALL", "L43", "All", 0.60, "2022-06-13", None),
            ("ALL", "L47", "All", 0.60, "2022-06-13", None),
            ("ALL", "H45", "All", 0.60, "2022-06-13", None),
            ("ALL", "LFD", "All", 0.41, "2022-06-13", None),
            ("ALL", "HZG", "All", 0.50, "2022-06-13", None),
            ("ALL", "LJL", "ALL", 0.60, None, None),
            ("ALL", "L38", "AC unit", 0.50, "2022-06-13", None),
            ("ALL", "L43", "AC unit", 0.50, "2022-06-13", None),
            ("ALL", "L47", "AC unit", 0.50, "2022-06-13", None),
            ("ALL", "H45", "AC unit", 0.50, "2022-06-13", None),
            ("ALL", "L38", "ACTUATOR-MODE", 0.41, "2022-06-13", None),
            ("ALL", "L43", "ACTUATOR-MODE", 0.41, "2022-06-13", None),
            ("ALL", "L47", "ACTUATOR-MODE", 0.41, "2022-06-13", None),
            ("ALL", "H45", "ACTUATOR-MODE", 0.41, "2022-06-13", None),
            ("ALL", "LFD", "CAP FILTER ASSY", 0.41, "2022-06-13", None),
            ("ALL", "HZG", "CAP FILTER ASSY", 0.50, "2022-06-13", None),
            ("ALL", "L38", "COND & LIQUID TANK ASSY", 0.50, "2022-06-13", None),
            ("ALL", "L43", "COND & LIQUID TANK ASSY", 0.50, "2022-06-13", None),
            ("ALL", "L47", "COND & LIQUID TANK ASSY", 0.50, "2022-06-13", None),
            ("ALL", "H45", "COND & LIQUID TANK ASSY", 0.50, "2022-06-13", None),
            ("ALL", "L38", "CONDENSER", 0.60, "2022-06-13", None),
            ("ALL", "L43", "CONDENSER", 0.60, "2022-06-13", None),
            ("ALL", "L47", "CONDENSER", 0.60, "2022-06-13", None),
            ("ALL", "H45", "CONDENSER", 0.60, "2022-06-13", None),
            ("ALL", "L38", "EVAPORATOR SENSER ASSY", 0.50, "2022-06-13", None),
            ("ALL", "L43", "EVAPORATOR SENSER ASSY", 0.50, "2022-06-13", None),
            ("ALL", "L47", "EVAPORATOR SENSER ASSY", 0.50, "2022-06-13", None),
            ("ALL", "H45", "EVAPORATOR SENSER ASSY", 0.50, "2022-06-13", None),
            ("ALL", "L38", "FAN UNIT-ENG COOLING", 0.60, "2022-06-13", None),
            ("ALL", "L43", "FAN UNIT-ENG COOLING", 0.60, "2022-06-13", None),
            ("ALL", "L47", "FAN UNIT-ENG COOLING", 0.60, "2022-06-13", None),
            ("ALL", "H45", "FAN UNIT-ENG COOLING", 0.60, "2022-06-13", None),
            ("ALL", "L38", "MODULE ASSY-POWER", 0.50, "2022-06-13", None),
            ("ALL", "L43", "MODULE ASSY-POWER", 0.50, "2022-06-13", None),
            ("ALL", "L47", "MODULE ASSY-POWER", 0.50, "2022-06-13", None),
            ("ALL", "H45", "MODULE ASSY-POWER", 0.50, "2022-06-13", None),
            ("ALL", "LFD", "MOTOR & FAN ASSY-FRONT BLOWER", 0.41, "2022-06-13", None),
            ("ALL", "HZG", "MOTOR & FAN ASSY-FRONT BLOWER", 0.50, "2022-06-13", None),
            ("ALL", "L38", "PIPE ASSY-CONDENSER", 0.50, "2022-06-13", None),
            ("ALL", "L43", "PIPE ASSY-CONDENSER", 0.50, "2022-06-13", None),
            ("ALL", "L47", "PIPE ASSY-CONDENSER", 0.50, "2022-06-13", None),
            ("ALL", "H45", "PIPE ASSY-CONDENSER", 0.50, "2022-06-13", None),
            ("ALL", "L38", "RADIATOR", 0.60, "2022-06-13", None),
            ("ALL", "L43", "RADIATOR", 0.60, "2022-06-13", None),
            ("ALL", "L47", "RADIATOR", 0.60, "2022-06-13", None),
            ("ALL", "H45", "RADIATOR", 0.60, "2022-06-13", None),
            ("ALL", "L38", "RADIATOR COMPL", 0.50, "2022-06-13", None),
            ("ALL", "L43", "RADIATOR COMPL", 0.50, "2022-06-13", None),
            ("ALL", "L47", "RADIATOR COMPL", 0.50, "2022-06-13", None),
            ("ALL", "H45", "RADIATOR COMPL", 0.50, "2022-06-13", None),
            ("ALL", "L38", "SHROUD ASSY-W/MOTOR FAN", 0.50, "2022-06-13", None),
            ("ALL", "L43", "SHROUD ASSY-W/MOTOR FAN", 0.50, "2022-06-13", None),
            ("ALL", "L47", "SHROUD ASSY-W/MOTOR FAN", 0.50, "2022-06-13", None),
            ("ALL", "H45", "SHROUD ASSY-W/MOTOR FAN", 0.50, "2022-06-13", None),
            ("ALL", "L38", "TANK ASSY-RESVR", 0.50, "2022-06-13", None),
            ("ALL", "L43", "TANK ASSY-RESVR", 0.50, "2022-06-13", None),
            ("ALL", "L47", "TANK ASSY-RESVR", 0.50, "2022-06-13", None),
            ("ALL", "H45", "TANK ASSY-RESVR", 0.50, "2022-06-13", None),
            ("ALL", "L38", "EXPANSION VALVE", 0.50, "2022-06-13", None),
            ("ALL", "L43", "EXPANSION VALVE", 0.50, "2022-06-13", None),
            ("ALL", "L47", "EXPANSION VALVE", 0.50, "2022-06-13", None),
            ("ALL", "H45", "EXPANSION VALVE", 0.50, "2022-06-13", None),
        ],
        "rule_B205": [
            ("all", "L38", "SUNROOF", 44, None, None),
        ],
        "rule_B508": [
            ("All", "L38", "All", 60, None, None),
            ("All", "L43", "All", 60, None, None),
            ("All", "L47", "All", 60, None, None),
            ("All", "H45", "All", 60, None, None),
            ("All", "LFD", "All", 50, None, None),
            ("All", "HZG", "All", 50, None, None),
        ],
        "rule_B904": [
            ("all", "L47", "All", 35, None, None),
            ("all", "LFD", "All", 35, None, None),
        ],
        "rule_B907": [
            ("ALL", "L38", "All", 50, None, None),
            ("ALL", "L43", "All", 50, None, None),
            ("ALL", "L47", "All", 50, None, None),
            ("ALL", "L38", "Head Lamp Xenon", 65, None, None),
            ("ALL", "L43", "Head Lamp Xenon", 65, None, None),
            ("ALL", "L47", "Head Lamp Xenon", 65, None, None),
            ("ALL", "H45", "All Lamp", 50, None, None),
            ("ALL", "LFD", "Head Lamp", 50, None, None),
            ("ALL", "HZG", "Head Lamp", 50, None, None),
        ],
        "rule_B908": [
            ("all", "LFD", "All", 50, None, None),
            ("all", "HZG", "All", 50, None, None),
        ],
        "rule_B923": [
            ("all", "L43", "all", 50, None, None),
            ("all", "L47", "all", 50, None, None),
            ("all", "L38", "MECHANISM-WS WIPER", 51, None, None),
            ("all", "L43", "MECHANISM-WS WIPER", 70, None, None),
            ("all", "L47", "MECHANISM-WS WIPER", 70, None, None),
            ("all", "HZG", "All", 50, None, None),
        ],
        "rule_B928": [
            ("all", "H45", "Wiper", 50, None, None),
            ("all", "LFD", "Wiper", 50, None, None),
        ],
        "rule_B932": [
            ("all", "H45", "GLASS RUN-DOOR WDW", 16, None, None),
            ("all", "HZG", "All", 50, None, None),
        ],
        "rule_B933": [
            ("all", "H45", "all", 50.00, None, None),
            ("all", "H45", "Condenser", 33.00, None, None),
            ("all", "H45", "Horn assy", 35.00, None, None),
            ("all", "H45", "Windshild washer tank", 50.00, None, None),
            ("all", "H45", "Pump washer", 67.00, None, None),
            ("all", "H45", "Radiator", 75.00, None, None),
            ("all", "H45", "Bolt", 80.00, None, None),
            ("all", "L43", "All", 50.00, None, None),
            ("DOMESTIC", "H45", "MOTOR ASSY", 5, None, None),
            ("DOMESTIC", "H45", "SHROUD ASSY-W/MOTOR FAN", 5, None, None),
            ("OVERSEAS", "H45", "MOTOR ASSY", 0, None, None),
            ("OVERSEAS", "H45", "SHROUD ASSY-W/MOTOR FAN", 0, None, None),
        ],
    }

    bulk_insert_rules(rules_to_insert)
