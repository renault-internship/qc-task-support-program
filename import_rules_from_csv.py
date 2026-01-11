import csv
import re
import sqlite3
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

# 실행방법: python import_rules_from_csv.py <rules.csv>

DB_PATH = Path("data/TestDB.sqlite")

CSV_COLUMNS = [
    "rule_",
    "repair_region",
    "project_code",
    "exclude_project_code",
    "vehicle_classification",
    "part_name",
    "part_no",
    "engine_form",
    "mileage_cap",
    "period_cap",
    "liability_ratio",
    "cap_type",
    "cap_value",
    "note",
    "valid_from",
    "valid_to",
]

# (repair_region, project_code, exclude_project_code, vehicle_classification,
#  part_no, part_name, engine_form, mileage_cap, period_cap,
#  liability_ratio, cap_type, cap_value, note, valid_from, valid_to)
FullRuleRow = Tuple[
    str, str, Optional[str], str,
    str, str, str, Optional[int], Optional[int],
    Optional[float], str, Optional[int], Optional[str],
    Optional[str], Optional[str]
]


def _strip_smart_quotes(s: str) -> str:
    return s.replace("‘", "").replace("’", "").replace('"', "").strip()


def safe_table_name(name: str) -> str:
    n = _strip_smart_quotes(str(name).strip())
    if not re.fullmatch(r"rule_[A-Za-z0-9_]+", n):
        raise ValueError(f"Invalid table name: {name}")
    return n


def normalize_text(v: Any, default: str) -> str:
    if v is None:
        return default
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return default
    if s.lower() == "all":
        return "ALL"
    return s


def normalize_nullable_text(v: Any) -> Optional[str]:
    if v is None:
        return None
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return None
    return s


def normalize_nullable_date(v: Any) -> Optional[str]:
    return normalize_nullable_text(v)


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
    if v is None:
        return "NONE"
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return "NONE"
    s_up = s.upper()

    if s_up in {"NONE", "NO", "N/A", "NA"}:
        return "NONE"
    if s_up == "LABOR":
        return "LABOR"
    if s_up in {"OUTSOURCE_LABOR", "OUTSOURCELABOR", "OUTSOURCE"}:
        return "OUTSOURCE_LABOR"
    if s_up in {"BOTH_LABOR", "BOTH", "LABOR+OUTSOURCE", "LABOR_OUTSOURCE"}:
        return "BOTH_LABOR"

    return "NONE"


def normalize_liability_ratio(v: Any) -> Optional[float]:
    if v is None:
        return None
    s = _strip_smart_quotes(str(v).strip())
    if s == "":
        return None

    s = s.replace("%", "").strip()
    try:
        lr = float(s)
    except Exception:
        return None

    if lr > 1.0:
        lr = lr / 100.0
    return lr


def get_existing_table_name(conn: sqlite3.Connection, requested: str) -> Optional[str]:
    req = safe_table_name(requested)

    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (req,),
    ).fetchone()
    if row:
        return row[0]

    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND lower(name)=lower(?)",
        (req,),
    ).fetchone()
    if row:
        return row[0]

    return None


def get_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
    cur = conn.execute(f'PRAGMA table_info("{table_name}")')
    return [r[1] for r in cur.fetchall()]


def ensure_unique_index(conn: sqlite3.Connection, table_name: str) -> None:
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


def parse_csv_row_to_full_rule_row(d: Dict[str, str]) -> FullRuleRow:
    rr = normalize_text(d.get("repair_region"), "ALL")
    pc = normalize_text(d.get("project_code"), "ALL")
    epc = normalize_nullable_text(d.get("exclude_project_code"))
    vc = normalize_text(d.get("vehicle_classification"), "ALL")

    pname = normalize_text(d.get("part_name"), "ALL")
    pno = normalize_text(d.get("part_no"), "ALL")
    eng = normalize_text(d.get("engine_form"), "ALL")

    mcap = normalize_nullable_int(d.get("mileage_cap"))
    pcap = normalize_nullable_int(d.get("period_cap"))

    lr = normalize_liability_ratio(d.get("liability_ratio"))

    # ✅ 핵심: 보증 오버라이드(mcap/pcap)가 있으면 lr 없어도 OK
    # lr도 없고 보증도 없으면 에러
    if lr is None and (mcap is None and pcap is None):
        raise ValueError(f"liability_ratio is required when no mileage_cap/period_cap. row={d}")

    ct = normalize_cap_type(d.get("cap_type"))
    cv = normalize_nullable_int(d.get("cap_value"))

    note = normalize_nullable_text(d.get("note"))
    vf = normalize_nullable_date(d.get("valid_from"))
    vt = normalize_nullable_date(d.get("valid_to"))

    return (
        rr, pc, epc, vc,
        pno, pname, eng, mcap, pcap,
        lr, ct, cv, note,
        vf, vt
    )


def load_rules_from_csv(csv_path: Path) -> Dict[str, List[tuple]]:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {csv_path.resolve()}")

    rules_to_insert: Dict[str, List[tuple]] = {}

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)

        header = reader.fieldnames or []
        missing = [c for c in CSV_COLUMNS if c not in header]
        if missing:
            raise ValueError(f"CSV header missing columns: {missing}. got={header}")

        for line_no, d in enumerate(reader, start=2):
            table_name_raw = (d.get("rule_") or "").strip()
            if not table_name_raw:
                continue

            table_name = safe_table_name(table_name_raw)

            try:
                row15 = parse_csv_row_to_full_rule_row(d)
            except Exception as e:
                raise ValueError(f"CSV parse error at line {line_no}: {e}")

            rules_to_insert.setdefault(table_name, []).append(row15)

    return rules_to_insert


def bulk_insert_rules(rules_to_insert: Dict[str, List[tuple]]):
    """
    ✅ 변경점:
    - 테이블별로 '기존 데이터 전체 삭제' 후 CSV 데이터를 다시 넣는다.
    - 즉, 중복 스킵(INSERT OR IGNORE) 같은 개념이 아니라 "덮어쓰기"이다.
    """
    if not DB_PATH.exists():
        raise FileNotFoundError(f"DB not found: {DB_PATH.resolve()}")

    with sqlite3.connect(str(DB_PATH)) as conn:
        conn.row_factory = sqlite3.Row

        print(f"[DB] {DB_PATH.resolve()}")
        conn.execute("BEGIN IMMEDIATE")
        per_table_stats: List[Dict[str, Any]] = []

        try:
            total_attempted = 0
            total_inserted = 0
            total_deleted = 0
            total_tables_ok = 0
            total_tables_skipped = 0

            for table_name_req, rows in rules_to_insert.items():
                attempted_here = len(rows)
                total_attempted += attempted_here

                if attempted_here == 0:
                    print(f"[SKIP] {table_name_req}: rows empty")
                    total_tables_skipped += 1
                    per_table_stats.append({
                        "table": table_name_req,
                        "attempted": 0,
                        "deleted": 0,
                        "inserted": 0,
                        "before": None,
                        "after": None,
                        "status": "SKIP(rows empty)"
                    })
                    continue

                actual_table_name = get_existing_table_name(conn, table_name_req)
                if not actual_table_name:
                    print(f"[SKIP] {table_name_req}: table not found")
                    total_tables_skipped += 1
                    per_table_stats.append({
                        "table": table_name_req,
                        "attempted": attempted_here,
                        "deleted": 0,
                        "inserted": 0,
                        "before": None,
                        "after": None,
                        "status": "SKIP(table not found)"
                    })
                    continue

                cols = set(get_columns(conn, actual_table_name))

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
                    print(f"[SKIP] {actual_table_name}: missing columns {sorted(missing)}")
                    total_tables_skipped += 1
                    per_table_stats.append({
                        "table": actual_table_name,
                        "attempted": attempted_here,
                        "deleted": 0,
                        "inserted": 0,
                        "before": None,
                        "after": None,
                        "status": f"SKIP(missing columns: {sorted(missing)})"
                    })
                    continue

                # 유니크 인덱스는 유지/재생성(있어도 무해)
                ensure_unique_index(conn, actual_table_name)

                before = conn.execute(
                    f'SELECT COUNT(*) AS cnt FROM "{actual_table_name}"'
                ).fetchone()["cnt"]

                # ✅ 핵심: 기존 데이터 전부 삭제
                conn.execute(f'DELETE FROM "{actual_table_name}"')
                deleted_here = before
                total_deleted += deleted_here

                # ✅ CSV 데이터를 "그대로" 다시 삽입
                conn.executemany(
                    f"""
                    INSERT INTO "{actual_table_name}"
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
                    rows,
                )

                after = conn.execute(
                    f'SELECT COUNT(*) AS cnt FROM "{actual_table_name}"'
                ).fetchone()["cnt"]

                inserted_here = after  # 삭제 후 넣었으니 after가 곧 inserted
                total_inserted += inserted_here
                total_tables_ok += 1

                print(
                    f"[OK] {actual_table_name}: attempted={attempted_here}, deleted={deleted_here}, inserted={inserted_here}, before={before}, after={after}"
                )

                per_table_stats.append({
                    "table": actual_table_name,
                    "attempted": attempted_here,
                    "deleted": deleted_here,
                    "inserted": inserted_here,
                    "before": before,
                    "after": after,
                    "status": "OK"
                })

            conn.commit()

            print("")
            print("============================================================")
            print(f"[SUMMARY] tables_ok={total_tables_ok}, tables_skipped={total_tables_skipped}")
            print(f"[SUMMARY] attempted_rows={total_attempted}, deleted_rows={total_deleted}, inserted_rows={total_inserted}")
            print("============================================================")

            # 테이블별 요약(삽입된 것 우선 정렬)
            per_table_stats_sorted = sorted(
                per_table_stats,
                key=lambda x: (-(x["inserted"] or 0), x["table"])
            )
            for s in per_table_stats_sorted:
                b = s["before"]
                a = s["after"]
                b_str = "-" if b is None else str(b)
                a_str = "-" if a is None else str(a)
                del_str = str(s.get("deleted", 0)) if b is not None else "-"
                print(
                    f"- {s['table']}: {s['status']} | attempted={s['attempted']} deleted={del_str} inserted={s['inserted']} before={b_str} after={a_str}"
                )
            print("============================================================")

        except Exception as e:
            conn.rollback()
            print("[ROLLBACK] error:", repr(e))
            raise


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Usage: python import_rules_from_csv.py <rules.csv>")

    csv_path = Path(sys.argv[1])
    rules_to_insert = load_rules_from_csv(csv_path)
    bulk_insert_rules(rules_to_insert)


if __name__ == "__main__":
    main()
