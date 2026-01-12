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

# ✅ 여기만 바꾸면 됨: L38/L43/L47 프로젝트코드 행은 CSV에서 읽을 때부터 스킵
SKIP_PROJECT_CODES = {"L38", "L43", "L47"}  # 필요 없으면 set() 로 비워두면 됨

# ✅ PK(autoincrement) 초기화까지 하고 싶으면 True
RESET_AUTOINCREMENT = True

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

    # 80 -> 0.8 같은 케이스
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


def reset_autoincrement_if_possible(conn: sqlite3.Connection, table_name: str) -> None:
    """
    - SQLite에서 AUTOINCREMENT 쓰는 테이블이면 sqlite_sequence에 기록됨
    - DELETE 후 sqlite_sequence 행 제거하면 다음 insert가 다시 1부터 시작
    - AUTOINCREMENT가 아니면 sqlite_sequence에 없을 수 있으니 조용히 무시
    """
    try:
        conn.execute("DELETE FROM sqlite_sequence WHERE name=?", (table_name,))
    except Exception:
        pass


def parse_csv_row_to_full_rule_row(d: Dict[str, str]) -> FullRuleRow:
    # ✅ (B) 여기서도 깨진 CSV를 더 친절하게 잡기
    if None in d:
        raise ValueError(
            "CSV column mismatch detected (extra columns). "
            "Likely unquoted comma inside a field. "
            f"extras={d.get(None)} row={d}"
        )

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

    ct = normalize_cap_type(d.get("cap_type"))
    cv = normalize_nullable_int(d.get("cap_value"))

    # ✅ (C-1) cap_type에 숫자가 들어간 경우 보정 (컬럼이 밀린 경우)
    if lr is None and ct == "NONE":
        # cap_type 원본 값 확인
        cap_type_raw = d.get("cap_type", "").strip()
        if cap_type_raw:
            try:
                tmp_lr = normalize_liability_ratio(cap_type_raw)
                if tmp_lr is not None:
                    lr = tmp_lr
                    ct = "NONE"  # cap_type은 비워둠
            except Exception:
                pass

    # ✅ (C-2) 임시 보정: 보증 없고 lr 비었는데 cap_type은 없고 cap_value만 0~1이면 lr로 구제
    if (mcap is None and pcap is None) and (lr is None):
        # normalize_cap_type은 비어도 "NONE"을 반환하니까, "NONE"이면 cap_type 없는 취급
        if (ct == "NONE") and (cv is not None):
            try:
                tmp = float(str(cv).strip())
                if 0 < tmp <= 1:
                    lr = tmp
                    cv = None
            except Exception:
                pass

    # ✅ 보증 오버라이드(mcap/pcap)가 있으면 lr 없어도 OK
    # lr도 없고 보증도 없으면 에러
    if lr is None and (mcap is None and pcap is None):
        raise ValueError(f"liability_ratio is required when no mileage_cap/period_cap. row={d}")

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
            # ✅ (A-1) CSV 컬럼 불일치 처리: None 키가 있으면 제거하고 경고
            if None in d:
                extras = d.get(None, [])
                # 빈 문자열만 있는 경우는 무시하고 계속 진행
                if extras and any(str(e).strip() for e in extras):
                    # 실제 값이 있는 경우에만 에러
                    raw_line = ""
                    try:
                        with csv_path.open("r", encoding="utf-8-sig", newline="") as f2:
                            lines = f2.readlines()
                            if line_no - 1 < len(lines):
                                raw_line = lines[line_no - 1].strip()
                    except:
                        pass
                    
                    raise ValueError(
                        f"CSV column mismatch at line {line_no} (unquoted comma inside a field).\n"
                        f"Expected {len(header)} columns, but found more.\n"
                        f"Raw line: {raw_line}\n"
                        f"Parsed row: {d}\n"
                        f"Extras: {extras}\n"
                        f"Please check if fields with commas are properly quoted."
                    )
                else:
                    # 빈 값만 있으면 제거하고 계속 진행
                    del d[None]
                    print(f"Warning: Line {line_no} has extra empty columns, ignoring...")

            table_name_raw = (d.get("rule_") or "").strip()
            if not table_name_raw:
                continue

            # ✅ 프로젝트코드 스킵: L38/L43/L47 + "L38 EV" / "L38/..." 같은 변형도 스킵
            pc_norm = normalize_text(d.get("project_code"), "ALL")
            pc_tokens = re.split(r"[\s/]+", pc_norm.strip())
            if any(t in SKIP_PROJECT_CODES for t in pc_tokens if t):
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
    - 테이블별로 '기존 데이터 전체 삭제' 후 CSV 데이터를 다시 넣는다. (덮어쓰기)
    - RESET_AUTOINCREMENT=True면 가능할 때 PK 시퀀스도 초기화 시도
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

                ensure_unique_index(conn, actual_table_name)

                before = conn.execute(
                    f'SELECT COUNT(*) AS cnt FROM "{actual_table_name}"'
                ).fetchone()["cnt"]

                # ✅ 기존 데이터 전부 삭제
                conn.execute(f'DELETE FROM "{actual_table_name}"')
                if RESET_AUTOINCREMENT:
                    reset_autoincrement_if_possible(conn, actual_table_name)

                deleted_here = before
                total_deleted += deleted_here

                 # ✅ CSV 데이터를 "그대로" 다시 삽입 (중복은 스킵)
                conn.executemany(
                    f"""
                    INSERT OR IGNORE INTO "{actual_table_name}"
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

                inserted_here = after
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
