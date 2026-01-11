# check_csv_columns.py
# 사용법:
#   python check_csv_columns.py rule_a.csv
# 옵션:
#   python check_csv_columns.py rule_a.csv --fix   (끝의 불필요한 콤마 때문에 생기는 "빈 초과 컬럼"만 자동 제거해서 *_fixed.csv 생성)

import csv
import sys
from pathlib import Path
from typing import List, Optional

EXPECTED_COLUMNS = [
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
EXPECTED_N = len(EXPECTED_COLUMNS)


def read_header_and_validate(path: Path) -> None:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        header = next(r, None)
        if header is None:
            raise SystemExit("CSV is empty")
        if header != EXPECTED_COLUMNS:
            # 헤더 순서/이름이 다르면 DictReader가 엉킬 수 있어서 강하게 안내
            print("[WARN] Header mismatch.")
            print("  expected:", EXPECTED_COLUMNS)
            print("  got     :", header)
        else:
            print("[OK] Header matches expected columns.")


def is_all_empty(fields: List[str]) -> bool:
    return all((x is None) or (str(x).strip() == "") for x in fields)


def check_rows(path: Path, fix: bool = False) -> Optional[Path]:
    out_path = None
    out_f = None
    w = None

    bad_count = 0
    fixed_count = 0
    total = 0

    if fix:
        out_path = path.with_name(path.stem + "_fixed" + path.suffix)
        out_f = out_path.open("w", encoding="utf-8", newline="")
        w = csv.writer(out_f)

    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            r = csv.reader(f)
            header = next(r, None)
            if header is None:
                raise SystemExit("CSV is empty")

            # write header
            if fix and w is not None:
                w.writerow(header)

            for line_no, row in enumerate(r, start=2):
                total += 1
                n = len(row)

                if n == EXPECTED_N:
                    if fix and w is not None:
                        w.writerow(row)
                    continue

                # n != 16 이면 에러 후보
                # 1) 더 많은 경우: 끝에 콤마가 여러 개 붙어서 빈 필드가 초과된 케이스는 자동 복구 가능
                if n > EXPECTED_N:
                    extras = row[EXPECTED_N:]
                    if fix and is_all_empty(extras):
                        row_fixed = row[:EXPECTED_N]
                        w.writerow(row_fixed)
                        fixed_count += 1
                        continue

                bad_count += 1
                preview = ",".join(row[:min(len(row), 20)])
                print(f"[BAD] line {line_no}: fields={n} (expected {EXPECTED_N}) | head={preview}")

                if fix and w is not None:
                    # 자동 수정 불가능한 케이스는 그대로 기록(원형 보존)
                    w.writerow(row)

        print("")
        print("========================================")
        print(f"[SUMMARY] total_rows={total}")
        print(f"[SUMMARY] bad_rows={bad_count}")
        if fix:
            print(f"[SUMMARY] auto_fixed_rows={fixed_count}")
            print(f"[OUT] wrote: {out_path}")
        print("========================================")
        return out_path

    finally:
        if out_f is not None:
            out_f.close()


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Usage: python check_csv_columns.py <file.csv> [--fix]")

    csv_path = Path(sys.argv[1])
    fix = ("--fix" in sys.argv[2:])

    if not csv_path.exists():
        raise SystemExit(f"File not found: {csv_path.resolve()}")

    read_header_and_validate(csv_path)
    check_rows(csv_path, fix=fix)


if __name__ == "__main__":
    main()
