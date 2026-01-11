"""
sap.remark만 업데이트하는 스크립트
- insert_sap_data_bulk.py는 그대로 두고 remark만 별도 관리
사용법:
  1) python insert_sap_data_bulk.py
  2) python update_sap_remarks.py
"""
import sqlite3
from pathlib import Path

from src.database import init_database
from sap_remarks import SAP_REMARKS

DB_PATH = Path("data/TestDB.sqlite")


def ensure_sap_remark_column(conn: sqlite3.Connection) -> None:
    """
    sap 테이블에 remark 컬럼이 없으면 추가
    """
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(sap)")
    cols = {row[1] for row in cur.fetchall()}  # row[1] = column name
    if "remark" not in cols:
        cur.execute("ALTER TABLE sap ADD COLUMN remark TEXT")
        conn.commit()


def main() -> None:
    init_database()  # sap 테이블 존재 보장
    conn = sqlite3.connect(str(DB_PATH))
    try:
        ensure_sap_remark_column(conn)

        updated = 0
        missing = 0
        empty = 0

        cur = conn.cursor()

        for sap_code_raw, remark_raw in SAP_REMARKS.items():
            sap_code = str(sap_code_raw).strip().upper()
            remark = (remark_raw or "").strip()

            if not remark:
                empty += 1
                continue

            # 존재 확인
            cur.execute("SELECT 1 FROM sap WHERE sap_code = ? LIMIT 1", (sap_code,))
            if cur.fetchone() is None:
                missing += 1
                continue

            # 업데이트
            cur.execute(
                "UPDATE sap SET remark = ? WHERE sap_code = ?",
                (remark, sap_code),
            )
            updated += cur.rowcount

        conn.commit()

        print("=" * 70)
        print("sap.remark 업데이트 완료")
        print(f"- 업데이트: {updated}건")
        print(f"- DB에 없는 sap_code: {missing}건")
        print(f"- remark 비어있음(스킵): {empty}건")
        print("=" * 70)

    finally:
        conn.close()


if __name__ == "__main__":
    main()
