import sqlite3
from pathlib import Path

DB_PATH = Path("data/TestDB.sqlite")

DEFAULT_MILEAGE = 60000
DEFAULT_PERIOD = 3


def table_exists(conn: sqlite3.Connection, name: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (name,),
    ).fetchone() is not None


def migrate():
    if not DB_PATH.exists():
        raise FileNotFoundError(f"DB not found: {DB_PATH.resolve()}")

    with sqlite3.connect(str(DB_PATH)) as conn:
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute("BEGIN")

        try:
            # 1) warranty 테이블 생성 (전역 1행)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS warranty (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    warranty_mileage INTEGER NOT NULL,
                    warranty_period  INTEGER NOT NULL,
                    created_at TEXT DEFAULT (DATETIME('now','localtime')),
                    updated_at TEXT DEFAULT (DATETIME('now','localtime'))
                );
            """)

            # 2) 기본값 1행 보장
            conn.execute("""
                INSERT OR IGNORE INTO warranty (id, warranty_mileage, warranty_period)
                VALUES (1, ?, ?);
            """, (DEFAULT_MILEAGE, DEFAULT_PERIOD))

            # 3) 예전 sap_warranty 테이블 있으면 삭제
            if table_exists(conn, "sap_warranty"):
                conn.execute('DROP TABLE "sap_warranty"')

            conn.commit()
            print("[DONE] warranty table ready (single row), sap_warranty dropped")

        except Exception:
            conn.rollback()
            raise
        finally:
            conn.execute("PRAGMA foreign_keys = ON")


if __name__ == "__main__":
    migrate()
