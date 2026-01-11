from src.database import init_database
import sqlite3
from pathlib import Path

DB_PATH = Path("data/TestDB.sqlite")

init_database()

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
cur = conn.cursor()

rows = cur.execute("""
    SELECT
        sap_code,
        sap_name,
        renault_code,
        rule_table_name,
        created_at,
        updated_at
    FROM sap
    ORDER BY sap_code
""").fetchall()

print("=" * 100)
print(f"총 {len(rows)}개 협력사")
print("=" * 100)

for r in rows:
    print(
        f"{r['sap_code']:8} | "
        f"{(r['sap_name'] or ''):40} | "
        f"{(r['renault_code'] or ''):10} | "
        f"{(r['rule_table_name'] or '')}"
    )

conn.close()