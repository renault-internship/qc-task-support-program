import sqlite3
from src.database import DB_PATH

with sqlite3.connect(str(DB_PATH)) as conn:
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM sap")
    sap_cnt = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM sap_warranty")
    w_cnt = cur.fetchone()[0]

    cur.execute("""
        SELECT COUNT(*)
        FROM sap s
        LEFT JOIN sap_warranty w ON w.sap_code = s.sap_code
        WHERE w.sap_code IS NULL
    """)
    missing = cur.fetchone()[0]

    print("sap count:", sap_cnt)
    print("sap_warranty count:", w_cnt)
    print("missing warranty rows:", missing)
