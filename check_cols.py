import sqlite3
from src.database import DB_PATH

conn = sqlite3.connect(str(DB_PATH))
cur = conn.cursor()

cur.execute("PRAGMA table_info(sap)")
print("sap columns:", [r[1] for r in cur.fetchall()])

cur.execute("PRAGMA table_info(sap_warranty)")
print("sap_warranty columns:", [r[1] for r in cur.fetchall()])

conn.close()
