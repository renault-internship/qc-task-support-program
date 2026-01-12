import sqlite3

conn = sqlite3.connect("data/TestDB.sqlite")
cur = conn.cursor()

cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'rule_%' ORDER BY name")
tables = [r[0] for r in cur.fetchall()]

print("rule tables:", len(tables))
print("--------------------------------")

for t in tables:
    cur.execute(f"SELECT COUNT(*) FROM {t}")
    cnt = cur.fetchone()[0]
    print(t.ljust(20), cnt)

conn.close()
