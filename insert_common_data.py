"""
common_project_liability 테이블에 LJL 기본 구상률 60(%) 넣기
- 60 / 0.6 둘 다 받아서 DB에는 0~1 스케일로 저장(60 -> 0.6)
사용법: python insert_common_ljl.py
"""

from src.database import init_database, upsert_common_project_liability, get_all_common_project_liabilities
# database.py가 src/가 아니라 루트면 아래로 바꿔
# from database import init_database, upsert_common_project_liability, get_all_common_project_liabilities


def normalize_ratio(x) -> float:
    v = float(x)
    if v > 1.0:
        v = v / 100.0
    return v


def main():
    init_database()

    project_code = "LJL"
    ratio_input = 60  # 너가 원하는 값: 60

    ratio = normalize_ratio(ratio_input)  # 60 -> 0.6
    upsert_common_project_liability(project_code, ratio)

    print("[OK] inserted/updated common_project_liability")
    print("project_code =", project_code, "liability_ratio =", ratio)

    print("\n[DB] all rows:")
    for r in get_all_common_project_liabilities():
        print(r)


if __name__ == "__main__":
    main()
