"""
룰테이블 우선순위 일괄 업데이트 스크립트

모든 룰테이블의 priority를 행 순서(rule_id 순서)에 따라 1, 2, 3... 으로 업데이트합니다.

사용법:
    python update_rule_priorities.py
"""

import sys
from pathlib import Path
import sqlite3
from typing import List

# 데이터베이스 경로
DB_PATH = Path("data/TestDB.sqlite")


def get_all_rule_tables() -> List[str]:
    """모든 rule_* 테이블 목록 조회"""
    if not DB_PATH.exists():
        return []

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    cursor.execute("""
        SELECT name FROM sqlite_master
        WHERE type='table' AND name LIKE 'rule_%'
        ORDER BY name
    """)

    rows = cursor.fetchall()
    conn.close()
    return [row[0] for row in rows] if rows else []


def get_rule_ids_ordered(rule_table_name: str) -> List[int]:
    """룰테이블에서 rule_id를 오름차순으로 조회"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cursor.execute(f"""
            SELECT rule_id FROM "{rule_table_name}"
            ORDER BY rule_id ASC
        """)
        rows = cursor.fetchall()
        return [row[0] for row in rows] if rows else []
    except sqlite3.OperationalError as e:
        print(f"  ⚠️  테이블 조회 실패: {e}")
        return []
    finally:
        conn.close()


def update_rule_priorities(rule_table_name: str, rule_ids_in_order: List[int]) -> bool:
    """룰테이블의 우선순위를 rule_id 순서에 따라 1, 2, 3... 으로 업데이트"""
    if not rule_table_name or not rule_table_name.startswith("rule_"):
        raise ValueError(f"유효하지 않은 rule 테이블명: {rule_table_name}")

    if not rule_ids_in_order:
        return True

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        for new_priority, rule_id in enumerate(rule_ids_in_order, start=1):
            cursor.execute(f'''
                UPDATE "{rule_table_name}"
                SET priority = ?, updated_at = DATETIME('now', 'localtime')
                WHERE rule_id = ?
            ''', (new_priority, rule_id))

        conn.commit()
        return True
    except sqlite3.Error as e:
        conn.rollback()
        raise ValueError(f"우선순위 업데이트 실패: {str(e)}")
    finally:
        conn.close()


def main():
    """메인 실행 함수"""
    print("=" * 60)
    print("룰테이블 우선순위 일괄 업데이트")
    print("=" * 60)
    print()

    # 데이터베이스 파일 확인
    if not DB_PATH.exists():
        print(f"❌ 데이터베이스 파일을 찾을 수 없습니다: {DB_PATH}")
        print("   현재 디렉토리:", Path.cwd())
        sys.exit(1)

    # 모든 룰테이블 조회
    rule_tables = get_all_rule_tables()
    
    if not rule_tables:
        print("⚠️  룰테이블이 없습니다.")
        sys.exit(0)

    print(f"📋 발견된 룰테이블: {len(rule_tables)}개")
    print()

    # 각 룰테이블 처리
    total_updated = 0
    total_rules = 0
    failed_tables = []

    for table_name in rule_tables:
        print(f"처리 중: {table_name}")
        
        try:
            # rule_id 순서대로 조회
            rule_ids = get_rule_ids_ordered(table_name)
            
            if not rule_ids:
                print(f"  ⚠️  규칙이 없습니다. 건너뜁니다.")
                print()
                continue

            # 우선순위 업데이트
            update_rule_priorities(table_name, rule_ids)
            
            total_updated += 1
            total_rules += len(rule_ids)
            print(f"  ✅ 완료: {len(rule_ids)}개 규칙의 우선순위를 업데이트했습니다.")
            print(f"     (priority: 1 ~ {len(rule_ids)})")
            print()

        except Exception as e:
            print(f"  ❌ 실패: {str(e)}")
            failed_tables.append((table_name, str(e)))
            print()

    # 결과 요약
    print("=" * 60)
    print("결과 요약")
    print("=" * 60)
    print(f"✅ 성공한 테이블: {total_updated}개")
    print(f"📊 총 업데이트된 규칙: {total_rules}개")
    
    if failed_tables:
        print(f"❌ 실패한 테이블: {len(failed_tables)}개")
        for table_name, error in failed_tables:
            print(f"   - {table_name}: {error}")
    else:
        print("✅ 모든 테이블이 성공적으로 처리되었습니다.")
    
    print()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  사용자에 의해 중단되었습니다.")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ 예상치 못한 오류가 발생했습니다: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
