"""
DB 스키마 마이그레이션 도구 - 독립 실행
사용법: python db_schema_migrate.py

현재 기준 변경 사항(재실행 안전):
1) sap 테이블에 remark 컬럼 추가 (없을 때만)
2) sap 테이블에 renault_code 컬럼 추가 (없을 때만)
3) rule_* 테이블들에 note 컬럼 추가 (없을 때만)

주의:
- 이미 한 번 실행된 DB에 다시 실행해도 중복/에러 없이 "스킵"되도록 설계
- rule 테이블별로 실패해도 전체 중단하지 않고 계속 진행(로그에 남김)
"""

import sys
import re
from pathlib import Path
import sqlite3
from typing import List

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QLabel, QMessageBox
)

# 데이터베이스 경로
DB_PATH = Path("data/TestDB.sqlite")


# -------------------------
# Helpers
# -------------------------

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


def safe_ident(name: str) -> str:
    """
    테이블명 안전성 체크
    - rule_* 형태만 허용
    - 영문/숫자/_ 만 허용
    """
    if not re.fullmatch(r"rule_[A-Za-z0-9_]+", name):
        raise ValueError(f"Invalid rule table name: {name}")
    return name


def get_table_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
    """테이블의 컬럼 목록 조회 (연결 재사용)"""
    # PRAGMA 는 파라미터 바인딩이 안 되므로, 식별자만 엄격히 검증/인용
    cursor = conn.cursor()
    cursor.execute(f'PRAGMA table_info("{table_name}")')
    rows = cursor.fetchall()
    return [row[1] for row in rows]  # row[1]은 컬럼명


# -------------------------
# Migrations
# -------------------------

def add_remark_to_sap() -> str:
    """sap 테이블에 remark 컬럼 추가 (없을 때만)"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cols = get_table_columns(conn, "sap")
        if "remark" in cols:
            return "sap: remark 컬럼이 이미 존재합니다."

        cursor.execute("ALTER TABLE sap ADD COLUMN remark TEXT DEFAULT ''")
        conn.commit()
        return "sap: remark 컬럼이 추가되었습니다."
    except sqlite3.OperationalError as e:
        conn.rollback()
        raise Exception(f"sap: remark 컬럼 추가 실패: {str(e)}")
    finally:
        conn.close()


def add_renault_code_to_sap() -> str:
    """sap 테이블에 renault_code 컬럼 추가 (없을 때만)"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cols = get_table_columns(conn, "sap")
        if "renault_code" in cols:
            return "sap: renault_code 컬럼이 이미 존재합니다."

        cursor.execute("ALTER TABLE sap ADD COLUMN renault_code TEXT DEFAULT ''")
        conn.commit()
        return "sap: renault_code 컬럼이 추가되었습니다."
    except sqlite3.OperationalError as e:
        conn.rollback()
        raise Exception(f"sap: renault_code 컬럼 추가 실패: {str(e)}")
    finally:
        conn.close()


def add_note_to_rule_table(table_name: str) -> str:
    """
    rule_* 테이블에 note 컬럼 추가 (없을 때만)
    - SQLite는 ADD COLUMN 지원하므로 테이블 복사 방식 불필요
    """
    safe_ident(table_name)

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    try:
        cols = get_table_columns(conn, table_name)
        if "note" in cols:
            return f'{table_name}: note 컬럼이 이미 존재합니다.'

        # note는 NULL 허용 + 기본값 ''(원하면 DEFAULT NULL로 바꿔도 됨)
        cursor.execute(f'ALTER TABLE "{table_name}" ADD COLUMN note TEXT DEFAULT ""')
        conn.commit()
        return f'{table_name}: note 컬럼이 추가되었습니다.'
    except sqlite3.OperationalError as e:
        conn.rollback()
        raise Exception(f'{table_name}: note 컬럼 추가 실패: {str(e)}')
    finally:
        conn.close()


# -------------------------
# UI
# -------------------------

class SchemaMigrationWindow(QWidget):
    """스키마 마이그레이션 윈도우"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DB 스키마 마이그레이션")
        self.setFixedSize(650, 520)

        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        info_label = QLabel(
            "변경 사항(재실행 안전):\n"
            "1) sap 테이블에 remark 컬럼 추가\n"
            "2) sap 테이블에 renault_code 컬럼 추가\n"
            "3) rule_* 테이블들에 note 컬럼 추가\n\n"
            "※ 이미 적용된 항목은 자동으로 스킵됩니다."
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        log_label = QLabel("실행 로그:")
        layout.addWidget(log_label)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text, 1)

        button_layout = QHBoxLayout()

        self.btn_migrate = QPushButton("마이그레이션 실행")
        self.btn_migrate.clicked.connect(self.run_migration)
        button_layout.addWidget(self.btn_migrate)

        self.btn_close = QPushButton("닫기")
        self.btn_close.clicked.connect(self.close)
        button_layout.addWidget(self.btn_close)

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def log(self, message: str) -> None:
        self.log_text.append(message)
        QApplication.processEvents()

    def run_migration(self) -> None:
        if not DB_PATH.exists():
            QMessageBox.critical(self, "오류", f"데이터베이스 파일을 찾을 수 없습니다:\n{DB_PATH}")
            return

        reply = QMessageBox.question(
            self, "확인",
            "스키마 변경을 실행하시겠습니까?\n\n"
            "주의: 이 작업은 되돌릴 수 없습니다.\n"
            "데이터베이스 백업을 권장합니다.",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.log_text.clear()
        self.btn_migrate.setEnabled(False)

        try:
            # 1) sap.remark
            self.log("=" * 60)
            self.log("1) sap 테이블에 remark 컬럼 추가 중...")
            self.log(add_remark_to_sap())

            # 2) sap.renault_code
            self.log("=" * 60)
            self.log("2) sap 테이블에 renault_code 컬럼 추가 중...")
            self.log(add_renault_code_to_sap())

            # 3) rule_*.note
            self.log("=" * 60)
            self.log("3) rule_* 테이블들에 note 컬럼 추가 중...")

            rule_tables = get_all_rule_tables()
            if not rule_tables:
                self.log("rule_* 테이블이 없습니다.")
            else:
                ok_cnt = 0
                fail_cnt = 0
                for table_name in rule_tables:
                    self.log(f"  - {table_name} 처리 중...")
                    try:
                        res = add_note_to_rule_table(table_name)
                        self.log(f"    {res}")
                        ok_cnt += 1
                    except Exception as e:
                        # ✅ 테이블 하나 실패해도 전체 중단하지 않음
                        self.log(f"    ✗ 실패: {table_name} - {e}")
                        fail_cnt += 1

                self.log(f"\n[요약] rule_* 처리: 성공 {ok_cnt} / 실패 {fail_cnt}")

            self.log("=" * 60)
            self.log("마이그레이션 완료!")
            QMessageBox.information(self, "완료", "스키마 마이그레이션이 완료되었습니다.")
        except Exception as e:
            self.log(f"오류 발생: {str(e)}")
            QMessageBox.critical(self, "오류", f"마이그레이션 실패:\n{str(e)}")
        finally:
            self.btn_migrate.setEnabled(True)


def main():
    app = QApplication(sys.argv)
    window = SchemaMigrationWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
