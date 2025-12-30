"""
DB 스키마 마이그레이션 도구 - 독립 실행
사용법: python db_schema_migrate.py

변경 사항:
1. sap 테이블에 remark 컬럼 추가
2. rule_* 테이블들에서 note 컬럼 삭제
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


def get_table_columns(table_name: str) -> List[str]:
    """테이블의 컬럼 목록 조회"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    cursor.execute(f"PRAGMA table_info({table_name})")
    rows = cursor.fetchall()
    conn.close()
    
    return [row[1] for row in rows]  # row[1]은 컬럼명


def add_remark_to_sap():
    """sap 테이블에 remark 컬럼 추가"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    try:
        # 이미 존재하는지 확인
        columns = get_table_columns("sap")
        if "remark" in columns:
            conn.close()
            return "sap 테이블에 remark 컬럼이 이미 존재합니다."
        
        # 컬럼 추가
        cursor.execute("ALTER TABLE sap ADD COLUMN remark TEXT DEFAULT ''")
        conn.commit()
        conn.close()
        return "sap 테이블에 remark 컬럼이 추가되었습니다."
    except sqlite3.OperationalError as e:
        conn.close()
        raise Exception(f"sap 테이블 컬럼 추가 실패: {str(e)}")


def drop_note_from_rule_table(table_name: str) -> str:
    """rule 테이블에서 note 컬럼 삭제 (SQLite는 DROP COLUMN을 직접 지원하지 않으므로 복잡한 과정 필요)"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    try:
        # 컬럼 목록 확인
        columns = get_table_columns(table_name)
        if "note" not in columns:
            conn.close()
            return f"{table_name}: note 컬럼이 존재하지 않습니다."
        
        # 1. 원본 테이블 구조 조회
        cursor.execute(f"PRAGMA table_info({table_name})")
        column_info = cursor.fetchall()
        
        # 2. 원본 CREATE TABLE 문 가져오기 (AUTOINCREMENT 등 제약조건 확인용)
        cursor.execute(f"SELECT sql FROM sqlite_master WHERE type='table' AND name='{table_name}'")
        create_sql_result = cursor.fetchone()
        original_sql = create_sql_result[0] if create_sql_result else ""
        has_autoincrement = "AUTOINCREMENT" in original_sql.upper()
        
        # 3. 원본 CREATE TABLE 문에서 note 컬럼만 제거
        cursor.execute(f"SELECT sql FROM sqlite_master WHERE type='table' AND name='{table_name}'")
        create_sql_result = cursor.fetchone()
        
        if not create_sql_result:
            conn.close()
            raise Exception(f"{table_name} 테이블의 CREATE 문을 찾을 수 없습니다.")
        
        original_sql = create_sql_result[0]
        
        # 원본 SQL을 줄 단위로 분리하여 note 컬럼 라인 제거
        lines = original_sql.split('\n')
        new_lines = []
        skip_next_empty = False
        
        for i, line in enumerate(lines):
            # note 컬럼 라인 찾기 (대소문자 구분 없이)
            line_stripped = line.strip()
            if re.match(r'^"?note"?\s+', line_stripped, re.IGNORECASE):
                # note 컬럼 라인은 제외
                skip_next_empty = True
                continue
            
            # note 컬럼 다음 빈 줄도 제거 (선택적)
            if skip_next_empty and not line_stripped:
                skip_next_empty = False
                continue
            
            skip_next_empty = False
            new_lines.append(line)
        
        new_sql = '\n'.join(new_lines)
        
        # 임시 테이블 생성
        temp_table = f"{table_name}_temp"
        # 테이블명만 변경
        new_sql = new_sql.replace(f'CREATE TABLE "{table_name}"', f'CREATE TABLE "{temp_table}"')
        new_sql = new_sql.replace(f"CREATE TABLE {table_name}", f"CREATE TABLE {temp_table}")
        
        # SQL 실행
        cursor.execute(new_sql)
        
        # 4. 데이터 복사 (note 제외)
        old_columns = [col[1] for col in column_info if col[1] != "note"]
        old_cols_str = ", ".join([f'"{col}"' for col in old_columns])
        
        cursor.execute(f"""
            INSERT INTO "{temp_table}" ({old_cols_str})
            SELECT {old_cols_str} FROM "{table_name}"
        """)
        
        # 5. 기존 테이블 삭제
        cursor.execute(f'DROP TABLE "{table_name}"')
        
        # 6. 새 테이블 이름 변경
        cursor.execute(f'ALTER TABLE "{temp_table}" RENAME TO "{table_name}"')
        
        conn.commit()
        conn.close()
        return f"{table_name}: note 컬럼이 삭제되었습니다."
    except sqlite3.OperationalError as e:
        conn.rollback()
        conn.close()
        raise Exception(f"{table_name} 테이블 컬럼 삭제 실패: {str(e)}")
    except Exception as e:
        conn.rollback()
        conn.close()
        raise Exception(f"{table_name} 테이블 컬럼 삭제 실패: {str(e)}")


class SchemaMigrationWindow(QWidget):
    """스키마 마이그레이션 윈도우"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DB 스키마 마이그레이션")
        self.setFixedSize(600, 500)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        # 설명
        info_label = QLabel(
            "변경 사항:\n"
            "1. sap 테이블에 remark 컬럼 추가\n"
            "2. rule_* 테이블들에서 note 컬럼 삭제"
        )
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # 로그 영역
        log_label = QLabel("실행 로그:")
        layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text, 1)
        
        # 버튼들
        button_layout = QHBoxLayout()
        
        self.btn_migrate = QPushButton("마이그레이션 실행")
        self.btn_migrate.clicked.connect(self.run_migration)
        button_layout.addWidget(self.btn_migrate)
        
        self.btn_close = QPushButton("닫기")
        self.btn_close.clicked.connect(self.close)
        button_layout.addWidget(self.btn_close)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def log(self, message: str):
        """로그 메시지 추가"""
        self.log_text.append(message)
        QApplication.processEvents()  # UI 업데이트
    
    def run_migration(self):
        """마이그레이션 실행"""
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
            # 1. sap 테이블에 remark 추가
            self.log("=" * 50)
            self.log("1. sap 테이블에 remark 컬럼 추가 중...")
            result = add_remark_to_sap()
            self.log(result)
            
            # 2. rule_* 테이블들에서 note 삭제
            self.log("=" * 50)
            self.log("2. rule_* 테이블들에서 note 컬럼 삭제 중...")
            
            rule_tables = get_all_rule_tables()
            if not rule_tables:
                self.log("rule_* 테이블이 없습니다.")
            else:
                for table_name in rule_tables:
                    self.log(f"  - {table_name} 처리 중...")
                    result = drop_note_from_rule_table(table_name)
                    self.log(f"    {result}")
            
            self.log("=" * 50)
            self.log("마이그레이션 완료!")
            
            QMessageBox.information(self, "완료", "스키마 마이그레이션이 완료되었습니다.")
        except Exception as e:
            self.log(f"오류 발생: {str(e)}")
            QMessageBox.critical(self, "오류", f"마이그레이션 실패:\n{str(e)}")
        finally:
            self.btn_migrate.setEnabled(True)


def main():
    """메인 함수"""
    app = QApplication(sys.argv)
    window = SchemaMigrationWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

