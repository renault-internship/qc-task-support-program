"""
DB 뷰어 - TestDB 구조 및 데이터 확인 (독립 실행)
사용법: python db_viewer.py
"""
import sys
from pathlib import Path
import sqlite3
from typing import List, Dict, Any

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem,
    QLabel, QGroupBox, QMessageBox
)

# 데이터베이스 경로
DB_PATH = Path("data/TestDB.sqlite")


def get_all_tables() -> List[str]:
    """모든 테이블 목록 조회"""
    if not DB_PATH.exists():
        return []
    
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT name FROM sqlite_master 
        WHERE type='table' AND name NOT LIKE 'sqlite_%'
        ORDER BY name
    """)
    
    rows = cursor.fetchall()
    conn.close()
    
    return [row[0] for row in rows] if rows else []


def get_table_schema(table_name: str) -> List[Dict[str, Any]]:
    """테이블의 스키마 정보 조회"""
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    
    cursor.execute(f"PRAGMA table_info({table_name})")
    rows = cursor.fetchall()
    conn.close()
    
    # 컬럼 정보: (cid, name, type, notnull, default_value, pk)
    return [
        {
            "cid": row[0],
            "name": row[1],
            "type": row[2],
            "notnull": bool(row[3]),
            "default_value": row[4],
            "pk": bool(row[5]),
        }
        for row in rows
    ]


def get_table_data(table_name: str, limit: int = 1000) -> List[Dict[str, Any]]:
    """테이블의 데이터 조회"""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    try:
        cursor.execute(f'SELECT * FROM "{table_name}" LIMIT ?', (limit,))
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows] if rows else []
    except sqlite3.OperationalError as e:
        conn.close()
        raise ValueError(f"데이터 조회 실패: {str(e)}")


class DBViewerWindow(QWidget):
    """DB 뷰어 윈도우"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DB 뷰어 - TestDB")
        self.resize(1200, 800)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        
        # 왼쪽: 테이블 목록
        left_panel = QVBoxLayout()
        left_panel.addWidget(QLabel("테이블 목록"))
        
        self.table_list = QListWidget()
        self.table_list.setMaximumWidth(200)
        self.table_list.itemClicked.connect(self.on_table_selected)
        left_panel.addWidget(self.table_list)
        
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        layout.addWidget(left_widget)
        
        # 오른쪽: 스키마 + 데이터
        right_layout = QVBoxLayout()
        
        # 스키마 정보
        schema_group = QGroupBox("테이블 구조")
        schema_layout = QVBoxLayout()
        self.schema_table = QTableWidget()
        self.schema_table.setColumnCount(6)
        self.schema_table.setHorizontalHeaderLabels([
            "CID", "컬럼명", "타입", "NOT NULL", "기본값", "PK"
        ])
        self.schema_table.horizontalHeader().setStretchLastSection(True)
        self.schema_table.setAlternatingRowColors(True)
        schema_layout.addWidget(self.schema_table)
        schema_group.setLayout(schema_layout)
        right_layout.addWidget(schema_group)
        
        # 데이터
        data_group = QGroupBox("데이터 (최대 1000행)")
        data_layout = QVBoxLayout()
        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        self.data_table.horizontalHeader().setStretchLastSection(True)
        data_layout.addWidget(self.data_table)
        data_group.setLayout(data_layout)
        right_layout.addWidget(data_group, 1)
        
        right_widget = QWidget()
        right_widget.setLayout(right_layout)
        layout.addWidget(right_widget, 1)
        
        self.setLayout(layout)
        
        # 초기화
        self.load_tables()
    
    def load_tables(self):
        """테이블 목록 로드"""
        self.table_list.clear()
        
        if not DB_PATH.exists():
            QMessageBox.warning(self, "오류", f"데이터베이스 파일을 찾을 수 없습니다:\n{DB_PATH}")
            return
        
        tables = get_all_tables()
        
        if not tables:
            QMessageBox.information(self, "안내", "테이블이 없습니다.")
            return
        
        for table in tables:
            item = QListWidgetItem(table)
            self.table_list.addItem(item)
        
        # 첫 번째 테이블 자동 선택
        if tables:
            self.table_list.setCurrentRow(0)
            self.on_table_selected(self.table_list.item(0))
    
    def on_table_selected(self, item: QListWidgetItem):
        """테이블 선택 시 스키마와 데이터 로드"""
        if not item:
            return
        
        table_name = item.text()
        
        # 스키마 로드
        try:
            schema = get_table_schema(table_name)
            self.load_schema(schema)
        except Exception as e:
            self.schema_table.setRowCount(0)
            QMessageBox.warning(self, "오류", f"스키마 로드 실패: {str(e)}")
        
        # 데이터 로드
        try:
            data = get_table_data(table_name)
            self.load_data(data)
        except Exception as e:
            self.data_table.setRowCount(0)
            QMessageBox.warning(self, "오류", f"데이터 로드 실패: {str(e)}")
    
    def load_schema(self, schema: List[Dict[str, Any]]):
        """스키마 테이블에 표시"""
        self.schema_table.setRowCount(len(schema))
        
        for row, col_info in enumerate(schema):
            self.schema_table.setItem(row, 0, QTableWidgetItem(str(col_info["cid"])))
            self.schema_table.setItem(row, 1, QTableWidgetItem(col_info["name"]))
            self.schema_table.setItem(row, 2, QTableWidgetItem(col_info["type"]))
            self.schema_table.setItem(row, 3, QTableWidgetItem("YES" if col_info["notnull"] else "NO"))
            default_val = str(col_info["default_value"]) if col_info["default_value"] is not None else ""
            self.schema_table.setItem(row, 4, QTableWidgetItem(default_val))
            self.schema_table.setItem(row, 5, QTableWidgetItem("YES" if col_info["pk"] else "NO"))
        
        self.schema_table.resizeColumnsToContents()
    
    def load_data(self, data: List[Dict[str, Any]]):
        """데이터 테이블에 표시"""
        if not data:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            return
        
        # 컬럼명 추출
        columns = list(data[0].keys())
        self.data_table.setColumnCount(len(columns))
        self.data_table.setHorizontalHeaderLabels(columns)
        
        # 데이터 채우기
        self.data_table.setRowCount(len(data))
        
        for row, record in enumerate(data):
            for col, col_name in enumerate(columns):
                value = record.get(col_name)
                item = QTableWidgetItem(str(value) if value is not None else "")
                self.data_table.setItem(row, col, item)
        
        self.data_table.resizeColumnsToContents()


def main():
    """메인 함수"""
    app = QApplication(sys.argv)
    window = DBViewerWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()