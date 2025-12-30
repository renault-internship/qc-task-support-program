"""
미리보기 컨테이너 - 시트 선택 + 엑셀 테이블
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QTableView, QComboBox, QLabel
)


class PreviewContainer(QWidget):
    """엑셀 미리보기 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        
        # 시트 선택
        sheet_row = QHBoxLayout()
        sheet_row.addWidget(QLabel("시트"))
        self.sheet_combo = QComboBox()
        sheet_row.addWidget(self.sheet_combo, 1)
        layout.addLayout(sheet_row)
        
        # 엑셀 테이블
        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(False)
        layout.addWidget(self.table, 1)
        
        self.setLayout(layout)
    
    def get_sheet_combo(self) -> QComboBox:
        """시트 선택 콤보박스 반환"""
        return self.sheet_combo
    
    def get_table(self) -> QTableView:
        """엑셀 테이블 반환"""
        return self.table

