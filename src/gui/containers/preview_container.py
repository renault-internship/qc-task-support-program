"""
미리보기 컨테이너 - 엑셀 테이블
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout,
    QTableView
)


class PreviewContainer(QWidget):
    """엑셀 미리보기 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # 엑셀 테이블
        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(False)
        layout.addWidget(self.table, 1)
        
        self.setLayout(layout)
    
    def get_table(self) -> QTableView:
        """엑셀 테이블 반환"""
        return self.table

