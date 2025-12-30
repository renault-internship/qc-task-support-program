"""
하단 Export 패널 - export 버튼들
"""
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QPushButton
)


class ExportPanel(QWidget):
    """하단 Export 패널"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QHBoxLayout()
        layout.addStretch()
        
        self.btn_export_rule = QPushButton("export (rule)")
        self.btn_export_final = QPushButton("export (최종 엑셀)")
        
        layout.addWidget(self.btn_export_rule)
        layout.addWidget(self.btn_export_final)
        
        self.setLayout(layout)
    
    def get_export_rule_button(self) -> QPushButton:
        """export (rule) 버튼 반환"""
        return self.btn_export_rule
    
    def get_export_final_button(self) -> QPushButton:
        """export (최종 엑셀) 버튼 반환"""
        return self.btn_export_final

