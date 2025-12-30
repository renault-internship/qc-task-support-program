"""
정보 패널 컨테이너 - 기업명, 비고, Rule 정보
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout,
    QLabel, QPushButton
)


class InfoPanel(QWidget):
    """정보 패널 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QVBoxLayout()
        
        info_group = QGroupBox("정보")
        form = QFormLayout()

        self.lbl_company = QLabel("-")
        self.lbl_remark = QLabel("-")
        self.lbl_editable = QLabel("-")
        
        # 규칙 추가 버튼
        self.btn_add_rule = QPushButton("+ 규칙 추가")
        self.btn_add_rule.setToolTip("규칙 추가")

        for lbl in (self.lbl_company, self.lbl_remark, self.lbl_editable):
            lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        
        # lbl_editable 클릭 시 Rule 목록 보기 (스타일만 설정, 이벤트는 외부에서 연결)
        self.lbl_editable.setCursor(Qt.CursorShape.PointingHandCursor)
        self.lbl_editable.setStyleSheet("QLabel { color: blue; text-decoration: underline; }")

        form.addRow("기업명", self.lbl_company)
        form.addRow("비고(remark)", self.lbl_remark)
        form.addRow("변경가능(rule)", self.lbl_editable)
        form.addRow("", self.btn_add_rule)
        info_group.setLayout(form)
        
        layout.addWidget(info_group)
        self.setLayout(layout)
    
    def set_company(self, text: str):
        """기업명 설정"""
        self.lbl_company.setText(text)
    
    def set_remark(self, text: str):
        """비고 설정"""
        self.lbl_remark.setText(text)
    
    def set_editable(self, text: str):
        """변경가능(rule) 설정"""
        self.lbl_editable.setText(text)
    
    def get_add_rule_button(self) -> QPushButton:
        """규칙 추가 버튼 반환"""
        return self.btn_add_rule
    
    def get_editable_label(self) -> QLabel:
        """변경가능 라벨 반환 (클릭 이벤트 연결용)"""
        return self.lbl_editable

