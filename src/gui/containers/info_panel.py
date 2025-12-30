"""
정보 패널 컨테이너 - 비고 + Rule 링크
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout,
    QLabel
)


class InfoPanel(QWidget):
    """정보 패널 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QVBoxLayout()
        
        info_group = QGroupBox("정보")
        form = QFormLayout()

        # 비고
        self.lbl_remark = QLabel("-")
        self.lbl_remark.setTextInteractionFlags(Qt.TextSelectableByMouse)

        # Rule (클릭 가능 형태)
        self.lbl_editable = QLabel("-")
        self.lbl_editable.setTextInteractionFlags(Qt.TextSelectableByMouse)

        # 파란색 링크 스타일 + 포인터 커서
        self.lbl_editable.setCursor(Qt.PointingHandCursor)
        self.lbl_editable.setStyleSheet(
            "QLabel { color: blue; text-decoration: underline; }"
        )

        # 레이아웃 구성
        form.addRow("비고(remark)", self.lbl_remark)
        form.addRow("변경가능(rule)", self.lbl_editable)

        info_group.setLayout(form)
        layout.addWidget(info_group)
        self.setLayout(layout)
    
    def set_remark(self, text: str):
        """비고 설정"""
        self.lbl_remark.setText(text)

    def set_editable(self, text: str):
        """Rule 텍스트 설정"""
        self.lbl_editable.setText(text)

    def get_editable_label(self) -> QLabel:
        """클릭 이벤트 연결용"""
        return self.lbl_editable
