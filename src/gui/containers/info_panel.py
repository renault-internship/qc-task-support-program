from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel
)


class InfoPanel(QWidget):
    """정보 패널 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        # ===== GroupBox =====
        info_group = QGroupBox("정보")

        main_layout = QHBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(12, 8, 12, 8)


        # ================= Rule 카드 =================
        rule_card = QWidget()
        rule_card.setStyleSheet("""
            QWidget {
                background: #FAFAFA;
            }
        """)

        rule_layout = QVBoxLayout(rule_card)
        rule_layout.setContentsMargins(10, 8, 10, 8)
        rule_layout.setSpacing(6)

        lbl_rule_title = QLabel("변경가능(rule)")
        lbl_rule_title.setStyleSheet("font-weight: 600; color: #555;")

        # 회사명 + 코드 라벨 추가
        self.lbl_company_info = QLabel("-")
        self.lbl_company_info.setStyleSheet("""
            QLabel {
                color: #777;
                font-size: 11px;
            }
        """)

        self.lbl_editable = QLabel("-")
        self.lbl_editable.setWordWrap(True)
        self.lbl_editable.setCursor(Qt.PointingHandCursor)
        self.lbl_editable.setStyleSheet("""
            QLabel {
                color: #1E6EDB;
                text-decoration: underline;
            }
        """)

        rule_layout.addWidget(lbl_rule_title)
        rule_layout.addWidget(self.lbl_company_info)
        rule_layout.addWidget(self.lbl_editable)


        # ================= Remark 카드 =================
        remark_card = QWidget()
        remark_card.setStyleSheet("""
            QWidget {
                background: #FAFAFA;
            }
        """)

        remark_layout = QVBoxLayout(remark_card)
        remark_layout.setContentsMargins(10, 8, 10, 8)
        remark_layout.setSpacing(6)

        lbl_remark_title = QLabel("비고(remark)")
        lbl_remark_title.setStyleSheet("font-weight: 600; color: #555;")

        self.lbl_remark = QLabel("등록된 비고 없음")
        self.lbl_remark.setWordWrap(True)
        self.lbl_remark.setStyleSheet("color: #555;")

        remark_layout.addWidget(lbl_remark_title)
        remark_layout.addWidget(self.lbl_remark)


        # ================= 배치 =================
        main_layout.addWidget(rule_card)
        main_layout.addWidget(remark_card)

        # 카드 비율 (Rule 더 넓게)
        main_layout.setStretch(0, 3)
        main_layout.setStretch(1, 2)

        info_group.setLayout(main_layout)
        outer_layout.addWidget(info_group)


    # ===== 함수 유지 =====
    def set_remark(self, text: str):
        self.lbl_remark.setText(text if text else "등록된 비고 없음")

    def set_editable(self, text: str):
        self.lbl_editable.setText(text or "-")

    def get_editable_label(self):
        return self.lbl_editable

    def set_company_info(self, name: str, code: str):
        if name and code:
            self.lbl_company_info.setText(f"{name} ({code})")
        else:
            self.lbl_company_info.setText("-")
