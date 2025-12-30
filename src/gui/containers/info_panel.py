"""
정보 패널 컨테이너 - 비고 + Rule 링크
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QFrame
)


class InfoPanel(QWidget):
    """정보 패널 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        outer_layout = QVBoxLayout()
        info_group = QGroupBox("정보")

        # ===== 메인 가로 레이아웃 =====
        main_layout = QHBoxLayout()
        main_layout.setSpacing(20)   # 좌/우 영역 간격

        # ---------------- Rule 영역 (왼쪽) ----------------
        self.lbl_rule_title = QLabel("변경가능(rule)")
        self.lbl_rule_title.setStyleSheet("""
            QLabel {
                font-weight: 600;
                color: #555;
            }
        """)

        self.lbl_editable = QLabel("-")
        self.lbl_editable.setWordWrap(True)
        self.lbl_editable.setCursor(Qt.PointingHandCursor)
        self.lbl_editable.setStyleSheet("""
            QLabel {
                color: #1E6EDB;
                text-decoration: underline;
                line-height: 140%;
            }
        """)

        rule_layout = QVBoxLayout()
        rule_layout.setSpacing(6)   # 제목-내용 간격
        rule_layout.addWidget(self.lbl_rule_title)
        rule_layout.addWidget(self.lbl_editable)


        # ---------------- 가운데 구분선 ----------------
        line = QFrame()
        line.setFrameShape(QFrame.VLine)
        line.setFrameShadow(QFrame.Sunken)


        # ---------------- Remark 영역 (오른쪽) ----------------
        self.lbl_remark_title = QLabel("비고(remark)")
        self.lbl_remark_title.setStyleSheet("""
            QLabel {
                font-weight: 600;
                color: #555;
            }
        """)

        self.lbl_remark = QLabel("-")
        self.lbl_remark.setWordWrap(True)
        self.lbl_remark.setStyleSheet("""
            QLabel {
                line-height: 140%;
            }
        """)

        remark_layout = QVBoxLayout()
        remark_layout.setSpacing(6)
        remark_layout.addWidget(self.lbl_remark_title)
        remark_layout.addWidget(self.lbl_remark)


        # ---------------- main layout 구성 ----------------
        main_layout.addLayout(rule_layout)
        main_layout.addWidget(line)
        main_layout.addLayout(remark_layout)

        # 비율 지정 (rule : remark = 3 : 2)
        main_layout.setStretch(0, 3)
        main_layout.setStretch(1, 0)
        main_layout.setStretch(2, 2)

        info_group.setLayout(main_layout)
        outer_layout.addWidget(info_group)
        self.setLayout(outer_layout)


    # ===== 기존 함수 유지 =====
    def set_remark(self, text: str):
        """비고 설정"""
        self.lbl_remark.setText(text or "-")

    def set_editable(self, text: str):
        """Rule 텍스트 설정"""
        self.lbl_editable.setText(text or "-")

    def get_editable_label(self) -> QLabel:
        """클릭 이벤트 연결용"""
        return self.lbl_editable
