from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QTextEdit, QPushButton,
    QScrollArea
)
from src.database import update_company_remark


class InfoPanel(QWidget):
    """정보 패널 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)

        # 최소 높이 설정
        self.setMinimumHeight(50)

        self.current_sap_code: str | None = None
        self.original_remark: str = ""

        # ===== 외부 레이아웃 =====
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        # ===== GroupBox =====
        # GroupBox 제목을 빈 문자열로 설정하고, 커스텀 헤더 사용
        info_group = QGroupBox("")
        info_group.setStyleSheet("""
            QGroupBox {

                padding-top: 10px;
            }
        """)
        
        # GroupBox 내부 전체 레이아웃
        group_inner_layout = QVBoxLayout()
        group_inner_layout.setContentsMargins(0, 0, 0, 0)
        group_inner_layout.setSpacing(0)
        
        # COMEX 제목과 기업명을 표시할 헤더 위젯 (GroupBox 내부 상단)
        header_widget = QWidget()
        header_widget.setStyleSheet("background: transparent;")
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(4, 0, 4, 0)
        header_layout.setSpacing(10)
        
        # COMEX 레이블
        lbl_comx = QLabel("COMEX")
        lbl_comx.setStyleSheet("font-size: 12pt; color: black; font-weight: bold;")
        
        # 기업명 표시 레이블 (초기에는 숨김)
        self.lbl_company_header = QLabel("")
        self.lbl_company_header.setStyleSheet("font-size: 12pt; color: black; font-weight: normal;")
        
        header_layout.addWidget(lbl_comx)
        header_layout.addWidget(self.lbl_company_header)
        header_layout.addStretch()
        
        group_inner_layout.addWidget(header_widget)
        
        main_layout = QHBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(4, 8, 4, 8)

        # ================= Rule 카드 =================
        rule_card = QWidget()
        rule_card.setStyleSheet("background: #FAFAFA;")
        rule_card_layout = QVBoxLayout(rule_card)
        rule_card_layout.setContentsMargins(4, 8, 4, 8)
        rule_card_layout.setSpacing(6)

        # 1) 적용 규칙 텍스트 - 상단
        lbl_rule_title = QLabel("적용 규칙")
        lbl_rule_title.setStyleSheet("font-weight: bold; font-size: 10pt; color: #555;")
        rule_card_layout.addWidget(lbl_rule_title)

        # 2) Rule 목록을 스크롤 가능한 영역으로 만들기
        rule_scroll = QScrollArea()
        rule_scroll.setWidgetResizable(True)
        rule_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        rule_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        rule_scroll.setStyleSheet("QScrollArea { border: none; background: #FAFAFA; }")
        
        rule_content = QWidget()
        rule_layout = QVBoxLayout(rule_content)
        rule_layout.setContentsMargins(0, 0, 0, 0)
        rule_layout.setSpacing(4)
        
        # Rule 목록 레이아웃 (기업명은 COMEX 우측에 표시되므로 제거)
        self.rule_list_layout = QVBoxLayout()
        self.rule_list_layout.setContentsMargins(0, 0, 0, 0)
        self.rule_list_layout.setSpacing(4)
        rule_layout.addLayout(self.rule_list_layout)
        rule_layout.addStretch()
        
        rule_scroll.setWidget(rule_content)
        rule_card_layout.addWidget(rule_scroll)

        # ================= Remark 카드 =================
        remark_card = QWidget()
        remark_card.setStyleSheet("background: #FAFAFA;")
        remark_layout = QVBoxLayout(remark_card)
        remark_layout.setContentsMargins(4, 8, 4, 8)
        remark_layout.setSpacing(6)

        # 비고(Remark) 텍스트와 저장 버튼을 같은 줄에 배치 (우측 상단)
        remark_title_widget = QWidget()
        remark_title_layout = QHBoxLayout(remark_title_widget)
        remark_title_layout.setContentsMargins(0, 0, 0, 0)
        remark_title_layout.setSpacing(0)
        
        lbl_remark_title = QLabel("비고(Remark)")
        lbl_remark_title.setStyleSheet("font-weight: bold; font-size: 10pt; color: #555;")
        remark_title_layout.addWidget(lbl_remark_title)
        remark_title_layout.addStretch()
        
        self.btn_save_remark = QPushButton("저장")
        self.btn_save_remark.setEnabled(False)
        # 저장 버튼 높이를 제목 텍스트 높이와 맞추기
        self.btn_save_remark.setFixedHeight(20)
        remark_title_layout.addWidget(self.btn_save_remark)
        
        # 제목 위젯의 높이를 고정하여 적용 규칙 제목과 동일한 높이 유지
        remark_title_widget.setFixedHeight(20)
        remark_layout.addWidget(remark_title_widget)
        
        # 빈 공간 제거 (기업명 레이블이 없으므로 불필요)
        
        # Remark 텍스트 영역 (동적으로 크기 조정)
        self.remark_text = QTextEdit()
        # 고정 높이 제거하여 정보 섹션이 움직일 때 따라가도록 함
        remark_layout.addWidget(self.remark_text)
        remark_layout.addStretch()  # 하단 여백 추가

        # ================= 배치 =================
        main_layout.addWidget(rule_card)
        main_layout.addWidget(remark_card)
        main_layout.setStretch(0, 1)  # 룰 섹션: 1
        main_layout.setStretch(1, 1)  # 리마크 섹션: 1 (반반)

        group_inner_layout.addLayout(main_layout)
        info_group.setLayout(group_inner_layout)
        
        outer_layout.addWidget(info_group)

        # ===== 이벤트 연결 =====
        self.remark_text.textChanged.connect(self._on_remark_changed)
        self.btn_save_remark.clicked.connect(self._on_save_remark)

    # ================= 외부 호출용 =================
    def set_company_info(self, name: str, code: str):
        if name and code:
            self.current_sap_code = code
            # COMEX 제목 옆의 기업명 표시
            self.lbl_company_header.setText(f"{name} ({code})")
            self.lbl_company_header.setVisible(True)
        else:
            self.current_sap_code = None
            self.lbl_company_header.setText("")
            self.lbl_company_header.setVisible(False)

    def set_remark(self, remark: str):
        """main_page에서 회사 선택 시 호출"""
        self.original_remark = remark or ""
        self.remark_text.setText(self.original_remark)
        self.btn_save_remark.setEnabled(False)

    def set_rules(self, rules: list[dict]):
        """기업 선택 시 Rule 목록 표시"""
        self._clear_rule_list()

        if not rules:
            lbl = QLabel("등록된 Rule 없음")
            lbl.setStyleSheet("color: #999; font-size: 10pt;")
            self.rule_list_layout.addWidget(lbl)
            return

        # 우선순위 오름차순 정렬
        sorted_rules = sorted(rules, key=lambda r: r.get("priority", 999))
        for rule in sorted_rules:
            status = rule.get("status", "")
            changes = self._format_rule_changes(rule)
            text = f"{status} | {changes}"  # 우선순위 표시 제거

            lbl = QLabel(text)
            lbl.setWordWrap(True)
            lbl.setStyleSheet("font-size: 10pt;")

            if status.upper() == "ACTIVE":
                lbl.setStyleSheet("font-size: 10pt; color: #2E7D32;")
            elif status.upper() == "INACTIVE":
                lbl.setStyleSheet("font-size: 10pt; color: #888;")

            self.rule_list_layout.addWidget(lbl)

    # ================= 내부 로직 =================
    def _on_remark_changed(self):
        if not self.current_sap_code:
            self.btn_save_remark.setEnabled(False)
            return
        current = self.remark_text.toPlainText()
        self.btn_save_remark.setEnabled(current != self.original_remark)

    def _on_save_remark(self):
        if not self.current_sap_code:
            return
        new_remark = self.remark_text.toPlainText()
        success = update_company_remark(self.current_sap_code, new_remark)
        if success:
            self.original_remark = new_remark
            self.btn_save_remark.setEnabled(False)

    def _clear_rule_list(self):
        while self.rule_list_layout.count():
            item = self.rule_list_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    # ===== dialogs.py 의 format_rule_changes 로직 그대로 =====
    def _format_rule_changes(self, rule: dict) -> str:
        changes = []

        def valid(val, ignore=("ALL", "NONE")):
            return val and str(val).strip().upper() not in ignore

        if valid(rule.get("repair_region")):
            changes.append(f"수리지역: {rule['repair_region']}")
        if valid(rule.get("project_code")):
            changes.append(f"프로젝트: {rule['project_code']}")
        if rule.get("exclude_project_code"):
            changes.append(f"제외: {rule['exclude_project_code']}")
        if valid(rule.get("vehicle_classification")):
            changes.append(f"차계: {rule['vehicle_classification']}")
        if valid(rule.get("part_name")):
            changes.append(f"부품: {rule['part_name']}")
        if valid(rule.get("part_no")):
            changes.append(f"부품번호: {rule['part_no']}")
        if valid(rule.get("engine_form")):
            changes.append(f"엔진: {rule['engine_form']}")
        if rule.get("liability_ratio") is not None:
            changes.append(f"구상율: {rule['liability_ratio']}%")
        if rule.get("warranty_mileage_override") is not None:
            changes.append(f"주행거리: {rule['warranty_mileage_override']}km")
        if rule.get("warranty_period_override") is not None:
            years = rule["warranty_period_override"] / 365.0
            changes.append(f"보증기간: {years:.1f}년")
        if rule.get("amount_cap_value") is not None and valid(rule.get("amount_cap_type")):
            changes.append(f"상한: {rule['amount_cap_value']} ({rule['amount_cap_type']})")
        if rule.get("valid_from"):
            changes.append(f"시작일: {rule['valid_from']}")
        if rule.get("valid_to"):
            changes.append(f"종료일: {rule['valid_to']}")

        return " | ".join(changes) if changes else "기본 규칙"
