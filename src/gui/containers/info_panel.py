from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QTextEdit, QPushButton
)
from src.database import update_company_remark


class InfoPanel(QWidget):
    """정보 패널 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)

        self.current_sap_code: str | None = None
        self.original_remark: str = ""

        # ===== 외부 레이아웃 =====
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)

        # ===== GroupBox =====
        info_group = QGroupBox("정보")
        main_layout = QHBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(12, 8, 12, 8)

        # ================= Rule 카드 =================
        rule_card = QWidget()
        rule_card.setStyleSheet("background: #FAFAFA;")
        rule_layout = QVBoxLayout(rule_card)
        rule_layout.setContentsMargins(10, 8, 10, 8)
        rule_layout.setSpacing(6)

        # 1) 적용 Rule 텍스트 - 상단
        lbl_rule_title = QLabel("적용 Rule")
        lbl_rule_title.setStyleSheet("font-weight: 600; color: #555;")
        rule_layout.addWidget(lbl_rule_title)

        # 2) 회사명과 회사코드 표시
        self.lbl_company_info = QLabel("-")
        self.lbl_company_info.setStyleSheet("color: #777; font-size: 11px;")
        rule_layout.addWidget(self.lbl_company_info)

        # 3) Rule 목록 레이아웃
        self.rule_list_layout = QVBoxLayout()
        self.rule_list_layout.setSpacing(4)
        rule_layout.addLayout(self.rule_list_layout)

        rule_layout.addStretch()

        # ================= Remark 카드 =================
        remark_card = QWidget()
        remark_card.setStyleSheet("background: #FAFAFA;")
        remark_layout = QVBoxLayout(remark_card)
        remark_layout.setContentsMargins(10, 8, 10, 8)
        remark_layout.setSpacing(6)

        lbl_remark_title = QLabel("비고(remark)")
        lbl_remark_title.setStyleSheet("font-weight: 600; color: #555;")
        self.remark_text = QTextEdit()
        self.remark_text.setMaximumHeight(80)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        self.btn_save_remark = QPushButton("저장")
        self.btn_save_remark.setEnabled(False)
        btn_layout.addWidget(self.btn_save_remark)

        remark_layout.addWidget(lbl_remark_title)
        remark_layout.addWidget(self.remark_text)
        remark_layout.addLayout(btn_layout)

        # ================= 배치 =================
        main_layout.addWidget(rule_card)
        main_layout.addWidget(remark_card)
        main_layout.setStretch(0, 3)
        main_layout.setStretch(1, 2)

        info_group.setLayout(main_layout)
        outer_layout.addWidget(info_group)

        # ===== 이벤트 연결 =====
        self.remark_text.textChanged.connect(self._on_remark_changed)
        self.btn_save_remark.clicked.connect(self._on_save_remark)

    # ================= 외부 호출용 =================
    def set_company_info(self, name: str, code: str):
        if name and code:
            self.current_sap_code = code
            self.lbl_company_info.setText(f"{name} ({code})")
        else:
            self.current_sap_code = None
            self.lbl_company_info.setText("-")

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
            lbl.setStyleSheet("color: #999; font-size: 11px;")
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
            lbl.setStyleSheet("font-size: 11px;")

            if status.upper() == "ACTIVE":
                lbl.setStyleSheet("font-size: 11px; color: #2E7D32;")
            elif status.upper() == "INACTIVE":
                lbl.setStyleSheet("font-size: 11px; color: #888;")

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
