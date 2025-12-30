"""
Rule 관련 다이얼로그
"""
from typing import Dict, Any, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QFormLayout, QHBoxLayout, QVBoxLayout,
    QPushButton, QSpinBox, QDoubleSpinBox, QLineEdit,
    QComboBox, QTableWidget, QTableWidgetItem
)


class AddRuleDialog(QDialog):
    """Rule 추가/수정 다이얼로그"""
    def __init__(self, rule_table_name: str, parent=None, rule_data: Dict[str, Any] = None):
        super().__init__(parent)
        self.rule_table_name = rule_table_name
        self.rule_data = rule_data
        self.is_edit_mode = rule_data is not None
        
        title = f"규칙 수정 - {rule_table_name}" if self.is_edit_mode else f"규칙 추가 - {rule_table_name}"
        self.setWindowTitle(title)
        self.setFixedSize(500, 600)
        
        layout = QFormLayout()
        
        # Priority
        self.priority_spin = QSpinBox()
        self.priority_spin.setRange(1, 999)
        self.priority_spin.setValue(100)
        layout.addRow("우선순위 *:", self.priority_spin)
        
        # Status
        self.status_combo = QComboBox()
        self.status_combo.addItems(["ACTIVE", "INACTIVE"])
        layout.addRow("상태 *:", self.status_combo)
        
        # Repair Region
        self.repair_region_edit = QLineEdit()
        layout.addRow("수리 지역 *:", self.repair_region_edit)
        
        # Vehicle Classification
        self.vehicle_class_edit = QLineEdit()
        layout.addRow("차량 분류 *:", self.vehicle_class_edit)
        
        # Project Code (선택)
        self.project_code_edit = QLineEdit()
        self.project_code_edit.setPlaceholderText("비워두면 모든 프로젝트")
        layout.addRow("프로젝트 코드:", self.project_code_edit)
        
        # Part Name (선택)
        self.part_name_edit = QLineEdit()
        self.part_name_edit.setPlaceholderText("비워두면 모든 부품")
        layout.addRow("부품명:", self.part_name_edit)
        
        # Part No (선택)
        self.part_no_edit = QLineEdit()
        self.part_no_edit.setPlaceholderText("선택사항")
        layout.addRow("부품 번호:", self.part_no_edit)
        
        # Liability Ratio (필수)
        self.liability_ratio_spin = QDoubleSpinBox()
        self.liability_ratio_spin.setRange(0.0, 100.0)
        self.liability_ratio_spin.setDecimals(2)
        self.liability_ratio_spin.setSuffix(" %")
        self.liability_ratio_spin.setValue(0.0)
        layout.addRow("구상율 *:", self.liability_ratio_spin)
        
        # Amount Cap Type
        self.amount_cap_combo = QComboBox()
        self.amount_cap_combo.addItems(["NONE", "FIXED", "PERCENTAGE"])
        layout.addRow("금액 상한 타입 *:", self.amount_cap_combo)
        
        # Amount Cap Value
        self.amount_cap_spin = QSpinBox()
        self.amount_cap_spin.setRange(0, 999999999)
        self.amount_cap_spin.setValue(0)
        layout.addRow("금액 상한 값:", self.amount_cap_spin)
        
        # Exclude Project Code
        self.exclude_project_code_edit = QLineEdit()
        self.exclude_project_code_edit.setPlaceholderText("제외할 프로젝트 코드")
        layout.addRow("제외 프로젝트 코드:", self.exclude_project_code_edit)
        
        # Note
        self.note_edit = QLineEdit()
        layout.addRow("비고:", self.note_edit)
        
        # 버튼
        button_layout = QHBoxLayout()
        self.save_btn = QPushButton("저장")
        self.save_btn.clicked.connect(self.accept)
        self.cancel_btn = QPushButton("취소")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addRow(button_layout)
        
        self.setLayout(layout)
        
        # 수정 모드인 경우 기존 데이터로 채우기
        if self.is_edit_mode and rule_data:
            self._load_rule_data(rule_data)
    
    def _load_rule_data(self, rule_data: Dict[str, Any]):
        """기존 규칙 데이터로 폼 채우기"""
        if "priority" in rule_data:
            self.priority_spin.setValue(rule_data["priority"])
        if "status" in rule_data:
            idx = self.status_combo.findText(rule_data["status"])
            if idx >= 0:
                self.status_combo.setCurrentIndex(idx)
        if "repair_region" in rule_data:
            self.repair_region_edit.setText(str(rule_data["repair_region"]))
        if "vehicle_classification" in rule_data:
            self.vehicle_class_edit.setText(str(rule_data["vehicle_classification"]))
        if "project_code" in rule_data:
            self.project_code_edit.setText(str(rule_data["project_code"]))
        if "part_name" in rule_data:
            self.part_name_edit.setText(str(rule_data["part_name"]))
        if "part_no" in rule_data:
            self.part_no_edit.setText(str(rule_data["part_no"]))
        if "liability_ratio" in rule_data:
            self.liability_ratio_spin.setValue(float(rule_data["liability_ratio"]))
        if "amount_cap_type" in rule_data:
            idx = self.amount_cap_combo.findText(rule_data["amount_cap_type"])
            if idx >= 0:
                self.amount_cap_combo.setCurrentIndex(idx)
        if "amount_cap_value" in rule_data and rule_data["amount_cap_value"]:
            self.amount_cap_spin.setValue(int(rule_data["amount_cap_value"]))
        if "exclude_project_code" in rule_data:
            self.exclude_project_code_edit.setText(str(rule_data["exclude_project_code"]))
        if "note" in rule_data:
            self.note_edit.setText(str(rule_data["note"]))
    
    def get_data(self) -> Dict[str, Any]:
        """입력된 데이터 반환"""
        return {
            "priority": self.priority_spin.value(),
            "status": self.status_combo.currentText(),
            "repair_region": self.repair_region_edit.text().strip(),
            "vehicle_classification": self.vehicle_class_edit.text().strip(),
            "project_code": self.project_code_edit.text().strip(),
            "part_name": self.part_name_edit.text().strip(),
            "part_no": self.part_no_edit.text().strip(),
            "liability_ratio": self.liability_ratio_spin.value(),
            "amount_cap_type": self.amount_cap_combo.currentText(),
            "amount_cap_value": self.amount_cap_spin.value() if self.amount_cap_spin.value() > 0 else None,
            "exclude_project_code": self.exclude_project_code_edit.text().strip(),
            "note": self.note_edit.text().strip(),
        }


class ViewRulesDialog(QDialog):
    """Rule 목록 보기 다이얼로그 (변경점만 표시)"""
    def __init__(self, rules: List[Dict[str, Any]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rule 목록")
        self.setFixedSize(800, 500)
        
        layout = QVBoxLayout()
        
        # Rule 목록 테이블
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["우선순위", "상태", "변경점"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        
        # Rule 데이터 채우기
        self.populate_rules(rules)
        
        layout.addWidget(self.table)
        
        # 닫기 버튼
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def format_rule_changes(self, rule: Dict[str, Any]) -> str:
        """Rule의 변경점만 포맷팅하여 반환"""
        changes = []
        
        # 프로젝트 코드
        project_code = rule.get("project_code", "").strip()
        if project_code:
            changes.append(f"프로젝트: {project_code}")
        
        # 부품명
        part_name = rule.get("part_name", "").strip()
        if part_name:
            changes.append(f"부품: {part_name}")
        
        # 부품 번호
        part_no = rule.get("part_no", "").strip()
        if part_no:
            changes.append(f"부품번호: {part_no}")
        
        # 제외 프로젝트
        exclude_project = rule.get("exclude_project_code", "").strip()
        if exclude_project:
            changes.append(f"제외: {exclude_project}")
        
        # 구상율 (항상 표시)
        liability_ratio = rule.get("liability_ratio", 0)
        if liability_ratio is not None:
            changes.append(f"구상율: {liability_ratio}%")
        
        # 보증 주행거리 오버라이드
        warranty_mileage = rule.get("warranty_mileage_override")
        if warranty_mileage:
            changes.append(f"주행거리: {warranty_mileage}km")
        
        # 보증 기간 오버라이드
        warranty_period = rule.get("warranty_period_override")
        if warranty_period:
            years = warranty_period / 365.0
            changes.append(f"보증기간: {years:.1f}년")
        
        # 금액 상한
        amount_cap_value = rule.get("amount_cap_value")
        if amount_cap_value:
            cap_type = rule.get("amount_cap_type", "NONE")
            if cap_type != "NONE":
                changes.append(f"상한: {amount_cap_value} ({cap_type})")
        
        # 비고
        note = rule.get("note", "").strip()
        if note:
            changes.append(f"비고: {note}")
        
        return " | ".join(changes) if changes else "기본 규칙"
    
    def populate_rules(self, rules: List[Dict[str, Any]]):
        """Rule 목록을 테이블에 채우기"""
        self.table.setRowCount(len(rules))
        
        for row, rule in enumerate(rules):
            # 우선순위
            priority_item = QTableWidgetItem(str(rule.get("priority", "")))
            priority_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, priority_item)
            
            # 상태
            status = rule.get("status", "")
            status_item = QTableWidgetItem(status)
            status_item.setTextAlignment(Qt.AlignCenter)
            # ACTIVE는 초록색, INACTIVE는 회색으로 표시
            if status.upper() == "ACTIVE":
                status_item.setForeground(Qt.GlobalColor.green)
            elif status.upper() == "INACTIVE":
                status_item.setForeground(Qt.GlobalColor.gray)
            self.table.setItem(row, 1, status_item)
            
            # 변경점
            changes_item = QTableWidgetItem(self.format_rule_changes(rule))
            self.table.setItem(row, 2, changes_item)
        
        self.table.resizeColumnsToContents()

