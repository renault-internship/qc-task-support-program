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
        self.setFixedSize(500, 700)  # 높이 증가
        
        layout = QFormLayout()
        
        # Priority (DEFAULT -1, 트리거로 자동 채움)
        self.priority_spin = QSpinBox()
        self.priority_spin.setRange(-1, 999)
        self.priority_spin.setValue(-1)
        self.priority_spin.setSpecialValueText("자동 (트리거)")
        layout.addRow("우선순위:", self.priority_spin)
        
        # Status (DEFAULT 'ACTIVE', CHECK IN ('ACTIVE','INACTIVE'))
        self.status_combo = QComboBox()
        self.status_combo.addItems(["ACTIVE", "INACTIVE"])
        self.status_combo.setCurrentText("ACTIVE")  # 기본값
        layout.addRow("상태 *:", self.status_combo)
        
        # Repair Region (CHECK IN ('DOMESTIC','OVERSEAS','ALL'))
        self.repair_region_combo = QComboBox()
        self.repair_region_combo.addItems(["DOMESTIC", "OVERSEAS", "ALL"])
        self.repair_region_combo.setCurrentText("ALL")  # 기본값
        layout.addRow("수리 지역 *:", self.repair_region_combo)
        
        # Project Code (DEFAULT 'ALL')
        self.project_code_edit = QLineEdit()
        self.project_code_edit.setPlaceholderText("기본값: ALL")
        self.project_code_edit.setText("ALL")  # 기본값
        layout.addRow("프로젝트 코드 *:", self.project_code_edit)
        
        # Exclude Project Code (NULL 허용)
        self.exclude_project_code_edit = QLineEdit()
        self.exclude_project_code_edit.setPlaceholderText("제외할 프로젝트 코드 (선택사항)")
        layout.addRow("제외 프로젝트 코드:", self.exclude_project_code_edit)
        
        # Vehicle Classification (DEFAULT 'ALL')
        self.vehicle_class_edit = QLineEdit()
        self.vehicle_class_edit.setPlaceholderText("기본값: ALL")
        self.vehicle_class_edit.setText("ALL")  # 기본값
        layout.addRow("차량 분류 *:", self.vehicle_class_edit)
        
        # Part No (NOT NULL DEFAULT 'ALL')
        self.part_no_edit = QLineEdit()
        self.part_no_edit.setPlaceholderText("기본값: ALL")
        self.part_no_edit.setText("ALL")  # 기본값
        layout.addRow("부품 번호 *:", self.part_no_edit)
        
        # Part Name (NOT NULL DEFAULT 'ALL')
        self.part_name_edit = QLineEdit()
        self.part_name_edit.setPlaceholderText("기본값: ALL")
        self.part_name_edit.setText("ALL")  # 기본값
        layout.addRow("부품명 *:", self.part_name_edit)
        
        # Engine Form (NOT NULL DEFAULT 'ALL')
        self.engine_form_edit = QLineEdit()
        self.engine_form_edit.setPlaceholderText("기본값: ALL")
        self.engine_form_edit.setText("ALL")  # 기본값
        layout.addRow("엔진 형태 *:", self.engine_form_edit)
        
        # Warranty Mileage Override (NULL 허용)
        self.warranty_mileage_spin = QSpinBox()
        self.warranty_mileage_spin.setRange(0, 1000000)
        self.warranty_mileage_spin.setValue(0)
        self.warranty_mileage_spin.setSpecialValueText("없음")
        layout.addRow("보증 주행거리 오버라이드 (km):", self.warranty_mileage_spin)
        
        # Warranty Period Override (NULL 허용, 일 단위)
        self.warranty_period_spin = QSpinBox()
        self.warranty_period_spin.setRange(0, 3650)
        self.warranty_period_spin.setValue(0)
        self.warranty_period_spin.setSpecialValueText("없음")
        layout.addRow("보증 기간 오버라이드 (일):", self.warranty_period_spin)
        
        # Amount Cap Type (DEFAULT 'NONE', CHECK IN ('LABOR','OUTSOURCE_LABOR','BOTH_LABOR','NONE'))
        self.amount_cap_combo = QComboBox()
        self.amount_cap_combo.addItems(["NONE", "LABOR", "OUTSOURCE_LABOR", "BOTH_LABOR"])
        self.amount_cap_combo.setCurrentText("NONE")  # 기본값
        layout.addRow("금액 상한 타입 *:", self.amount_cap_combo)
        
        # Amount Cap Value (NULL 허용)
        self.amount_cap_spin = QSpinBox()
        self.amount_cap_spin.setRange(0, 999999999)
        self.amount_cap_spin.setValue(0)
        self.amount_cap_spin.setSpecialValueText("없음")
        layout.addRow("금액 상한 값:", self.amount_cap_spin)
        
        # Liability Ratio (선택사항 - LABOR 최댓값 규칙의 경우 NULL 가능)
        self.liability_ratio_spin = QDoubleSpinBox()
        self.liability_ratio_spin.setRange(0.0, 100.0)
        self.liability_ratio_spin.setDecimals(2)
        self.liability_ratio_spin.setSuffix(" %")
        self.liability_ratio_spin.setValue(0.0)
        self.liability_ratio_spin.setSpecialValueText("없음 (LABOR 최댓값 규칙용)")
        layout.addRow("구상율:", self.liability_ratio_spin)
        
        # amount_cap_type과 amount_cap_value 변경 시 구상율 필수 여부 업데이트
        self.amount_cap_combo.currentTextChanged.connect(self._update_liability_ratio_required)
        self.amount_cap_spin.valueChanged.connect(self._update_liability_ratio_required)
        
        # Valid From (날짜 형식)
        self.valid_from_edit = QLineEdit()
        self.valid_from_edit.setPlaceholderText("YYYY-MM-DD (선택사항)")
        layout.addRow("유효 시작일:", self.valid_from_edit)
        
        # Valid To (날짜 형식)
        self.valid_to_edit = QLineEdit()
        self.valid_to_edit.setPlaceholderText("YYYY-MM-DD (선택사항)")
        layout.addRow("유효 종료일:", self.valid_to_edit)
        
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
            idx = self.repair_region_combo.findText(rule_data["repair_region"])
            if idx >= 0:
                self.repair_region_combo.setCurrentIndex(idx)
        
        if "project_code" in rule_data:
            self.project_code_edit.setText(str(rule_data["project_code"]))
        
        if "exclude_project_code" in rule_data:
            exclude_code = rule_data["exclude_project_code"]
            self.exclude_project_code_edit.setText(exclude_code if exclude_code is not None else "")
        
        if "vehicle_classification" in rule_data:
            self.vehicle_class_edit.setText(str(rule_data["vehicle_classification"]))
        
        if "part_no" in rule_data:
            self.part_no_edit.setText(str(rule_data["part_no"]))
        
        if "part_name" in rule_data:
            self.part_name_edit.setText(str(rule_data["part_name"]))
        
        if "engine_form" in rule_data:
            self.engine_form_edit.setText(str(rule_data["engine_form"]))
        
        if "warranty_mileage_override" in rule_data and rule_data["warranty_mileage_override"]:
            self.warranty_mileage_spin.setValue(int(rule_data["warranty_mileage_override"]))
        
        if "warranty_period_override" in rule_data and rule_data["warranty_period_override"]:
            self.warranty_period_spin.setValue(int(rule_data["warranty_period_override"]))
        
        if "liability_ratio" in rule_data:
            # liability_ratio가 None일 수 있음
            liability_ratio = rule_data.get("liability_ratio")
            if liability_ratio is not None:
                self.liability_ratio_spin.setValue(float(liability_ratio))
            else:
                self.liability_ratio_spin.setValue(0.0)  # None이면 0으로 표시 (SpecialValueText)
        
        if "amount_cap_type" in rule_data:
            idx = self.amount_cap_combo.findText(rule_data["amount_cap_type"])
            if idx >= 0:
                self.amount_cap_combo.setCurrentIndex(idx)
        
        if "amount_cap_value" in rule_data and rule_data["amount_cap_value"]:
            self.amount_cap_spin.setValue(int(rule_data["amount_cap_value"]))
        
        if "valid_from" in rule_data:
            self.valid_from_edit.setText(str(rule_data["valid_from"]) if rule_data["valid_from"] else "")
        
        if "valid_to" in rule_data:
            self.valid_to_edit.setText(str(rule_data["valid_to"]) if rule_data["valid_to"] else "")
    
    def _update_liability_ratio_required(self):
        """amount_cap_type과 amount_cap_value에 따라 구상율 필수 여부 업데이트"""
        # LABOR 최댓값 규칙인 경우 구상율은 선택사항
        # 이 함수는 필요시 확장 가능 (예: 툴팁 변경 등)
        pass
    
    def get_data(self) -> Dict[str, Any]:
        """입력된 데이터 반환"""
        # Priority: -1이면 None으로 전달 (트리거가 자동으로 채움)
        priority = self.priority_spin.value()
        if priority == -1:
            priority = None
        
        # Warranty overrides: 0이면 None
        warranty_mileage = self.warranty_mileage_spin.value()
        if warranty_mileage == 0:
            warranty_mileage = None
        
        warranty_period = self.warranty_period_spin.value()
        if warranty_period == 0:
            warranty_period = None
        
        # Amount cap value: 0이면 None
        amount_cap_value = self.amount_cap_spin.value()
        if amount_cap_value == 0:
            amount_cap_value = None
        
        return {
            "priority": priority,
            "status": self.status_combo.currentText(),
            "repair_region": self.repair_region_combo.currentText(),
            "project_code": self.project_code_edit.text().strip() or "ALL",
            "exclude_project_code": self.exclude_project_code_edit.text().strip() or None,
            "vehicle_classification": self.vehicle_class_edit.text().strip() or "ALL",
            "part_no": self.part_no_edit.text().strip() or "ALL",
            "part_name": self.part_name_edit.text().strip() or "ALL",
            "engine_form": self.engine_form_edit.text().strip() or "ALL",
            "warranty_mileage_override": warranty_mileage,
            "warranty_period_override": warranty_period,
            "liability_ratio": self.liability_ratio_spin.value() if self.liability_ratio_spin.value() > 0.0 else None,
            "amount_cap_type": self.amount_cap_combo.currentText(),
            "amount_cap_value": amount_cap_value,
            "valid_from": self.valid_from_edit.text().strip() or None,
            "valid_to": self.valid_to_edit.text().strip() or None,
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
        """Rule의 변경점만 포맷팅하여 반환 (NULL, "ALL", "NONE" 제외)"""
        changes = []
        
        # 수리 지역 (ALL이 아닐 때만)
        repair_region = rule.get("repair_region")
        if repair_region and str(repair_region).strip().upper() != "ALL":
            changes.append(f"수리지역: {repair_region}")
        
        # 프로젝트 코드 (ALL이 아닐 때만)
        project_code = rule.get("project_code")
        if project_code:
            project_code = str(project_code).strip()
            if project_code and project_code.upper() != "ALL":
                changes.append(f"프로젝트: {project_code}")
        
        # 제외 프로젝트 (NULL이 아닐 때만)
        exclude_project = rule.get("exclude_project_code")
        if exclude_project:
            exclude_project = str(exclude_project).strip()
            if exclude_project:
                changes.append(f"제외: {exclude_project}")
        
        # 차계 (ALL이 아닐 때만)
        vehicle_classification = rule.get("vehicle_classification")
        if vehicle_classification:
            vehicle_classification = str(vehicle_classification).strip()
            if vehicle_classification and vehicle_classification.upper() != "ALL":
                changes.append(f"차계: {vehicle_classification}")
        
        # 부품명 (ALL이 아닐 때만)
        part_name = rule.get("part_name")
        if part_name:
            part_name = str(part_name).strip()
            if part_name and part_name.upper() != "ALL":
                changes.append(f"부품: {part_name}")
        
        # 부품 번호 (ALL이 아닐 때만)
        part_no = rule.get("part_no")
        if part_no:
            part_no = str(part_no).strip()
            if part_no and part_no.upper() != "ALL":
                changes.append(f"부품번호: {part_no}")
        
        # 엔진 형식 (ALL이 아닐 때만)
        engine_form = rule.get("engine_form")
        if engine_form:
            engine_form = str(engine_form).strip()
            if engine_form and engine_form.upper() != "ALL":
                changes.append(f"엔진: {engine_form}")
        
        # 구상율 (항상 표시)
        liability_ratio = rule.get("liability_ratio")
        if liability_ratio is not None:
            changes.append(f"구상율: {liability_ratio}%")
        
        # 보증 주행거리 오버라이드 (NULL이 아닐 때만)
        warranty_mileage = rule.get("warranty_mileage_override")
        if warranty_mileage is not None:
            changes.append(f"주행거리: {warranty_mileage}km")
        
        # 보증 기간 오버라이드 (NULL이 아닐 때만)
        warranty_period = rule.get("warranty_period_override")
        if warranty_period is not None:
            years = warranty_period / 365.0
            changes.append(f"보증기간: {years:.1f}년")
        
        # 금액 상한 (NULL이 아니고 NONE이 아닐 때만)
        amount_cap_value = rule.get("amount_cap_value")
        if amount_cap_value is not None:
            cap_type = rule.get("amount_cap_type", "NONE")
            if cap_type and str(cap_type).strip().upper() != "NONE":
                changes.append(f"상한: {amount_cap_value} ({cap_type})")
        
        # 적용 시작일 (NULL이 아닐 때만)
        valid_from = rule.get("valid_from")
        if valid_from:
            valid_from = str(valid_from).strip()
            if valid_from:
                changes.append(f"시작일: {valid_from}")
        
        # 적용 종료일 (NULL이 아닐 때만)
        valid_to = rule.get("valid_to")
        if valid_to:
            valid_to = str(valid_to).strip()
            if valid_to:
                changes.append(f"종료일: {valid_to}")
        
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

