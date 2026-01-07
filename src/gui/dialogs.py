"""
Rule 관련 다이얼로그
"""
from typing import Dict, Any, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QFormLayout, QHBoxLayout, QVBoxLayout,
    QPushButton, QSpinBox, QDoubleSpinBox, QLineEdit,
    QComboBox, QTableWidget, QTableWidgetItem, QGroupBox, QLabel, QWidget, QFrame, QCheckBox
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
        self.setFixedSize(800, 600)  # 가로로 넓게, 세로 높이 증가
        
        # 메인 레이아웃 (가로 배치)
        main_layout = QHBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # ========== 섹션1: 규칙 적용 대상 ==========
        section1_group = QGroupBox("규칙 적용 대상")
        section1_group.setStyleSheet("QGroupBox::title { color: black; font-size: 16pt; font-weight: bold; }")
        section1_layout = QVBoxLayout()
        section1_layout.setSpacing(15)
        
        # Repair Region (CHECK IN ('DOMESTIC','OVERSEAS','ALL'))
        repair_region_widget = QWidget()
        repair_region_layout = QVBoxLayout(repair_region_widget)
        repair_region_layout.setContentsMargins(0, 0, 0, 0)
        repair_region_label = QLabel("수리지역 *:")
        repair_region_label.setStyleSheet("font-size: 8pt;")
        repair_region_layout.addWidget(repair_region_label)
        self.repair_region_combo = QComboBox()
        self.repair_region_combo.addItems(["DOMESTIC", "OVERSEAS", "ALL"])
        self.repair_region_combo.setCurrentText("ALL")  # 기본값
        repair_region_layout.addWidget(self.repair_region_combo)
        section1_layout.addWidget(repair_region_widget)
        
        # Project Code (DEFAULT 'ALL')
        project_code_widget = QWidget()
        project_code_layout = QVBoxLayout(project_code_widget)
        project_code_layout.setContentsMargins(0, 0, 0, 0)
        project_code_label = QLabel("프로젝트코드 *:")
        project_code_label.setStyleSheet("font-size: 8pt;")
        project_code_layout.addWidget(project_code_label)
        self.project_code_edit = QLineEdit()
        self.project_code_edit.setPlaceholderText("기본값: ALL")
        self.project_code_edit.setText("ALL")  # 기본값
        project_code_layout.addWidget(self.project_code_edit)
        section1_layout.addWidget(project_code_widget)
        
        # Exclude Project Code (NULL 허용)
        exclude_project_code_widget = QWidget()
        exclude_project_code_layout = QVBoxLayout(exclude_project_code_widget)
        exclude_project_code_layout.setContentsMargins(0, 0, 0, 0)
        exclude_project_code_label = QLabel("제외프로젝트코드:")
        exclude_project_code_label.setStyleSheet("font-size: 8pt;")
        exclude_project_code_layout.addWidget(exclude_project_code_label)
        self.exclude_project_code_edit = QLineEdit()
        self.exclude_project_code_edit.setPlaceholderText("제외할 프로젝트 코드 (선택사항)")
        exclude_project_code_layout.addWidget(self.exclude_project_code_edit)
        section1_layout.addWidget(exclude_project_code_widget)
        
        # Vehicle Classification (DEFAULT 'ALL')
        vehicle_class_widget = QWidget()
        vehicle_class_layout = QVBoxLayout(vehicle_class_widget)
        vehicle_class_layout.setContentsMargins(0, 0, 0, 0)
        vehicle_class_label = QLabel("차계 *:")
        vehicle_class_label.setStyleSheet("font-size: 8pt;")
        vehicle_class_layout.addWidget(vehicle_class_label)
        self.vehicle_class_edit = QLineEdit()
        self.vehicle_class_edit.setPlaceholderText("기본값: ALL")
        self.vehicle_class_edit.setText("ALL")  # 기본값
        vehicle_class_layout.addWidget(self.vehicle_class_edit)
        section1_layout.addWidget(vehicle_class_widget)
        
        # Part No (NOT NULL DEFAULT 'ALL')
        part_no_widget = QWidget()
        part_no_layout = QVBoxLayout(part_no_widget)
        part_no_layout.setContentsMargins(0, 0, 0, 0)
        part_no_label = QLabel("부품번호 *:")
        part_no_label.setStyleSheet("font-size: 8pt;")
        part_no_layout.addWidget(part_no_label)
        self.part_no_edit = QLineEdit()
        self.part_no_edit.setPlaceholderText("기본값: ALL")
        self.part_no_edit.setText("ALL")  # 기본값
        part_no_layout.addWidget(self.part_no_edit)
        section1_layout.addWidget(part_no_widget)
        
        # Part Name (NOT NULL DEFAULT 'ALL')
        part_name_widget = QWidget()
        part_name_layout = QVBoxLayout(part_name_widget)
        part_name_layout.setContentsMargins(0, 0, 0, 0)
        part_name_label = QLabel("부품명 *:")
        part_name_label.setStyleSheet("font-size: 8pt;")
        part_name_layout.addWidget(part_name_label)
        self.part_name_edit = QLineEdit()
        self.part_name_edit.setPlaceholderText("기본값: ALL")
        self.part_name_edit.setText("ALL")  # 기본값
        part_name_layout.addWidget(self.part_name_edit)
        section1_layout.addWidget(part_name_widget)
        
        # Engine Form (NOT NULL DEFAULT 'ALL')
        engine_form_widget = QWidget()
        engine_form_layout = QVBoxLayout(engine_form_widget)
        engine_form_layout.setContentsMargins(0, 0, 0, 0)
        engine_form_label = QLabel("엔진형태 *:")
        engine_form_label.setStyleSheet("font-size: 8pt;")
        engine_form_layout.addWidget(engine_form_label)
        self.engine_form_edit = QLineEdit()
        self.engine_form_edit.setPlaceholderText("기본값: ALL")
        self.engine_form_edit.setText("ALL")  # 기본값
        engine_form_layout.addWidget(self.engine_form_edit)
        section1_layout.addWidget(engine_form_widget)
        
        section1_group.setLayout(section1_layout)
        main_layout.addWidget(section1_group, 1)  # stretch factor 1
        
        # ========== 섹션2: 규칙 타입 ==========
        section2_group = QGroupBox("규칙 타입(택 1)")
        section2_group.setStyleSheet("QGroupBox::title { color: black; font-size: 16pt; font-weight: bold; }")
        section2_layout = QVBoxLayout()
        section2_layout.setSpacing(15)
        
        # 옵션 1: 구상률
        liability_group_widget = QWidget()
        liability_group_layout = QVBoxLayout(liability_group_widget)
        liability_group_layout.setContentsMargins(0, 0, 0, 0)
        liability_group_layout.setSpacing(3)
        
        self.liability_checkbox = QCheckBox("구상률")
        self.liability_checkbox.setStyleSheet("font-size: 10pt; font-weight: bold;")
        liability_group_layout.addWidget(self.liability_checkbox)
        
        liability_ratio_widget = QWidget()
        liability_ratio_layout = QVBoxLayout(liability_ratio_widget)
        liability_ratio_layout.setContentsMargins(20, 0, 0, 0)  # 체크박스와 정렬
        self.liability_ratio_label = QLabel("구상률:")
        self.liability_ratio_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
        liability_ratio_layout.addWidget(self.liability_ratio_label)
        self.liability_ratio_spin = QDoubleSpinBox()
        self.liability_ratio_spin.setRange(0.0, 100.0)
        self.liability_ratio_spin.setDecimals(2)
        self.liability_ratio_spin.setSuffix(" %")
        self.liability_ratio_spin.setValue(0.0)
        self.liability_ratio_spin.setSpecialValueText("없음")
        self.liability_ratio_spin.setEnabled(False)  # 초기 비활성화
        liability_ratio_layout.addWidget(self.liability_ratio_spin)
        liability_group_layout.addWidget(liability_ratio_widget)
        section2_layout.addWidget(liability_group_widget)
        
        # 구분선 1
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.Shape.HLine)
        separator1.setFrameShadow(QFrame.Shadow.Sunken)
        separator1.setStyleSheet("color: #ccc;")
        section2_layout.addWidget(separator1)
        
        # 옵션 2: 공임비상한
        amount_cap_group_widget = QWidget()
        amount_cap_group_layout = QVBoxLayout(amount_cap_group_widget)
        amount_cap_group_layout.setContentsMargins(0, 0, 0, 0)
        amount_cap_group_layout.setSpacing(3)
        
        self.amount_cap_checkbox = QCheckBox("공임비")
        self.amount_cap_checkbox.setStyleSheet("font-size: 10pt; font-weight: bold;")
        amount_cap_group_layout.addWidget(self.amount_cap_checkbox)
        
        # 공임타입
        amount_cap_type_widget = QWidget()
        amount_cap_type_layout = QVBoxLayout(amount_cap_type_widget)
        amount_cap_type_layout.setContentsMargins(20, 0, 0, 0)  # 체크박스와 정렬
        self.amount_cap_type_label = QLabel("공임타입 *:")
        self.amount_cap_type_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
        amount_cap_type_layout.addWidget(self.amount_cap_type_label)
        self.amount_cap_combo = QComboBox()
        self.amount_cap_combo.addItems(["NONE", "LABOR", "OUTSOURCE_LABOR", "BOTH_LABOR"])
        self.amount_cap_combo.setCurrentText("NONE")  # 기본값
        self.amount_cap_combo.setEnabled(False)  # 초기 비활성화
        amount_cap_type_layout.addWidget(self.amount_cap_combo)
        amount_cap_group_layout.addWidget(amount_cap_type_widget)
        
        # 금액상한값
        amount_cap_value_widget = QWidget()
        amount_cap_value_layout = QVBoxLayout(amount_cap_value_widget)
        amount_cap_value_layout.setContentsMargins(20, 0, 0, 0)  # 체크박스와 정렬
        self.amount_cap_value_label = QLabel("금액상한값(원):")
        self.amount_cap_value_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
        amount_cap_value_layout.addWidget(self.amount_cap_value_label)
        self.amount_cap_spin = QSpinBox()
        self.amount_cap_spin.setRange(0, 999999999)
        self.amount_cap_spin.setValue(0)
        self.amount_cap_spin.setSpecialValueText("없음")
        self.amount_cap_spin.setEnabled(False)  # 초기 비활성화
        amount_cap_value_layout.addWidget(self.amount_cap_spin)
        amount_cap_group_layout.addWidget(amount_cap_value_widget)
        section2_layout.addWidget(amount_cap_group_widget)
        
        # 구분선 2
        separator2 = QFrame()
        separator2.setFrameShape(QFrame.Shape.HLine)
        separator2.setFrameShadow(QFrame.Shadow.Sunken)
        separator2.setStyleSheet("color: #ccc;")
        section2_layout.addWidget(separator2)
        
        # 옵션 3: 보증주행거리 및 보증기간
        warranty_group_widget = QWidget()
        warranty_group_layout = QVBoxLayout(warranty_group_widget)
        warranty_group_layout.setContentsMargins(0, 0, 0, 0)
        warranty_group_layout.setSpacing(3)
        
        self.warranty_checkbox = QCheckBox("보증주행거리 및 보증기간")
        self.warranty_checkbox.setStyleSheet("font-size: 10pt; font-weight: bold;")
        warranty_group_layout.addWidget(self.warranty_checkbox)
        
        # 보증주행거리
        warranty_mileage_widget = QWidget()
        warranty_mileage_layout = QVBoxLayout(warranty_mileage_widget)
        warranty_mileage_layout.setContentsMargins(20, 0, 0, 0)  # 체크박스와 정렬
        self.warranty_mileage_label = QLabel("보증주행거리(km):")
        self.warranty_mileage_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
        warranty_mileage_layout.addWidget(self.warranty_mileage_label)
        self.warranty_mileage_spin = QSpinBox()
        self.warranty_mileage_spin.setRange(0, 1000000)
        self.warranty_mileage_spin.setValue(0)
        self.warranty_mileage_spin.setSpecialValueText("없음")
        self.warranty_mileage_spin.setEnabled(False)  # 초기 비활성화
        warranty_mileage_layout.addWidget(self.warranty_mileage_spin)
        warranty_group_layout.addWidget(warranty_mileage_widget)
        
        # 보증기간
        warranty_period_widget = QWidget()
        warranty_period_layout = QVBoxLayout(warranty_period_widget)
        warranty_period_layout.setContentsMargins(20, 0, 0, 0)  # 체크박스와 정렬
        self.warranty_period_label = QLabel("보증기간(년):")
        self.warranty_period_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
        warranty_period_layout.addWidget(self.warranty_period_label)
        self.warranty_period_spin = QSpinBox()
        self.warranty_period_spin.setRange(0, 3650)
        self.warranty_period_spin.setValue(0)
        self.warranty_period_spin.setSpecialValueText("없음")
        self.warranty_period_spin.setEnabled(False)  # 초기 비활성화
        warranty_period_layout.addWidget(self.warranty_period_spin)
        warranty_group_layout.addWidget(warranty_period_widget)
        section2_layout.addWidget(warranty_group_widget)
        
        # 하단 여백 추가 (세 번째 섹션과 동일하게)
        section2_layout.addStretch()
        
        # 체크박스 이벤트 연결
        self.liability_checkbox.stateChanged.connect(self._on_rule_type_changed)
        self.amount_cap_checkbox.stateChanged.connect(self._on_rule_type_changed)
        self.warranty_checkbox.stateChanged.connect(self._on_rule_type_changed)
        
        # amount_cap_type과 amount_cap_value 변경 시 구상율 필수 여부 업데이트
        self.amount_cap_combo.currentTextChanged.connect(self._update_liability_ratio_required)
        self.amount_cap_spin.valueChanged.connect(self._update_liability_ratio_required)
        
        section2_group.setLayout(section2_layout)
        main_layout.addWidget(section2_group, 1)  # stretch factor 1
        
        # ========== 섹션3: 기타 ==========
        section3_group = QGroupBox("기타(선택사항)")
        section3_group.setStyleSheet("QGroupBox::title { color: black; font-size: 16pt; font-weight: bold; }")
        section3_layout = QVBoxLayout()
        section3_layout.setSpacing(15)
        
        # Valid From (날짜 형식)
        valid_from_widget = QWidget()
        valid_from_layout = QVBoxLayout(valid_from_widget)
        valid_from_layout.setContentsMargins(0, 0, 0, 0)
        valid_from_label = QLabel("유효시작일:")
        valid_from_label.setStyleSheet("font-size: 8pt;")
        valid_from_layout.addWidget(valid_from_label)
        self.valid_from_edit = QLineEdit()
        self.valid_from_edit.setPlaceholderText("YYYY-MM-DD (선택사항)")
        valid_from_layout.addWidget(self.valid_from_edit)
        section3_layout.addWidget(valid_from_widget)
        
        # Valid To (날짜 형식)
        valid_to_widget = QWidget()
        valid_to_layout = QVBoxLayout(valid_to_widget)
        valid_to_layout.setContentsMargins(0, 0, 0, 0)
        valid_to_label = QLabel("유효종료일:")
        valid_to_label.setStyleSheet("font-size: 8pt;")
        valid_to_layout.addWidget(valid_to_label)
        self.valid_to_edit = QLineEdit()
        self.valid_to_edit.setPlaceholderText("YYYY-MM-DD (선택사항)")
        valid_to_layout.addWidget(self.valid_to_edit)
        section3_layout.addWidget(valid_to_widget)
        
        # 하단 여백 추가
        section3_layout.addStretch()
        
        # 버튼 (우측 하단)
        button_layout = QHBoxLayout()
        button_layout.addStretch()  # 좌측 여백
        self.save_btn = QPushButton("저장")
        self.save_btn.clicked.connect(self.accept)
        self.cancel_btn = QPushButton("취소")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.cancel_btn)
        section3_layout.addLayout(button_layout)
        
        section3_group.setLayout(section3_layout)
        main_layout.addWidget(section3_group, 1)  # stretch factor 1
        
        self.setLayout(main_layout)
        
        # 수정 모드인 경우 기존 데이터로 채우기
        if self.is_edit_mode and rule_data:
            self._load_rule_data(rule_data)
    
    def _load_rule_data(self, rule_data: Dict[str, Any]):
        """기존 규칙 데이터로 폼 채우기"""
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
        
        # 규칙 타입 체크박스 설정 (데이터 기반으로 어떤 옵션이 선택되었는지 판단)
        has_liability = "liability_ratio" in rule_data and rule_data.get("liability_ratio") is not None and rule_data.get("liability_ratio") > 0.0
        has_amount_cap = "amount_cap_type" in rule_data and rule_data.get("amount_cap_type") and rule_data.get("amount_cap_type") != "NONE"
        has_warranty = ("warranty_mileage_override" in rule_data and rule_data.get("warranty_mileage_override")) or \
                      ("warranty_period_override" in rule_data and rule_data.get("warranty_period_override"))
        
        # 하나만 선택 가능하므로 우선순위: 구상률 > 공임비상한 > 보증
        # 체크박스 설정 시 stateChanged 시그널이 발생하여 _on_rule_type_changed가 자동 호출됨
        if has_liability:
            self.liability_checkbox.setChecked(True)
        elif has_amount_cap:
            self.amount_cap_checkbox.setChecked(True)
        elif has_warranty:
            self.warranty_checkbox.setChecked(True)
        
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
    
    def _on_rule_type_changed(self):
        """규칙 타입 체크박스 변경 시 다른 옵션 비활성화 및 활성화"""
        sender = self.sender()
        
        # 체크박스가 체크된 경우 다른 체크박스 해제
        if sender.isChecked():
            if sender == self.liability_checkbox:
                self.amount_cap_checkbox.setChecked(False)
                self.warranty_checkbox.setChecked(False)
                # 구상률 필드 활성화
                self.liability_ratio_spin.setEnabled(True)
                self.liability_ratio_label.setStyleSheet("font-size: 8pt; color: black;")
                # 공임비상한 필드 비활성화
                self.amount_cap_combo.setEnabled(False)
                self.amount_cap_spin.setEnabled(False)
                self.amount_cap_type_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                self.amount_cap_value_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                # 보증 필드 비활성화
                self.warranty_mileage_spin.setEnabled(False)
                self.warranty_period_spin.setEnabled(False)
                self.warranty_mileage_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                self.warranty_period_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
            elif sender == self.amount_cap_checkbox:
                self.liability_checkbox.setChecked(False)
                self.warranty_checkbox.setChecked(False)
                # 구상률 필드 비활성화
                self.liability_ratio_spin.setEnabled(False)
                self.liability_ratio_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                # 공임비상한 필드 활성화
                self.amount_cap_combo.setEnabled(True)
                self.amount_cap_spin.setEnabled(True)
                self.amount_cap_type_label.setStyleSheet("font-size: 8pt; color: black;")
                self.amount_cap_value_label.setStyleSheet("font-size: 8pt; color: black;")
                # 보증 필드 비활성화
                self.warranty_mileage_spin.setEnabled(False)
                self.warranty_period_spin.setEnabled(False)
                self.warranty_mileage_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                self.warranty_period_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
            elif sender == self.warranty_checkbox:
                self.liability_checkbox.setChecked(False)
                self.amount_cap_checkbox.setChecked(False)
                # 구상률 필드 비활성화
                self.liability_ratio_spin.setEnabled(False)
                self.liability_ratio_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                # 공임비상한 필드 비활성화
                self.amount_cap_combo.setEnabled(False)
                self.amount_cap_spin.setEnabled(False)
                self.amount_cap_type_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                self.amount_cap_value_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
                # 보증 필드 활성화
                self.warranty_mileage_spin.setEnabled(True)
                self.warranty_period_spin.setEnabled(True)
                self.warranty_mileage_label.setStyleSheet("font-size: 8pt; color: black;")
                self.warranty_period_label.setStyleSheet("font-size: 8pt; color: black;")
        else:
            # 체크박스가 해제된 경우 모든 필드 비활성화
            self.liability_ratio_spin.setEnabled(False)
            self.amount_cap_combo.setEnabled(False)
            self.amount_cap_spin.setEnabled(False)
            self.warranty_mileage_spin.setEnabled(False)
            self.warranty_period_spin.setEnabled(False)
            # 모든 레이블 회색으로 변경
            self.liability_ratio_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
            self.amount_cap_type_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
            self.amount_cap_value_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
            self.warranty_mileage_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
            self.warranty_period_label.setStyleSheet("font-size: 8pt; color: #cccccc;")
    
    def _update_liability_ratio_required(self):
        """amount_cap_type과 amount_cap_value에 따라 구상율 필수 여부 업데이트"""
        # LABOR 최댓값 규칙인 경우 구상율은 선택사항
        # 이 함수는 필요시 확장 가능 (예: 툴팁 변경 등)
        pass
    
    def get_data(self) -> Dict[str, Any]:
        """입력된 데이터 반환"""
        # Priority: None으로 설정 (데이터베이스에서 자동으로 최대값+1 처리)
        priority = None
        
        # Status: 기본값 "ACTIVE" (UI에서 제거되었지만 데이터베이스에 필요)
        status = "ACTIVE"
        
        # 체크박스 상태에 따라 데이터 설정
        liability_ratio = None
        amount_cap_type = "NONE"
        amount_cap_value = None
        warranty_mileage = None
        warranty_period = None
        
        if self.liability_checkbox.isChecked():
            # 구상률 선택
            liability_ratio = self.liability_ratio_spin.value() if self.liability_ratio_spin.value() > 0.0 else None
        elif self.amount_cap_checkbox.isChecked():
            # 공임비상한 선택
            amount_cap_type = self.amount_cap_combo.currentText()
            amount_cap_value = self.amount_cap_spin.value() if self.amount_cap_spin.value() > 0 else None
        elif self.warranty_checkbox.isChecked():
            # 보증주행거리 및 보증기간 선택
            warranty_mileage = self.warranty_mileage_spin.value() if self.warranty_mileage_spin.value() > 0 else None
            warranty_period = self.warranty_period_spin.value() if self.warranty_period_spin.value() > 0 else None
        
        return {
            "priority": priority,
            "status": status,
            "repair_region": self.repair_region_combo.currentText(),
            "project_code": self.project_code_edit.text().strip() or "ALL",
            "exclude_project_code": self.exclude_project_code_edit.text().strip() or None,
            "vehicle_classification": self.vehicle_class_edit.text().strip() or "ALL",
            "part_no": self.part_no_edit.text().strip() or "ALL",
            "part_name": self.part_name_edit.text().strip() or "ALL",
            "engine_form": self.engine_form_edit.text().strip() or "ALL",
            "warranty_mileage_override": warranty_mileage,
            "warranty_period_override": warranty_period,
            "liability_ratio": liability_ratio,
            "amount_cap_type": amount_cap_type,
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

