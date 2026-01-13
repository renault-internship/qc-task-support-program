"""
Rule 관련 다이얼로그
"""
from typing import Dict, Any, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QFormLayout, QHBoxLayout, QVBoxLayout,
    QPushButton, QSpinBox, QDoubleSpinBox, QLineEdit,
    QComboBox, QTableWidget, QTableWidgetItem, QGroupBox, QLabel, QWidget, QFrame, QCheckBox, QTextEdit
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
        has_liability = "liability_ratio" in rule_data and rule_data.get("liability_ratio") is not None and rule_data.get("liability_ratio") >= 0.0
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
                # DB 값(0~1 범위)을 퍼센티지(0~100%)로 변환
                self.liability_ratio_spin.setValue(float(liability_ratio) * 100)
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
            # 구상률 선택 (0도 허용)
            liability_ratio = self.liability_ratio_spin.value() / 100.0  # 퍼센티지를 0~1 범위로 변환
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
            changes.append(f"구상율: {liability_ratio * 100:.0f}%")
        
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


class PreprocessResultDialog(QDialog):
    """전처리 결과 표시 다이얼로그"""
    def __init__(self, result, parent=None):
        """
        Args:
            result: PreprocessResult 객체
        """
        super().__init__(parent)
        self.result = result
        self.setWindowTitle("청구서 전처리 결과 보고서")
        self.setMinimumSize(1000, 800)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(10)
        
        # 제목
        title_label = QLabel("청구서 전처리 결과 보고서")
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold; color: #2c3e50;")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 처리 일시
        time_label = QLabel(f"처리일시: {self.result.process_time}")
        time_label.setStyleSheet("font-size: 10pt; color: #6c757d;")
        time_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(time_label)
        
        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)
        
        # 스크롤 영역
        from PySide6.QtWidgets import QScrollArea
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; }")
        
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(15)
        
        # 1. 기본 정보
        scroll_layout.addWidget(self._create_basic_info_section())
        
        # 2. 처리 결과 요약
        scroll_layout.addWidget(self._create_summary_section())
        
        # 3. 차계 및 프로젝트 분포
        scroll_layout.addWidget(self._create_project_section())
        
        # 4. 구상률 적용
        scroll_layout.addWidget(self._create_liability_section())
        
        # 5. 보증 기준
        scroll_layout.addWidget(self._create_warranty_section())
        
        # 6. 룰 사용 현황
        scroll_layout.addWidget(self._create_rule_usage_section())
        
        # 7. 예외 사항
        if self.result.warnings:
            scroll_layout.addWidget(self._create_warnings_section())
        
        # 8. 비고 (제일 아래)
        scroll_layout.addWidget(self._create_remarks_section())
        
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll)
        
        # 닫기 버튼
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_close = QPushButton("확인")
        btn_close.setFixedWidth(120)
        btn_close.setFixedHeight(35)
        btn_close.setStyleSheet("""
            QPushButton {
                background-color: #0d6efd;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 11pt;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0b5ed7;
            }
        """)
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        main_layout.addLayout(btn_layout)
        
        self.setLayout(main_layout)
    
    def _create_section_box(self, title: str, icon: str = "") -> QWidget:
        """섹션 위젯 생성 (테두리 없음)"""
        container = QWidget()
        container.setStyleSheet("background-color: white;")
        
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 10, 0, 10)  # 상하 패딩 균등하게
        layout.setSpacing(8)
        
        # 제목
        title_label = QLabel(title)
        title_label.setStyleSheet("""
            font-size: 13pt;
            font-weight: bold;
            color: #2c3e50;
            padding: 0;
            margin: 0;
        """)
        layout.addWidget(title_label)
        
        # 내용 컨테이너
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 5, 0, 0)  # 제목과 내용 사이 간격
        layout.addWidget(content)
        
        # 하단 구분선
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("color: #dee2e6;")
        layout.addWidget(separator)
        
        container.content_layout = content_layout
        return container
    
    def _create_info_row(self, label: str, value: str) -> QHBoxLayout:
        """정보 행 생성"""
        row = QHBoxLayout()
        row.setSpacing(10)
        
        lbl = QLabel(f"• {label}:")
        lbl.setStyleSheet("font-size: 10pt; font-weight: bold; color: #495057;")
        lbl.setFixedWidth(150)
        
        val = QLabel(value)
        val.setStyleSheet("font-size: 10pt; color: #212529;")
        
        row.addWidget(lbl)
        row.addWidget(val)
        row.addStretch()
        
        return row
    
    def _create_stat_card(self, label: str, value: str, color: str = "#0d6efd") -> QWidget:
        """통계 카드 생성"""
        card = QWidget()
        card.setFixedHeight(100)  # 높이 고정
        card.setStyleSheet(f"""
            QWidget {{
                background-color: {color};
                border-radius: 8px;
            }}
        """)
        
        layout = QVBoxLayout(card)
        layout.setSpacing(2)
        layout.setContentsMargins(10, 15, 10, 15)
        
        val_label = QLabel(value)
        val_label.setStyleSheet("font-size: 28pt; font-weight: bold; color: white;")
        val_label.setAlignment(Qt.AlignCenter)
        
        lbl_label = QLabel(label)
        lbl_label.setStyleSheet("font-size: 11pt; color: white;")
        lbl_label.setAlignment(Qt.AlignCenter)
        lbl_label.setWordWrap(True)
        
        layout.addWidget(val_label)
        layout.addWidget(lbl_label)
        
        return card
    
    def _create_basic_info_section(self) -> QWidget:
        """기본 정보 섹션"""
        box = self._create_section_box("청구서 정보")
        
        region_text = "국내 청구서" if self.result.repair_region == "DOMESTIC" else "해외 청구서"
        box.content_layout.addLayout(self._create_info_row("유형", region_text))
        box.content_layout.addLayout(self._create_info_row("협력사", f"{self.result.company_name} ({self.result.company_code})"))
        box.content_layout.addLayout(self._create_info_row("룰 테이블", self.result.rule_table_name))
        
        return box
    
    def _create_summary_section(self) -> QWidget:
        """처리 결과 요약 섹션"""
        box = self._create_section_box("처리 결과 요약")
        
        # 통계 카드들
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(15)
        
        cards_layout.addWidget(self._create_stat_card("총 청구 건수", f"{self.result.total_rows:,}", "#6c757d"))
        cards_layout.addWidget(self._create_stat_card("정상 처리", f"{self.result.success_rows:,}", "#28a745"))
        cards_layout.addWidget(self._create_stat_card("예외 처리", f"{self.result.warning_rows:,}", "#ffc107"))
        cards_layout.addWidget(self._create_stat_card("오류 발생", f"{self.result.error_rows:,}", "#dc3545"))
        
        box.content_layout.addLayout(cards_layout)
        
        return box
    
    def _create_project_section(self) -> QWidget:
        """차계 및 프로젝트 분포 섹션"""
        box = self._create_section_box("차계 및 프로젝트 코드 분석")
        
        if self.result.project_stats:
            table = QTableWidget()
            table.setColumnCount(4)
            table.setHorizontalHeaderLabels(["프로젝트 코드", "건수", "비율", "기본 구상률"])
            table.setRowCount(len(self.result.project_stats))
            
            table.horizontalHeader().setStretchLastSection(True)
            table.verticalHeader().setVisible(False)
            table.setAlternatingRowColors(True)
            table.setStyleSheet("""
                QTableWidget {
                    border: 1px solid #dee2e6;
                    gridline-color: #dee2e6;
                }
                QHeaderView::section {
                    background-color: #e9ecef;
                    padding: 8px;
                    border: 1px solid #dee2e6;
                    font-weight: bold;
                }
            """)
            
            total = self.result.total_rows
            for row, (project_code, (count, ratio)) in enumerate(sorted(self.result.project_stats.items(), key=lambda x: x[1][0], reverse=True)):
                percentage = (count / total * 100) if total > 0 else 0
                ratio_str = f"{ratio * 100:.0f}%" if ratio is not None else "미설정"
                
                table.setItem(row, 0, QTableWidgetItem(project_code))
                table.setItem(row, 1, QTableWidgetItem(f"{count:,}건"))
                table.setItem(row, 2, QTableWidgetItem(f"{percentage:.1f}%"))
                table.setItem(row, 3, QTableWidgetItem(ratio_str))
                
                # 중앙 정렬
                for col in range(4):
                    table.item(row, col).setTextAlignment(Qt.AlignCenter)
            
            table.resizeColumnsToContents()
            table.setMaximumHeight(min(300, 50 + len(self.result.project_stats) * 35))
            
            box.content_layout.addWidget(table)
        else:
            no_data = QLabel("프로젝트 통계 없음")
            no_data.setStyleSheet("font-size: 10pt; color: #6c757d; padding: 20px;")
            no_data.setAlignment(Qt.AlignCenter)
            box.content_layout.addWidget(no_data)
        
        return box
    
    def _create_liability_section(self) -> QWidget:
        """구상률 적용 섹션"""
        box = self._create_section_box("구상률 적용 내역")
        
        # 기본 구상률
        basic_widget = QWidget()
        basic_layout = QVBoxLayout(basic_widget)
        basic_layout.setContentsMargins(10, 10, 10, 10)
        basic_widget.setStyleSheet("background-color: #f8f9fa; border-radius: 5px;")
        
        basic_title = QLabel("기본 구상률 적용 (Common Project Liability)")
        basic_title.setStyleSheet("font-size: 10pt; font-weight: bold; color: #495057;")
        basic_layout.addWidget(basic_title)
        
        basic_value = QLabel(f"적용 건수: {self.result.common_liability_applied:,}건")
        basic_value.setStyleSheet("font-size: 10pt; color: #212529; padding-left: 10px;")
        basic_layout.addWidget(basic_value)
        
        box.content_layout.addWidget(basic_widget)
        
        # 세부 룰
        rule_widget = QWidget()
        rule_layout = QVBoxLayout(rule_widget)
        rule_layout.setContentsMargins(10, 10, 10, 10)
        rule_widget.setStyleSheet("background-color: #f8f9fa; border-radius: 5px;")
        
        rule_title = QLabel("세부 룰에 의한 구상률 변경")
        rule_title.setStyleSheet("font-size: 10pt; font-weight: bold; color: #495057;")
        rule_layout.addWidget(rule_title)
        
        if self.result.liability_ratio_rules_applied > 0:
            rule_value = QLabel(f"룰 적용 건수: {self.result.liability_ratio_rules_applied:,}건")
            rule_value.setStyleSheet("font-size: 10pt; color: #212529; padding-left: 10px;")
            rule_layout.addWidget(rule_value)
            
            rule_desc = QLabel(f"총 {self.result.liability_ratio_rules_applied:,}건의 구상률이 세부 룰로 오버라이드됨")
            rule_desc.setStyleSheet("font-size: 9pt; color: #6c757d; padding-left: 10px;")
            rule_layout.addWidget(rule_desc)
        else:
            no_rule = QLabel("적용된 구상률 변경 없음")
            no_rule.setStyleSheet("font-size: 10pt; color: #6c757d; padding-left: 10px;")
            rule_layout.addWidget(no_rule)
        
        box.content_layout.addWidget(rule_widget)
        
        return box
    
    def _create_warranty_section(self) -> QWidget:
        """보증 기준 및 초과 섹션 (통합)"""
        box = self._create_section_box("보증 주행거리 및 보증 기간")
        
        # 기본값
        default_widget = QWidget()
        default_layout = QVBoxLayout(default_widget)
        default_layout.setContentsMargins(10, 10, 10, 10)
        default_widget.setStyleSheet("background-color: #e7f3ff; border-radius: 5px;")
        
        default_title = QLabel("기본 보증 기준")
        default_title.setStyleSheet("font-size: 10pt; font-weight: bold; color: #004085;")
        default_layout.addWidget(default_title)
        
        default_mileage = QLabel(f"• 보증 주행거리: {self.result.default_mileage_threshold:,}km")
        default_mileage.setStyleSheet("font-size: 10pt; color: #004085; padding-left: 10px;")
        default_layout.addWidget(default_mileage)
        
        default_period = QLabel(f"• 보증 기간: {self.result.default_warranty_years}년")
        default_period.setStyleSheet("font-size: 10pt; color: #004085; padding-left: 10px;")
        default_layout.addWidget(default_period)
        
        box.content_layout.addWidget(default_widget)
        
        # 오버라이드
        if self.result.mileage_overrides or self.result.period_overrides:
            override_widget = QWidget()
            override_layout = QVBoxLayout(override_widget)
            override_layout.setContentsMargins(10, 10, 10, 10)
            override_widget.setStyleSheet("background-color: #fff3cd; border-radius: 5px;")
            
            override_title = QLabel("세부 룰에 의한 변경")
            override_title.setStyleSheet("font-size: 10pt; font-weight: bold; color: #856404;")
            override_layout.addWidget(override_title)
            
            if self.result.mileage_overrides:
                for mileage, count in sorted(self.result.mileage_overrides.items()):
                    lbl = QLabel(f"• 보증 주행거리 {mileage:,}km 적용: {count:,}건")
                    lbl.setStyleSheet("font-size: 10pt; color: #856404; padding-left: 10px;")
                    override_layout.addWidget(lbl)
            
            if self.result.period_overrides:
                for years, count in sorted(self.result.period_overrides.items()):
                    lbl = QLabel(f"• 보증 기간 {years}년 적용: {count:,}건")
                    lbl.setStyleSheet("font-size: 10pt; color: #856404; padding-left: 10px;")
                    override_layout.addWidget(lbl)
            
            box.content_layout.addWidget(override_widget)
        
        # 보증 범위 초과 통계
        exceeded_widget = QWidget()
        exceeded_layout = QVBoxLayout(exceeded_widget)
        exceeded_layout.setContentsMargins(10, 10, 10, 10)
        exceeded_widget.setStyleSheet("background-color: #f8d7da; border-radius: 5px;")
        
        exceeded_title = QLabel("보증 범위 초과 (구상률 0% 적용)")
        exceeded_title.setStyleSheet("font-size: 10pt; font-weight: bold; color: #721c24;")
        exceeded_layout.addWidget(exceeded_title)
        
        exceeded_layout.addLayout(self._create_info_row("주행거리 초과", f"{self.result.mileage_exceeded_rows:,}건"))
        exceeded_layout.addLayout(self._create_info_row("보증기간 초과", f"{self.result.period_exceeded_rows:,}건"))
        exceeded_layout.addLayout(self._create_info_row("중복 초과 (거리+기간)", f"{self.result.both_exceeded_rows:,}건"))
        
        # 구분선
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #dc3545;")
        exceeded_layout.addWidget(sep)
        
        total_lbl = QLabel(f"총 개수: {self.result.warranty_highlighted_rows:,}개 (노란색 하이라이트)")
        total_lbl.setStyleSheet("font-size: 11pt; font-weight: bold; color: #721c24;")
        exceeded_layout.addWidget(total_lbl)
        
        box.content_layout.addWidget(exceeded_widget)
        
        return box
    
    def _create_rule_usage_section(self) -> QWidget:
        """룰 사용 현황 섹션"""
        box = self._create_section_box("룰 사용 현황")
        
        # 요약
        total_rules = len(self.result.rule_usage) + len(self.result.unused_rules)
        summary_layout = QHBoxLayout()
        summary_layout.setSpacing(10)
        
        total_lbl = QLabel(f"총 활성 룰: {total_rules}개")
        total_lbl.setStyleSheet("font-size: 10pt; font-weight: bold;")
        
        used_lbl = QLabel(f"적용됨: {len(self.result.rule_usage)}개")
        used_lbl.setStyleSheet("font-size: 10pt; color: #28a745;")
        
        unused_lbl = QLabel(f"미적용: {len(self.result.unused_rules)}개")
        unused_lbl.setStyleSheet("font-size: 10pt; color: #dc3545;")
        
        summary_layout.addWidget(total_lbl)
        summary_layout.addWidget(used_lbl)
        summary_layout.addWidget(unused_lbl)
        summary_layout.addStretch()
        
        box.content_layout.addLayout(summary_layout)
        
        # 적용된 룰 테이블
        if self.result.rule_usage:
            used_title = QLabel("적용된 룰")
            used_title.setStyleSheet("font-size: 10pt; font-weight: bold; color: #28a745; margin-top: 10px;")
            box.content_layout.addWidget(used_title)
            
            used_table = QTableWidget()
            used_table.setColumnCount(3)
            used_table.setHorizontalHeaderLabels(["룰 ID", "설명", "적용 횟수"])
            
            display_count = min(10, len(self.result.rule_usage))
            used_table.setRowCount(display_count)
            
            for row, (rule_id, (desc, count)) in enumerate(sorted(self.result.rule_usage.items(), key=lambda x: x[1][1], reverse=True)[:10]):
                used_table.setItem(row, 0, QTableWidgetItem(f"R-{rule_id}"))
                used_table.setItem(row, 1, QTableWidgetItem(desc))
                used_table.setItem(row, 2, QTableWidgetItem(f"{count:,}회"))
                
                used_table.item(row, 0).setTextAlignment(Qt.AlignCenter)
                used_table.item(row, 2).setTextAlignment(Qt.AlignCenter)
            
            used_table.horizontalHeader().setStretchLastSection(True)
            used_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #e9ecef; font-weight: bold; }")
            used_table.setAlternatingRowColors(True)
            used_table.resizeColumnsToContents()
            used_table.setMaximumHeight(min(300, 50 + display_count * 35))
            
            box.content_layout.addWidget(used_table)
            
            if len(self.result.rule_usage) > 10:
                more_lbl = QLabel(f"... 외 {len(self.result.rule_usage) - 10}개 더")
                more_lbl.setStyleSheet("font-size: 9pt; color: #6c757d;")
                box.content_layout.addWidget(more_lbl)
        
        # 미적용 룰 테이블
        if self.result.unused_rules:
            unused_title = QLabel("미적용 룰 (확인 필요)")
            unused_title.setStyleSheet("font-size: 10pt; font-weight: bold; color: #dc3545; margin-top: 10px;")
            box.content_layout.addWidget(unused_title)
            
            unused_table = QTableWidget()
            unused_table.setColumnCount(3)
            unused_table.setHorizontalHeaderLabels(["룰 ID", "설명", "미적용 이유"])
            
            display_count = min(10, len(self.result.unused_rules))
            unused_table.setRowCount(display_count)
            
            for row, (rule_id, desc, reason) in enumerate(self.result.unused_rules[:10]):
                unused_table.setItem(row, 0, QTableWidgetItem(f"R-{rule_id}"))
                unused_table.setItem(row, 1, QTableWidgetItem(desc))
                unused_table.setItem(row, 2, QTableWidgetItem(reason))
                
                unused_table.item(row, 0).setTextAlignment(Qt.AlignCenter)
            
            unused_table.horizontalHeader().setStretchLastSection(True)
            unused_table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #e9ecef; font-weight: bold; }")
            unused_table.setAlternatingRowColors(True)
            unused_table.resizeColumnsToContents()
            unused_table.setMaximumHeight(min(300, 50 + display_count * 35))
            
            box.content_layout.addWidget(unused_table)
            
            if len(self.result.unused_rules) > 10:
                more_lbl = QLabel(f"... 외 {len(self.result.unused_rules) - 10}개 더")
                more_lbl.setStyleSheet("font-size: 9pt; color: #6c757d;")
                box.content_layout.addWidget(more_lbl)
            
            # 권장 조치
            advice = QLabel("미적용 룰의 조건을 검토하거나 우선순위 조정이 필요합니다.")
            advice.setStyleSheet("font-size: 9pt; color: #856404; background-color: #fff3cd; padding: 8px; border-radius: 4px; margin-top: 5px;")
            advice.setWordWrap(True)
            box.content_layout.addWidget(advice)
        
        return box
    
    def _create_remarks_section(self) -> QWidget:
        """비고 섹션 (특이사항만 기록)"""
        box = self._create_section_box("비고")
        
        remarks_widget = QWidget()
        remarks_layout = QVBoxLayout(remarks_widget)
        remarks_layout.setContentsMargins(10, 10, 10, 10)
        remarks_widget.setStyleSheet("background-color: #f8f9fa; border-radius: 5px; border: 1px solid #dee2e6;")
        
        # 비고가 있으면 표시, 없으면 "기타 특이사항 없음"
        if self.result.remarks:
            remarks_text = QTextEdit()
            remarks_text.setReadOnly(True)
            remarks_text.setPlainText("\n".join(self.result.remarks))
            remarks_text.setStyleSheet("""
                QTextEdit {
                    font-size: 9pt;
                    color: #495057;
                    background-color: white;
                    border: 1px solid #ced4da;
                    border-radius: 4px;
                    padding: 8px;
                }
            """)
            remarks_text.setMaximumHeight(200)
            remarks_layout.addWidget(remarks_text)
        else:
            no_remarks = QLabel("기타 특이사항 없음")
            no_remarks.setStyleSheet("font-size: 10pt; color: #6c757d; font-style: italic;")
            no_remarks.setAlignment(Qt.AlignCenter)
            remarks_layout.addWidget(no_remarks)
        
        box.content_layout.addWidget(remarks_widget)
        
        return box
    
    def _create_warnings_section(self) -> QWidget:
        """예외 사항 섹션"""
        box = self._create_section_box("예외 사항 (수동 확인 필요)")
        
        warning_count = QLabel(f"총 {len(self.result.warnings)}건의 예외 사항이 발견되었습니다.")
        warning_count.setStyleSheet("font-size: 10pt; font-weight: bold; color: #dc3545; margin-bottom: 10px;")
        box.content_layout.addWidget(warning_count)
        
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["행번호", "차계", "부품번호", "사유"])
        
        display_count = min(30, len(self.result.warnings))
        table.setRowCount(display_count)
        
        for row, (row_num, vehicle, part_no, reason) in enumerate(self.result.warnings[:30]):
            table.setItem(row, 0, QTableWidgetItem(str(row_num)))
            table.setItem(row, 1, QTableWidgetItem(vehicle))
            table.setItem(row, 2, QTableWidgetItem(part_no))
            table.setItem(row, 3, QTableWidgetItem(reason))
            
            table.item(row, 0).setTextAlignment(Qt.AlignCenter)
        
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setStyleSheet("QHeaderView::section { background-color: #e9ecef; font-weight: bold; }")
        table.setAlternatingRowColors(True)
        table.resizeColumnsToContents()
        table.setMaximumHeight(min(400, 50 + display_count * 35))
        
        box.content_layout.addWidget(table)
        
        if len(self.result.warnings) > 30:
            more_lbl = QLabel(f"... 외 {len(self.result.warnings) - 30}건 더")
            more_lbl.setStyleSheet("font-size: 9pt; color: #6c757d;")
            box.content_layout.addWidget(more_lbl)
        
        note = QLabel("* 예외 항목은 기본값으로 처리되었으며, 수동 확인이 필요합니다.")
        note.setStyleSheet("font-size: 9pt; color: #6c757d; font-style: italic; margin-top: 5px;")
        box.content_layout.addWidget(note)
        
        return box
    
    def _generate_report(self) -> str:
        """레거시 메서드 - 사용하지 않음"""
        return ""

