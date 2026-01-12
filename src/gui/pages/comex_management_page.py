"""
comex 관리 페이지 - 협력사 목록 및 룰 관리
"""
from typing import Dict, Any, Optional, List

from PySide6.QtCore import Qt, QSortFilterProxyModel, QRegularExpression
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QListWidget, QListWidgetItem, QMessageBox, QDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QMenu
)

from src.database import (
    get_all_companies, get_all_companies_with_code, get_company_info, 
    get_rules_from_table, add_rule_to_table, update_rule_in_table, 
    delete_rule_from_table, upsert_company, update_company_remark,
    update_rule_priorities, update_company, delete_company
)
from src.gui.dialogs import AddRuleDialog


class AddCompanyDialog(QDialog):
    """협력사 추가 다이얼로그"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("협력사 추가")
        self.setFixedSize(400, 200)
        
        from PySide6.QtWidgets import QFormLayout
        
        layout = QFormLayout()
        
        self.sap_code_edit = QLineEdit()
        self.sap_code_edit.setPlaceholderText("예: B907")
        layout.addRow("협력사 코드 *:", self.sap_code_edit)
        
        self.sap_name_edit = QLineEdit()
        self.sap_name_edit.setPlaceholderText("예: AMS")
        layout.addRow("협력사 이름 *:", self.sap_name_edit)
        
        self.renault_code_edit = QLineEdit()
        self.renault_code_edit.setPlaceholderText("예: 247736")
        layout.addRow("르노 코드 (선택사항):", self.renault_code_edit)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        ok_btn = QPushButton("확인")
        cancel_btn = QPushButton("취소")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addRow("", button_layout)
        
        self.setLayout(layout)
    
    def get_data(self) -> Dict[str, Any]:
        """입력 데이터 반환"""
        sap_code = self.sap_code_edit.text().strip()
        sap_name = self.sap_name_edit.text().strip()
        renault_code = self.renault_code_edit.text().strip()
        
        # rule_table_name 자동 생성
        rule_table_name = f"rule_{sap_code}"  # rule_협력사코드
        
        return {
            "sap_code": sap_code,
            "sap_name": sap_name,
            "renault_code": renault_code,
            "rule_table_name": rule_table_name,
        }


class EditCompanyDialog(QDialog):
    """협력사 수정 다이얼로그"""
    def __init__(self, company_info: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.setWindowTitle("협력사 수정")
        self.setFixedSize(400, 200)
        self.old_sap_code = company_info.get("sap_code", "")
        
        from PySide6.QtWidgets import QFormLayout
        
        layout = QFormLayout()
        
        self.sap_code_edit = QLineEdit()
        self.sap_code_edit.setText(company_info.get("sap_code", ""))
        self.sap_code_edit.setPlaceholderText("예: B907")
        layout.addRow("협력사 코드 *:", self.sap_code_edit)
        
        self.sap_name_edit = QLineEdit()
        self.sap_name_edit.setText(company_info.get("sap_name", ""))
        self.sap_name_edit.setPlaceholderText("예: AMS")
        layout.addRow("협력사 이름 *:", self.sap_name_edit)
        
        self.renault_code_edit = QLineEdit()
        self.renault_code_edit.setText(company_info.get("renault_code", ""))
        self.renault_code_edit.setPlaceholderText("예: 247736")
        layout.addRow("르노 코드 (선택사항):", self.renault_code_edit)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        ok_btn = QPushButton("확인")
        cancel_btn = QPushButton("취소")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addRow("", button_layout)
        
        self.setLayout(layout)
    
    def get_data(self) -> Dict[str, Any]:
        """입력 데이터 반환"""
        return {
            "old_sap_code": self.old_sap_code,
            "new_sap_code": self.sap_code_edit.text().strip(),
            "sap_name": self.sap_name_edit.text().strip(),
            "renault_code": self.renault_code_edit.text().strip(),
        }


class RuleManagementWidget(QWidget):
    """규칙 관리 위젯 (선택한 협력사의 규칙 추가/수정/삭제)"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_company: Optional[str] = None
        self.current_rule_table: Optional[str] = None
        self.current_sap_code: Optional[str] = None
        self.original_remark: str = ""  # 원본 remark 저장
        self.rules: List[Dict[str, Any]] = []
        self.priority_edit_mode: bool = False  # 우선순위 변경 모드 플래그
        self._drag_start_row: Optional[int] = None  # 드래그 시작 row 추적용
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # 제목
        from PySide6.QtWidgets import QLabel, QGroupBox, QTextEdit
        self.title_label = QLabel("협력사를 선택하세요")
        self.title_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        layout.addWidget(self.title_label)
        
        # 상단: Remark 영역
        remark_group = QGroupBox()
        remark_group.setStyleSheet("QGroupBox { border: 1px solid #ccc; border-radius: 3px; margin-top: 10px; padding-top: 10px; }")
        remark_layout = QVBoxLayout(remark_group)
        remark_layout.setContentsMargins(4, 8, 4, 8)
        remark_layout.setSpacing(6)
        
        # Remark 제목과 저장 버튼을 같은 줄에 배치 (우측 상단)
        remark_title_widget = QWidget()
        remark_title_layout = QHBoxLayout(remark_title_widget)
        remark_title_layout.setContentsMargins(0, 0, 0, 0)
        remark_title_layout.setSpacing(0)
        
        lbl_remark_title = QLabel("Remark")
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
        
        # Remark 텍스트 영역
        self.remark_text = QTextEdit()
        self.remark_text.setReadOnly(False)  # 편집 가능
        self.remark_text.setMaximumHeight(100)
        remark_layout.addWidget(self.remark_text)
        
        layout.addWidget(remark_group)
        
        # 버튼들
        button_layout = QHBoxLayout()
        self.btn_add_rule = QPushButton("+ 규칙 추가")
        self.btn_edit_rule = QPushButton("규칙 수정")
        self.btn_delete_rule = QPushButton("규칙 삭제")
        self.btn_add_rule.setEnabled(False)
        self.btn_edit_rule.setEnabled(False)
        self.btn_delete_rule.setEnabled(False)
        
        button_layout.addWidget(self.btn_add_rule)
        button_layout.addWidget(self.btn_edit_rule)
        button_layout.addWidget(self.btn_delete_rule)
        button_layout.addStretch()
        self.btn_priority_mode = QPushButton("우선순위 변경")
        self.btn_priority_mode.setEnabled(False)
        button_layout.addWidget(self.btn_priority_mode)
        layout.addLayout(button_layout)
        
        # 하단: Rule 테이블 전체 출력
        rule_group = QGroupBox("규칙 테이블")
        rule_group.setStyleSheet("QGroupBox::title { color: black; font-size: 12pt; font-weight: bold; }")
        rule_layout = QVBoxLayout()
        self.table = QTableWidget()
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)  # 다중 선택 가능
        # 기본적으로 드래그 앤 드롭 비활성화 (우선순위 변경 모드에서만 활성화)
        self.table.setDragDropMode(QTableWidget.NoDragDrop)
        rule_layout.addWidget(self.table)
        rule_group.setLayout(rule_layout)
        layout.addWidget(rule_group, 1)
        
        self.setLayout(layout)
        
        # 이벤트 연결
        self.btn_add_rule.clicked.connect(self.on_add_rule)
        self.btn_edit_rule.clicked.connect(self.on_edit_rule)
        self.btn_delete_rule.clicked.connect(self.on_delete_rule)
        self.btn_priority_mode.clicked.connect(self.on_toggle_priority_mode)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
        self.btn_save_remark.clicked.connect(self.on_save_remark)
        self.remark_text.textChanged.connect(self.on_remark_changed)
    
    def set_company(self, company_name: str):
        """협력사 설정 및 규칙 로드"""
        self.current_company = company_name
        
        # 빈 문자열인 경우 초기화
        if not company_name or not company_name.strip():
            self.title_label.setText("협력사를 선택하세요")
            self.current_rule_table = None
            self.current_sap_code = None
            self.rules = []
            self.remark_text.clear()
            self.original_remark = ""
            self.refresh_table()
            self.btn_add_rule.setEnabled(False)
            self.btn_priority_mode.setEnabled(False)
            return
        
        company_info = get_company_info(company_name)
        
        if not company_info:
            self.title_label.setText(f"오류: {company_name} 정보를 찾을 수 없습니다")
            self.current_rule_table = None
            self.rules = []
            self.remark_text.clear()
            self.refresh_table()
            return
        
        self.current_rule_table = company_info.get("rule_table_name")
        self.title_label.setText(f"규칙 관리 - {company_name} ({self.current_rule_table or '테이블 없음'})")
        
        # Remark 표시
        remark = company_info.get("remark", "")
        self.original_remark = remark if remark else ""
        self.remark_text.setText(self.original_remark)
        self.current_sap_code = company_info.get("sap_code")
        self.btn_save_remark.setEnabled(False)  # 초기에는 저장 버튼 비활성화
        
        if self.current_rule_table:
            self.rules = get_rules_from_table(self.current_rule_table)
        else:
            self.rules = []
        
        self.refresh_table()
        self.btn_add_rule.setEnabled(self.current_rule_table is not None)
        self.btn_priority_mode.setEnabled(self.current_rule_table is not None and len(self.rules) > 0)
        
        # 협력사 변경 시 우선순위 변경 모드 해제
        if self.priority_edit_mode:
            self.priority_edit_mode = False
            self.table.setDragDropMode(QTableWidget.NoDragDrop)
            self.btn_priority_mode.setText("우선순위 변경")
    
    def refresh_table(self):
        """테이블 새로고침 (rule 테이블 전체 컬럼 출력)"""
        if not self.rules:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return
        
        # 컬럼명 한국어 매핑
        column_name_map = {
            "rule_id": "규칙 ID",
            "priority": "우선순위",
            "status": "규칙적용상태",
            "repair_region": "수리 지역",
            "project_code": "프로젝트 코드",
            "exclude_project_code": "제외 프로젝트 코드",
            "vehicle_classification": "차계",
            "part_no": "부품번호",
            "part_name": "부품명",
            "engine_form": "엔진 형식",
            "warranty_mileage_override": "보증 주행거리 상한",
            "warranty_period_override": "보증 기간",
            "liability_ratio": "구상율",
            "amount_cap_type": "금액 상한 유형",
            "amount_cap_value": "상한 금액",
            "valid_from": "적용 시작일",
            "valid_to": "적용 종료일",
            "created_at": "생성일시",
            "updated_at": "수정일시",
        }
        
        # 모든 컬럼 가져오기
        all_columns = set()
        for rule in self.rules:
            all_columns.update(rule.keys())
        
        # 컬럼 순서 정렬 (rule_id, priority, status 등을 앞에)
        column_order = ["rule_id", "priority", "status", "repair_region", "project_code","exclude_project_code",
                       "vehicle_classification", "part_name", "part_no", 
                       "liability_ratio", "amount_cap_type", "amount_cap_value",
                       "warranty_mileage_override", "warranty_period_override",
                       "valid_from", "valid_to", "engine_form",
                       "created_at", "updated_at"]
        
        # 순서가 정해진 컬럼 먼저, 나머지는 알파벳 순
        ordered_columns = []
        for col in column_order:
            if col in all_columns:
                ordered_columns.append(col)
                all_columns.remove(col)
        
        ordered_columns.extend(sorted(all_columns))
        
        # 한국어 헤더 레이블 생성
        korean_headers = [column_name_map.get(col, col) for col in ordered_columns]
        
        # 테이블 설정
        self.table.setColumnCount(len(ordered_columns))
        self.table.setHorizontalHeaderLabels(korean_headers)
        self.table.setRowCount(len(self.rules))
        
        # 데이터 채우기
        for row, rule in enumerate(self.rules):
            for col_idx, col_name in enumerate(ordered_columns):
                value = rule.get(col_name)
                
                if value is None:
                    item = QTableWidgetItem("")
                elif isinstance(value, (int, float)):
                    item = QTableWidgetItem(str(value))
                elif isinstance(value, bool):
                    item = QTableWidgetItem("TRUE" if value else "FALSE")
                else:
                    item = QTableWidgetItem(str(value))
                
                # 상태 컬럼은 색상 표시
                if col_name == "status":
                    status = str(value).upper()
                    item.setTextAlignment(Qt.AlignCenter)
                    if status == "ACTIVE":
                        item.setForeground(Qt.GlobalColor.green)
                    elif status == "INACTIVE":
                        item.setForeground(Qt.GlobalColor.gray)
                
                # 숫자 컬럼은 우측 정렬
                if col_name in ["rule_id", "priority", "liability_ratio", "amount_cap_value",
                               "warranty_mileage_override", "warranty_period_override"]:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                
                self.table.setItem(row, col_idx, item)
        
        self.table.resizeColumnsToContents()
    
    def format_rule_changes(self, rule: Dict[str, Any]) -> str:
        """Rule의 변경점만 포맷팅"""
        changes = []
        
        project_code = rule.get("project_code", "").strip()
        if project_code:
            changes.append(f"프로젝트: {project_code}")
        
        part_name = rule.get("part_name", "").strip()
        if part_name:
            changes.append(f"부품: {part_name}")
        
        liability_ratio = rule.get("liability_ratio", 0)
        if liability_ratio is not None:
            changes.append(f"구상율: {liability_ratio}%")
        
        return " | ".join(changes) if changes else "기본 규칙"
    
    def on_selection_changed(self):
        """선택 변경 시"""
        has_selection = len(self.table.selectedItems()) > 0
        self.btn_edit_rule.setEnabled(has_selection and self.current_rule_table is not None)
        self.btn_delete_rule.setEnabled(has_selection and self.current_rule_table is not None and not self.priority_edit_mode)
    
    def on_add_rule(self):
        """규칙 추가"""
        if not self.current_rule_table:
            QMessageBox.warning(self, "오류", "Rule 테이블이 없습니다.")
            return
        
        dialog = AddRuleDialog(self.current_rule_table, self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            try:
                rule_id = add_rule_to_table(
                    rule_table_name=self.current_rule_table,
                    priority=data.get("priority"),
                    status=data["status"],
                    repair_region=data["repair_region"],
                    vehicle_classification=data["vehicle_classification"],
                    liability_ratio=data["liability_ratio"],
                    amount_cap_type=data["amount_cap_type"],
                    project_code=data["project_code"],
                    part_name=data["part_name"],
                    part_no=data["part_no"],
                    engine_form=data.get("engine_form", "ALL"),
                    exclude_project_code=data.get("exclude_project_code"),
                    warranty_mileage_override=data.get("warranty_mileage_override"),
                    warranty_period_override=data.get("warranty_period_override"),
                    amount_cap_value=data.get("amount_cap_value"),
                    valid_from=data.get("valid_from"),
                    valid_to=data.get("valid_to"),
                )
                
                QMessageBox.information(self, "완료", f"규칙이 추가되었습니다. (ID: {rule_id})")
                self.set_company(self.current_company)  # 새로고침
            except Exception as e:
                QMessageBox.critical(self, "오류", f"규칙 추가 실패: {str(e)}")
    
    def on_edit_rule(self):
        """규칙 수정"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        row = selected_items[0].row()
        if row < 0 or row >= len(self.rules):
            return
        
        rule = self.rules[row]
        rule_id = rule.get("rule_id")
        
        if not rule_id:
            QMessageBox.warning(self, "오류", "규칙 ID를 찾을 수 없습니다.")
            return
        
        if not self.current_rule_table:
            QMessageBox.warning(self, "오류", "Rule 테이블이 없습니다.")
            return
        
        # 수정 다이얼로그 열기
        dialog = AddRuleDialog(self.current_rule_table, self, rule_data=rule)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            try:
                success = update_rule_in_table(
                    rule_table_name=self.current_rule_table,
                    rule_id=rule_id,
                    priority=data.get("priority"),
                    status=data.get("status"),
                    repair_region=data.get("repair_region"),
                    vehicle_classification=data.get("vehicle_classification"),
                    liability_ratio=data.get("liability_ratio"),
                    amount_cap_type=data.get("amount_cap_type"),
                    project_code=data.get("project_code"),
                    part_name=data.get("part_name"),
                    part_no=data.get("part_no"),
                    engine_form=data.get("engine_form"),
                    exclude_project_code=data.get("exclude_project_code"),
                    warranty_mileage_override=data.get("warranty_mileage_override"),
                    warranty_period_override=data.get("warranty_period_override"),
                    amount_cap_value=data.get("amount_cap_value"),
                    valid_from=data.get("valid_from"),
                    valid_to=data.get("valid_to"),
                )
                
                if success:
                    QMessageBox.information(self, "완료", "규칙이 수정되었습니다.")
                    self.set_company(self.current_company)  # 새로고침
                else:
                    QMessageBox.warning(self, "오류", "규칙 수정에 실패했습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"규칙 수정 실패: {str(e)}")
    
    def on_delete_rule(self):
        """규칙 삭제 (다중 선택 지원)"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        # 선택된 행들의 rule_id 수집
        selected_rows = set()
        for item in selected_items:
            row = item.row()
            if 0 <= row < len(self.rules):
                selected_rows.add(row)
        
        if not selected_rows:
            return
        
        # 선택된 규칙들의 rule_id와 정보 수집
        rule_ids_to_delete = []
        rules_info = []
        for row in selected_rows:
            rule = self.rules[row]
            rule_id = rule.get("rule_id")
            if rule_id:
                rule_ids_to_delete.append(rule_id)
                rules_info.append({
                    "rule_id": rule_id,
                    "priority": rule.get("priority"),
                    "status": rule.get("status")
                })
        
        if not rule_ids_to_delete:
            QMessageBox.warning(self, "오류", "삭제할 규칙을 찾을 수 없습니다.")
            return
        
        if not self.current_rule_table:
            QMessageBox.warning(self, "오류", "Rule 테이블이 없습니다.")
            return
        
        # 확인 메시지 (단일/다중 구분)
        if len(rule_ids_to_delete) == 1:
            rule_info = rules_info[0]
            message = f"이 규칙을 삭제하시겠습니까?\n(우선순위: {rule_info['priority']}, 상태: {rule_info['status']})"
        else:
            message = f"선택한 {len(rule_ids_to_delete)}개의 규칙을 삭제하시겠습니까?"
        
        reply = QMessageBox.question(
            self, "확인", 
            message,
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                # 모든 선택된 규칙 삭제
                success_count = 0
                failed_count = 0
                
                for rule_id in rule_ids_to_delete:
                    success = delete_rule_from_table(self.current_rule_table, rule_id)
                    if success:
                        success_count += 1
                    else:
                        failed_count += 1
                
                # 결과 메시지
                if success_count > 0:
                    # 삭제 후 남은 룰들의 우선순위 재정렬 (1, 2, 3...)
                    try:
                        # 남은 모든 룰 조회 (get_rules_from_table은 이미 priority 순서대로 반환)
                        remaining_rules = get_rules_from_table(self.current_rule_table)
                        if remaining_rules:
                            # priority 순서대로 정렬된 rule_id 리스트 생성
                            rule_ids_in_order = [r.get("rule_id") for r in remaining_rules if r.get("rule_id")]
                            
                            # 우선순위 재정렬 (1, 2, 3...)
                            if rule_ids_in_order:
                                update_rule_priorities(self.current_rule_table, rule_ids_in_order)
                    except Exception as e:
                        # 우선순위 재정렬 실패해도 삭제는 성공했으므로 경고만 표시
                        print(f"우선순위 재정렬 실패: {str(e)}")
                    
                    if len(rule_ids_to_delete) == 1:
                        QMessageBox.information(self, "완료", "규칙이 삭제되었습니다.")
                    else:
                        QMessageBox.information(self, "완료", f"{success_count}개의 규칙이 삭제되었습니다.")
                    self.set_company(self.current_company)  # 새로고침
                
                if failed_count > 0:
                    QMessageBox.warning(self, "경고", f"{failed_count}개의 규칙 삭제에 실패했습니다.")
                    
            except Exception as e:
                QMessageBox.critical(self, "오류", f"규칙 삭제 실패: {str(e)}")
    
    def on_remark_changed(self):
        """Remark 텍스트 변경 시 저장 버튼 활성화"""
        current_text = self.remark_text.toPlainText()
        if current_text != self.original_remark:
            self.btn_save_remark.setEnabled(True)
        else:
            self.btn_save_remark.setEnabled(False)
    
    def on_save_remark(self):
        """Remark 저장"""
        if not self.current_sap_code:
            QMessageBox.warning(self, "오류", "협력사 정보를 찾을 수 없습니다.")
            return
        
        new_remark = self.remark_text.toPlainText()
        
        try:
            success = update_company_remark(self.current_sap_code, new_remark)
            if success:
                self.original_remark = new_remark
                self.btn_save_remark.setEnabled(False)
                QMessageBox.information(self, "완료", "Remark가 저장되었습니다.")
            else:
                QMessageBox.warning(self, "오류", "Remark 저장에 실패했습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"Remark 저장 실패: {str(e)}")
    
    def on_toggle_priority_mode(self):
        """우선순위 변경 모드 토글"""
        if not self.current_rule_table or len(self.rules) == 0:
            return
        
        self.priority_edit_mode = not self.priority_edit_mode
        
        if self.priority_edit_mode:
            # 우선순위 변경 모드 활성화
            self.table.setDragDropMode(QTableWidget.InternalMove)
            self.table.setDragDropOverwriteMode(False)
            self.table.setDefaultDropAction(Qt.MoveAction)
            self.btn_priority_mode.setText("우선순위 변경 종료")
            self.btn_add_rule.setEnabled(False)
            self.btn_edit_rule.setEnabled(False)
            self.btn_delete_rule.setEnabled(False)
            QMessageBox.information(self, "안내", "드래그 앤 드롭으로 규칙 순서를 변경할 수 있습니다.\n모드 종료 시 변경사항이 저장됩니다.")
            
            # dropEvent 오버라이드하여 row 이동을 수동으로 처리
            # 드래그 시작 시점의 row를 저장
            self._drag_start_row = None
            original_start_drag = self.table.startDrag
            def handle_start_drag(supported_actions):
                self._drag_start_row = self.table.currentRow()
                original_start_drag(supported_actions)
            self.table.startDrag = handle_start_drag
            
            original_drop = self.table.dropEvent
            def handle_drop(event):
                if event.source() == self.table and event.dropAction() == Qt.MoveAction:
                    if self._drag_start_row is None or self._drag_start_row < 0:
                        event.ignore()
                        return
                    
                    drag_row = self._drag_start_row
                    
                    # 드롭 위치 계산
                    drop_pos = event.pos()
                    drop_row = self.table.rowAt(drop_pos.y())
                    
                    if drop_row < 0:
                        drop_row = self.table.rowCount()
                    else:
                        # 행의 중간 위치 확인
                        item_rect = self.table.visualItemRect(self.table.item(drop_row, 0))
                        if drop_pos.y() > item_rect.center().y():
                            drop_row += 1
                    
                    if drag_row == drop_row or drop_row == drag_row + 1:
                        event.ignore()
                        return
                    
                    # 전체 데이터 가져오기
                    all_rows = []
                    for r in range(self.table.rowCount()):
                        row_items = []
                        for c in range(self.table.columnCount()):
                            item = self.table.item(r, c)
                            row_items.append(item.clone() if item else None)
                        all_rows.append(row_items)
                    
                    # 순서 변경
                    moved = all_rows.pop(drag_row)
                    if drop_row > drag_row:
                        drop_row -= 1
                    all_rows.insert(drop_row, moved)
                    
                    # 테이블 재구성
                    self.table.clearContents()
                    self.table.setRowCount(len(all_rows))
                    for r, row_items in enumerate(all_rows):
                        for c, item in enumerate(row_items):
                            if item:
                                self.table.setItem(r, c, item)
                    
                    self._drag_start_row = None
                    event.accept()
                else:
                    original_drop(event)
            self.table.dropEvent = handle_drop
        else:
            # 우선순위 변경 모드 비활성화 - 변경사항 DB에 반영
            self._save_priority_changes()
            self.table.setDragDropMode(QTableWidget.NoDragDrop)
            self.btn_priority_mode.setText("우선순위 변경")
            self.btn_add_rule.setEnabled(self.current_rule_table is not None)
            # 편집/삭제 버튼은 선택 상태에 따라 활성화
            self.on_selection_changed()
            
            # 오버라이드 제거
            self._drag_start_row = None
    
    def _save_priority_changes(self):
        """현재 테이블 순서를 DB에 반영"""
        if not self.current_rule_table or not self.rules:
            return
        
        # 현재 테이블의 순서대로 rule_id 추출
        rule_ids_in_order = []
        for row in range(self.table.rowCount()):
            # rule_id 컬럼 찾기
            rule_id_col = None
            for col in range(self.table.columnCount()):
                header = self.table.horizontalHeaderItem(col)
                if header and header.text() == "규칙 ID":
                    rule_id_col = col
                    break
            
            if rule_id_col is not None:
                item = self.table.item(row, rule_id_col)
                if item:
                    try:
                        rule_id = int(item.text())
                        rule_ids_in_order.append(rule_id)
                    except ValueError:
                        pass
        
        # 순서가 변경되었는지 확인
        if len(rule_ids_in_order) != len(self.rules):
            return
        
        # 기존 순서와 비교
        current_order = [r.get("rule_id") for r in self.rules]
        if rule_ids_in_order == current_order:
            return  # 순서 변경 없음
        
        # 우선순위 업데이트
        try:
            update_rule_priorities(self.current_rule_table, rule_ids_in_order)
            # 규칙 목록 새로고침
            self.set_company(self.current_company)
            QMessageBox.information(self, "완료", "우선순위가 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", f"우선순위 업데이트 실패: {str(e)}")
            # 실패 시 원래대로 복구
            self.set_company(self.current_company)


class ComExManagementPageWidget(QWidget):
    """comex 관리 페이지"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.company_data = {}  # sap_name -> {sap_code, sap_name} 매핑
        
        layout = QHBoxLayout()
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(8)
        
        # 왼쪽: 협력사 목록 패널
        left_panel = QVBoxLayout()
        left_panel.setContentsMargins(0, 0, 0, 0)
        left_panel.setSpacing(4)
        
        # 버튼들
        button_layout = QHBoxLayout()
        self.btn_add_company = QPushButton("+ 협력사 추가")
        button_layout.addWidget(self.btn_add_company)
        left_panel.addLayout(button_layout)
        
        # 검색
        search_layout = QHBoxLayout()
        from PySide6.QtWidgets import QLabel
        search_layout.addWidget(QLabel("검색:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("협력사 검색")
        search_layout.addWidget(self.search_edit)
        left_panel.addLayout(search_layout)
        
        # 협력사 목록
        self.company_list = QListWidget()
        self.company_list.setMaximumWidth(250)
        self.company_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.company_list.customContextMenuRequested.connect(self.on_company_context_menu)
        left_panel.addWidget(self.company_list, 1)
        
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        layout.addWidget(left_widget)
        
        # 오른쪽: 규칙 관리 위젯
        self.rule_management = RuleManagementWidget(self)
        layout.addWidget(self.rule_management, 1)
        
        self.setLayout(layout)
        
        # 이벤트 연결
        self.btn_add_company.clicked.connect(self.on_add_company)
        self.company_list.itemClicked.connect(self.on_company_selected)
        self.search_edit.textChanged.connect(self.on_search_changed)
        
        # 초기화
        self.load_companies()
    
    def load_companies(self):
        """협력사 목록 로드 (sap_code와 sap_name 저장)"""
        self.company_list.clear()
        self.company_data = {}  # sap_name -> {sap_code, sap_name} 매핑
        
        companies = get_all_companies_with_code()
        
        for company in companies:
            sap_name = company["sap_name"]
            sap_code = company["sap_code"]
            
            # 표시는 sap_name으로, 데이터는 모두 저장
            item = QListWidgetItem(sap_name)
            self.company_list.addItem(item)
            self.company_data[sap_name] = {"sap_code": sap_code, "sap_name": sap_name}
        
        # 검색 필터 적용
        self.on_search_changed(self.search_edit.text())
    
    def on_search_changed(self, text: str):
        """검색어 변경 시 필터링 (대소문자 구분 없이, sap_code와 sap_name 모두 검색)"""
        search_text = text.strip().lower()
        
        if not search_text:
            # 검색어가 없으면 모두 표시
            for i in range(self.company_list.count()):
                self.company_list.item(i).setHidden(False)
            return
        
        for i in range(self.company_list.count()):
            item = self.company_list.item(i)
            sap_name = item.text()
            company_info = self.company_data.get(sap_name, {})
            sap_code = company_info.get("sap_code", "")
            
            # sap_name과 sap_code 모두 검색 (대소문자 구분 없음)
            sap_name_lower = sap_name.lower()
            sap_code_lower = sap_code.lower()
            
            matches = (
                search_text in sap_name_lower or 
                search_text in sap_code_lower
            )
            
            item.setHidden(not matches)
    
    def on_add_company(self):
        """협력사 추가"""
        dialog = AddCompanyDialog(self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            if not data["sap_code"]:
                QMessageBox.warning(self, "오류", "협력사 코드를 입력해주세요.")
                return
            if not data["sap_name"]:
                QMessageBox.warning(self, "오류", "협력사 이름을 입력해주세요.")
                return
            
            try:
                upsert_company(
                    sap_code=data["sap_code"],
                    sap_name=data["sap_name"],
                    renault_code=data["renault_code"],
                    rule_table_name=data["rule_table_name"],
                )
                QMessageBox.information(self, "완료", "협력사가 추가되었습니다.")
                self.load_companies()
            except Exception as e:
                QMessageBox.critical(self, "오류", f"협력사 추가 실패: {str(e)}")
    
    def on_company_selected(self, item: QListWidgetItem):
        """협력사 선택 시"""
        company_name = item.text()
        self.rule_management.set_company(company_name)
    
    def on_company_context_menu(self, pos):
        """협력사 목록 우클릭 메뉴"""
        item = self.company_list.itemAt(pos)
        if not item:
            return
        
        company_name = item.text()
        company_info = self.company_data.get(company_name, {})
        sap_code = company_info.get("sap_code")
        
        if not sap_code:
            return
        
        # 전체 회사 정보 가져오기
        full_company_info = get_company_info(sap_code)
        if not full_company_info:
            return
        
        menu = QMenu(self)
        act_edit = menu.addAction("수정")
        act_delete = menu.addAction("삭제")
        
        picked = menu.exec(self.company_list.mapToGlobal(pos))
        
        if picked == act_edit:
            self.on_edit_company(full_company_info)
        elif picked == act_delete:
            self.on_delete_company(full_company_info)
    
    def on_edit_company(self, company_info: Dict[str, Any]):
        """협력사 수정"""
        dialog = EditCompanyDialog(company_info, self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            if not data["new_sap_code"]:
                QMessageBox.warning(self, "오류", "협력사 코드를 입력해주세요.")
                return
            if not data["sap_name"]:
                QMessageBox.warning(self, "오류", "협력사 이름을 입력해주세요.")
                return
            
            try:
                success = update_company(
                    old_sap_code=data["old_sap_code"],
                    new_sap_code=data["new_sap_code"],
                    sap_name=data["sap_name"],
                    renault_code=data["renault_code"],
                )
                
                if success:
                    QMessageBox.information(self, "완료", "협력사가 수정되었습니다.")
                    # 현재 선택된 회사가 수정된 경우 규칙 관리도 새로고침
                    current_item = self.company_list.currentItem()
                    if current_item and current_item.text() == company_info.get("sap_name"):
                        # SAP 코드가 변경된 경우 새 이름으로 선택
                        if data["new_sap_code"] != data["old_sap_code"]:
                            # 목록 새로고침 후 새 이름으로 선택
                            self.load_companies()
                            # 새 이름으로 항목 찾기
                            for i in range(self.company_list.count()):
                                item = self.company_list.item(i)
                                if item.text() == data["sap_name"]:
                                    self.company_list.setCurrentItem(item)
                                    self.on_company_selected(item)
                                    break
                        else:
                            # 이름만 변경된 경우
                            self.load_companies()
                            for i in range(self.company_list.count()):
                                item = self.company_list.item(i)
                                if item.text() == data["sap_name"]:
                                    self.company_list.setCurrentItem(item)
                                    self.on_company_selected(item)
                                    break
                    else:
                        self.load_companies()
                else:
                    QMessageBox.warning(self, "오류", "협력사 수정에 실패했습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"협력사 수정 실패: {str(e)}")
    
    def on_delete_company(self, company_info: Dict[str, Any]):
        """협력사 삭제"""
        sap_code = company_info.get("sap_code")
        sap_name = company_info.get("sap_name")
        
        reply = QMessageBox.question(
            self, "확인",
            f"'{sap_name}' ({sap_code}) 협력사를 삭제하시겠습니까?\n\n"
            "주의: 이 작업은 되돌릴 수 없으며, 관련 rule 테이블도 함께 삭제됩니다.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                success = delete_company(sap_code)
                
                if success:
                    QMessageBox.information(self, "완료", "협력사가 삭제되었습니다.")
                    # 규칙 관리 초기화
                    self.rule_management.set_company("")
                    # 목록 새로고침
                    self.load_companies()
                else:
                    QMessageBox.warning(self, "오류", "협력사 삭제에 실패했습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"협력사 삭제 실패: {str(e)}")

