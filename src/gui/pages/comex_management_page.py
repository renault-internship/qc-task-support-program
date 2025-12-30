"""
comex 관리 페이지 - 협력사 목록 및 룰 관리
"""
from typing import Dict, Any, Optional, List

from PySide6.QtCore import Qt, QSortFilterProxyModel, QRegularExpression
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit,
    QListWidget, QListWidgetItem, QMessageBox, QDialog, QTableWidget,
    QTableWidgetItem, QHeaderView
)

from src.database import (
    get_all_companies, get_all_companies_with_code, get_company_info, 
    get_rules_from_table, add_rule_to_table, update_rule_in_table, 
    delete_rule_from_table, upsert_company
)
from src.gui.dialogs import AddRuleDialog


class AddCompanyDialog(QDialog):
    """협력사 추가 다이얼로그"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("협력사 추가")
        self.setFixedSize(400, 300)
        
        from PySide6.QtWidgets import QFormLayout, QSpinBox
        
        layout = QFormLayout()
        
        self.sap_code_edit = QLineEdit()
        self.sap_code_edit.setPlaceholderText("예: B907")
        layout.addRow("SAP 코드 *:", self.sap_code_edit)
        
        self.sap_name_edit = QLineEdit()
        self.sap_name_edit.setPlaceholderText("예: AMS")
        layout.addRow("SAP 기업명 *:", self.sap_name_edit)
        
        self.warranty_mileage_spin = QSpinBox()
        self.warranty_mileage_spin.setRange(0, 1000000)
        self.warranty_mileage_spin.setValue(50000)
        layout.addRow("보증 주행거리 (km):", self.warranty_mileage_spin)
        
        self.warranty_period_spin = QSpinBox()
        self.warranty_period_spin.setRange(0, 100)
        self.warranty_period_spin.setValue(3)
        self.warranty_period_spin.setSuffix(" 년")
        layout.addRow("보증 기간:", self.warranty_period_spin)
        
        self.rule_table_edit = QLineEdit()
        self.rule_table_edit.setPlaceholderText("예: rule_B907 (자동 생성됨)")
        layout.addRow("Rule 테이블명:", self.rule_table_edit)
        
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
        rule_table = self.rule_table_edit.text().strip()
        
        # rule_table이 비어있으면 자동 생성
        if not rule_table:
            rule_table = f"rule_{sap_code}"
        
        return {
            "sap_code": sap_code,
            "sap_name": sap_name,
            "warranty_mileage": self.warranty_mileage_spin.value(),
            "warranty_period": self.warranty_period_spin.value() * 365,  # 년을 일로 변환
            "rule_table_name": rule_table,
        }


class RuleManagementWidget(QWidget):
    """룰 관리 위젯 (선택한 협력사의 룰 추가/수정/삭제)"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_company: Optional[str] = None
        self.current_rule_table: Optional[str] = None
        self.rules: List[Dict[str, Any]] = []
        
        layout = QVBoxLayout()
        
        # 제목
        from PySide6.QtWidgets import QLabel
        self.title_label = QLabel("협력사를 선택하세요")
        self.title_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(self.title_label)
        
        # 버튼들
        button_layout = QHBoxLayout()
        self.btn_add_rule = QPushButton("+ 룰 추가")
        self.btn_edit_rule = QPushButton("룰 수정")
        self.btn_delete_rule = QPushButton("룰 삭제")
        self.btn_add_rule.setEnabled(False)
        self.btn_edit_rule.setEnabled(False)
        self.btn_delete_rule.setEnabled(False)
        
        button_layout.addWidget(self.btn_add_rule)
        button_layout.addWidget(self.btn_edit_rule)
        button_layout.addWidget(self.btn_delete_rule)
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # 룰 목록 테이블
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["우선순위", "상태", "수리지역", "변경점"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        layout.addWidget(self.table, 1)
        
        self.setLayout(layout)
        
        # 이벤트 연결
        self.btn_add_rule.clicked.connect(self.on_add_rule)
        self.btn_edit_rule.clicked.connect(self.on_edit_rule)
        self.btn_delete_rule.clicked.connect(self.on_delete_rule)
        self.table.itemSelectionChanged.connect(self.on_selection_changed)
    
    def set_company(self, company_name: str):
        """협력사 설정 및 룰 로드"""
        self.current_company = company_name
        company_info = get_company_info(company_name)
        
        if not company_info:
            self.title_label.setText(f"오류: {company_name} 정보를 찾을 수 없습니다")
            self.current_rule_table = None
            self.rules = []
            self.refresh_table()
            return
        
        self.current_rule_table = company_info.get("rule_table_name")
        self.title_label.setText(f"룰 관리 - {company_name} ({self.current_rule_table or '테이블 없음'})")
        
        if self.current_rule_table:
            self.rules = get_rules_from_table(self.current_rule_table)
        else:
            self.rules = []
        
        self.refresh_table()
        self.btn_add_rule.setEnabled(self.current_rule_table is not None)
    
    def refresh_table(self):
        """테이블 새로고침"""
        self.table.setRowCount(len(self.rules))
        
        for row, rule in enumerate(self.rules):
            # 우선순위
            priority_item = QTableWidgetItem(str(rule.get("priority", "")))
            priority_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, priority_item)
            
            # 상태
            status = rule.get("status", "")
            status_item = QTableWidgetItem(status)
            status_item.setTextAlignment(Qt.AlignCenter)
            if status.upper() == "ACTIVE":
                status_item.setForeground(Qt.GlobalColor.green)
            elif status.upper() == "INACTIVE":
                status_item.setForeground(Qt.GlobalColor.gray)
            self.table.setItem(row, 1, status_item)
            
            # 수리지역
            region_item = QTableWidgetItem(rule.get("repair_region", ""))
            self.table.setItem(row, 2, region_item)
            
            # 변경점
            changes = self.format_rule_changes(rule)
            changes_item = QTableWidgetItem(changes)
            self.table.setItem(row, 3, changes_item)
        
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
        self.btn_delete_rule.setEnabled(has_selection and self.current_rule_table is not None)
    
    def on_add_rule(self):
        """룰 추가"""
        if not self.current_rule_table:
            QMessageBox.warning(self, "오류", "Rule 테이블이 없습니다.")
            return
        
        dialog = AddRuleDialog(self.current_rule_table, self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            try:
                rule_id = add_rule_to_table(
                    rule_table_name=self.current_rule_table,
                    priority=data["priority"],
                    status=data["status"],
                    repair_region=data["repair_region"],
                    vehicle_classification=data["vehicle_classification"],
                    liability_ratio=data["liability_ratio"],
                    amount_cap_type=data["amount_cap_type"],
                    project_code=data["project_code"],
                    part_name=data["part_name"],
                    part_no=data["part_no"],
                    exclude_project_code=data["exclude_project_code"],
                    amount_cap_value=data["amount_cap_value"],
                    note=data["note"],
                )
                
                QMessageBox.information(self, "완료", f"룰이 추가되었습니다. (ID: {rule_id})")
                self.set_company(self.current_company)  # 새로고침
            except Exception as e:
                QMessageBox.critical(self, "오류", f"룰 추가 실패: {str(e)}")
    
    def on_edit_rule(self):
        """룰 수정"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        row = selected_items[0].row()
        if row < 0 or row >= len(self.rules):
            return
        
        rule = self.rules[row]
        rule_id = rule.get("rule_id")
        
        if not rule_id:
            QMessageBox.warning(self, "오류", "룰 ID를 찾을 수 없습니다.")
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
                    exclude_project_code=data.get("exclude_project_code"),
                    amount_cap_value=data.get("amount_cap_value"),
                    note=data.get("note"),
                )
                
                if success:
                    QMessageBox.information(self, "완료", "룰이 수정되었습니다.")
                    self.set_company(self.current_company)  # 새로고침
                else:
                    QMessageBox.warning(self, "오류", "룰 수정에 실패했습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"룰 수정 실패: {str(e)}")
    
    def on_delete_rule(self):
        """룰 삭제"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        row = selected_items[0].row()
        if row < 0 or row >= len(self.rules):
            return
        
        rule = self.rules[row]
        rule_id = rule.get("rule_id")
        
        if not rule_id:
            QMessageBox.warning(self, "오류", "룰 ID를 찾을 수 없습니다.")
            return
        
        reply = QMessageBox.question(
            self, "확인", 
            f"이 룰을 삭제하시겠습니까?\n(우선순위: {rule.get('priority')}, 상태: {rule.get('status')})",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                success = delete_rule_from_table(self.current_rule_table, rule_id)
                if success:
                    QMessageBox.information(self, "완료", "룰이 삭제되었습니다.")
                    self.set_company(self.current_company)  # 새로고침
                else:
                    QMessageBox.warning(self, "오류", "룰 삭제에 실패했습니다.")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"룰 삭제 실패: {str(e)}")


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
        left_panel.addWidget(self.company_list, 1)
        
        left_widget = QWidget()
        left_widget.setLayout(left_panel)
        layout.addWidget(left_widget)
        
        # 오른쪽: 룰 관리 위젯
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
                QMessageBox.warning(self, "오류", "SAP 코드를 입력해주세요.")
                return
            if not data["sap_name"]:
                QMessageBox.warning(self, "오류", "SAP 기업명을 입력해주세요.")
                return
            
            try:
                upsert_company(
                    sap_code=data["sap_code"],
                    sap_name=data["sap_name"],
                    warranty_mileage=data["warranty_mileage"],
                    warranty_period=data["warranty_period"],
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

