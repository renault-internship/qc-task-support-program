"""
메인 페이지 - 엑셀 전처리 도구
"""
from pathlib import Path
from typing import Dict, Any

from PySide6.QtCore import Qt, QSortFilterProxyModel, QRegularExpression
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFileDialog, QMessageBox, QDialog
)

from openpyxl.workbook.workbook import Workbook

from src.utils import load_workbook_safe, save_workbook_safe, AppError
from src.excel_processor import preprocess_inplace
from src.database import (
    get_company_info, get_all_companies, 
    get_rules_from_table, add_rule_to_table
)
from src.gui.containers import (
    PreviewContainer, InfoPanel, ControlPanel
)
from src.gui.models import ExcelSheetModel
from src.gui.dialogs import AddRuleDialog, ViewRulesDialog


class MainPageWidget(QWidget):
    """메인 페이지 - 엑셀 전처리 도구"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.file_path: Path | None = None
        self.wb: Workbook | None = None
        self.model: ExcelSheetModel | None = None
        self.proxy: QSortFilterProxyModel | None = None
        self.current_company_info: Dict[str, Any] | None = None
        
        # 컨테이너 생성
        self.control_panel = ControlPanel(self)
        self.preview_container = PreviewContainer(self)
        self.info_panel = InfoPanel(self)
        
        # 레이아웃 구성
        layout = QVBoxLayout()
        
        # 상단: 컨트롤 패널
        layout.addWidget(self.control_panel)
        layout.addSpacing(4)
        
        # 중앙: 미리보기 + 정보 패널
        preview_layout = QVBoxLayout()
        preview_layout.addWidget(self.preview_container, 1)
        preview_layout.addWidget(self.info_panel)
        preview_widget = QWidget()
        preview_widget.setLayout(preview_layout)
        layout.addWidget(preview_widget, 1)
        
        self.setLayout(layout)
        
        # 이벤트 연결
        self._connect_signals()
        
        # 초기화
        self._initialize()
    
    def _initialize(self):
        """초기화"""
        self._set_info_defaults()
        self.load_companies()
    
    def _connect_signals(self):
        """시그널 연결"""
        # 컨트롤 패널
        self.control_panel.get_upload_button().clicked.connect(self.open_file)
        self.control_panel.get_preprocess_button().clicked.connect(self.on_preprocess_clicked)
        self.control_panel.get_company_combo().currentTextChanged.connect(self._on_company_changed)
        self.control_panel.get_search_edit().textChanged.connect(self.on_search_changed)
        self.control_panel.get_edit_all_checkbox().stateChanged.connect(self.on_edit_mode_changed)
        
        # 미리보기 컨테이너
        self.preview_container.get_sheet_combo().currentTextChanged.connect(self.on_sheet_changed)
        
        # 정보 패널
        # self.info_panel.get_add_rule_button().clicked.connect(self.add_rule)
        self.info_panel.get_editable_label().mousePressEvent = self.show_rules_dialog
        
        # Export 버튼
        self.control_panel.get_export_final_button().clicked.connect(self.save_as_file)
    
    def _set_info_defaults(self):
        """정보 패널 기본값 설정"""
        # self.info_panel.set_company("-")
        self.info_panel.set_remark("-")
        self.info_panel.set_editable("-")
    
    
    def load_companies(self):
        """기업 목록 로드 (DB에서)"""
        combo = self.control_panel.get_company_combo()
        combo.clear()
        companies = get_all_companies()
        if companies:
            combo.addItem("선택")
            combo.addItems(companies)
        else:
            combo.addItem("선택")
    
    def _on_company_changed(self, name: str):
        """기업 선택 변경 시 정보 업데이트"""
        if name and name != "선택":
            # sap 테이블에서 기업 정보 가져오기
            company_info = get_company_info(name)
            if company_info:
                # self.info_panel.set_company(name)
                # 🔹🔹🔹 여기 추가 🔹🔹🔹
                remark = company_info.get("remark", "-")
                self.info_panel.set_remark(remark)
                # 🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹🔹
                
                # rule_table_name 가져오기
                rule_table_name = company_info.get("rule_table_name")
                if rule_table_name:
                    # rule 테이블에서 rule 정보 조회
                    rules = get_rules_from_table(rule_table_name)
                    if rules:
                        # rule 개수와 주요 정보 표시
                        rule_count = len(rules)
                        active_rules = [r for r in rules if r.get("status", "").upper() == "ACTIVE"]
                        active_count = len(active_rules)
                        self.info_panel.set_editable(f"Rule: {rule_count}개 (활성: {active_count}개)")
                    else:
                        self.info_panel.set_editable(f"Rule 테이블: {rule_table_name} (규칙 없음)")
                else:
                    self.info_panel.set_editable("Rule 테이블 없음")
                
                # 현재 선택된 기업 정보 저장 (rule 추가 시 사용)
                self.current_company_info = company_info
            else:
                # self.info_panel.set_company("-")
                self.info_panel.set_editable("-")
                self.current_company_info = None
        else:
            # self.info_panel.set_company("-")
            self.info_panel.set_editable("-")
            self.current_company_info = None
        
        # 편집 모드 라벨도 업데이트
        self._refresh_editable_label()
    
    def add_rule(self):
        """Rule 추가 다이얼로그 열기"""
        # 기업이 선택되어 있는지 확인
        company_name = self.control_panel.get_company_combo().currentText()
        if company_name == "선택" or not company_name:
            QMessageBox.warning(self, "오류", "먼저 기업을 선택해주세요.")
            return
        
        # 기업 정보 가져오기
        if not hasattr(self, 'current_company_info') or not self.current_company_info:
            company_info = get_company_info(company_name)
            if not company_info:
                QMessageBox.warning(self, "오류", f"기업정보를 찾을 수 없습니다: {company_name}")
                return
            self.current_company_info = company_info
        
        rule_table_name = self.current_company_info.get("rule_table_name")
        if not rule_table_name:
            QMessageBox.warning(self, "오류", "선택한 기업에 Rule 테이블이 없습니다.")
            return
        
        # Rule 추가 다이얼로그 열기
        dialog = AddRuleDialog(rule_table_name, self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            # 필수 필드 검증
            if not data["repair_region"]:
                QMessageBox.warning(self, "오류", "수리 지역을 입력해주세요.")
                return
            if not data["vehicle_classification"]:
                QMessageBox.warning(self, "오류", "차량 분류를 입력해주세요.")
                return
            
            try:
                rule_id = add_rule_to_table(
                    rule_table_name=rule_table_name,
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
                
                QMessageBox.information(self, "완료", f"Rule이 추가되었습니다. (ID: {rule_id})")
                
                # 정보 패널 업데이트 (rule 개수 갱신)
                self._on_company_changed(company_name)
            except Exception as e:
                QMessageBox.critical(self, "오류", f"Rule 추가 실패: {str(e)}")
    
    def show_rules_dialog(self, event):
        """Rule 목록 다이얼로그 표시"""
        company_name = self.control_panel.get_company_combo().currentText()
        if company_name == "선택" or not company_name:
            QMessageBox.information(self, "안내", "먼저 기업을 선택해주세요.")
            return
        
        company_info = get_company_info(company_name)
        if not company_info:
            QMessageBox.warning(self, "오류", f"기업정보를 찾을 수 없습니다: {company_name}")
            return
        
        rule_table_name = company_info.get("rule_table_name")
        if not rule_table_name:
            QMessageBox.information(self, "안내", "선택한 기업에 Rule 테이블이 없습니다.")
            return
        
        rules = get_rules_from_table(rule_table_name)
        if not rules:
            QMessageBox.information(self, "안내", "등록된 Rule이 없습니다.")
            return
        
        dialog = ViewRulesDialog(rules, self)
        dialog.exec()
    
    # ---------- 업로드 ----------
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 선택", "", "Excel Files (*.xlsx)")
        if not path:
            return

        self.file_path = Path(path)

        try:
            self.wb = load_workbook_safe(self.file_path)
        except AppError as e:
            QMessageBox.critical(self, "오류", str(e))
            return

        sheet_combo = self.preview_container.get_sheet_combo()
        sheet_combo.blockSignals(True)
        sheet_combo.clear()
        sheet_combo.addItems(self.wb.sheetnames)
        sheet_combo.blockSignals(False)

        if self.wb.sheetnames:
            sheet_combo.setCurrentIndex(0)
            self.load_sheet(self.wb.sheetnames[0])

        company = self.control_panel.get_company_combo().currentText()
        # self.info_panel.set_company(company if company != "선택" else "-")
        self.info_panel.set_remark("업로드 완료. 전처리 전 상태")
        self._refresh_editable_label()

    # ---------- 시트 로드/변경 ----------
    def load_sheet(self, sheet_name: str):
        if not self.wb:
            return

        ws = self.wb[sheet_name]
        self.model = ExcelSheetModel(ws, parent=self)

        # 편집 모드 반영
        edit_all = self.control_panel.get_edit_all_checkbox().isChecked()
        self.model.set_edit_all(edit_all)

        self.proxy = QSortFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy.setFilterKeyColumn(-1)  # 전체 컬럼 대상으로 검색

        table = self.preview_container.get_table()
        table.setModel(self.proxy)
        table.resizeColumnsToContents()

        # 기존 검색어 유지
        search_text = self.control_panel.get_search_edit().text()
        self.on_search_changed(search_text)

    def on_sheet_changed(self, sheet_name: str):
        if self.model:
            self.model.apply_dirty_to_sheet()
        self.load_sheet(sheet_name)

    # ---------- 검색/필터 ----------
    def on_search_changed(self, text: str):
        if not self.proxy:
            return
        t = (text or "").strip()
        if not t:
            self.proxy.setFilterRegularExpression(QRegularExpression(""))
            return
        rx = QRegularExpression(QRegularExpression.escape(t), QRegularExpression.CaseInsensitiveOption)
        self.proxy.setFilterRegularExpression(rx)

    # ---------- 편집 모드 ----------
    def on_edit_mode_changed(self):
        if self.model:
            edit_all = self.control_panel.get_edit_all_checkbox().isChecked()
            self.model.set_edit_all(edit_all)
            # flags만 바뀌는 거라 화면 갱신
            self.model.layoutChanged.emit()
        self._refresh_editable_label()

    def _refresh_editable_label(self):
        """편집 가능 라벨 업데이트"""
        # Rule 정보가 있으면 우선 표시, 없으면 편집 모드 표시
        company_name = self.control_panel.get_company_combo().currentText()
        if company_name and company_name != "선택":
            company_info = get_company_info(company_name)
            if company_info:
                rule_table_name = company_info.get("rule_table_name")
                if rule_table_name:
                    rules = get_rules_from_table(rule_table_name)
                    if rules:
                        rule_count = len(rules)
                        active_rules = [r for r in rules if r.get("status", "").upper() == "ACTIVE"]
                        active_count = len(active_rules)
                        self.info_panel.set_editable(f"Rule: {rule_count}개 (활성: {active_count}개)")
                        return
        
        # Rule 정보가 없으면 편집 모드 표시
        edit_all = self.control_panel.get_edit_all_checkbox().isChecked()
        if edit_all:
            self.info_panel.set_editable("현재: 전체 셀 편집 가능")
        else:
            # 구상율 컬럼 찾았는지 표시
            if self.model and self.model.editable_cols:
                cols = ", ".join(ExcelSheetModel.excel_col_name(c) for c in sorted(self.model.editable_cols))
                self.info_panel.set_editable(f"현재: 구상율 컬럼만 편집 가능 ({cols})")
            else:
                self.info_panel.set_editable("현재: 편집 제한(구상율 컬럼 미탐지 또는 없음)")

    # ---------- 전처리 ----------
    def on_preprocess_clicked(self):
        if not self.wb:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        # 미리보기에서 수정해둔 내용이 있으면 먼저 workbook에 반영
        if self.model:
            self.model.apply_dirty_to_sheet()

        company = self.control_panel.get_company_combo().currentText()
        keyword = self.control_panel.get_search_edit().text().strip()

        try:
            preprocess_inplace(self.wb, company=company, keyword=keyword)
        except AppError as e:
            QMessageBox.critical(self, "오류", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "오류", f"전처리 실패:\n{e}")
            return

        company = self.control_panel.get_company_combo().currentText()
        # self.info_panel.set_company(company if company != "선택" else "-")
        self.info_panel.set_remark("전처리 완료. 미리보기 갱신됨")
        self.refresh_preview_after_processing()

    def refresh_preview_after_processing(self):
        if not self.wb:
            return

        sheet_combo = self.preview_container.get_sheet_combo()
        current_sheet = sheet_combo.currentText()
        if not current_sheet or current_sheet not in self.wb.sheetnames:
            current_sheet = self.wb.sheetnames[0] if self.wb.sheetnames else ""

        sheet_combo.blockSignals(True)
        sheet_combo.clear()
        sheet_combo.addItems(self.wb.sheetnames)
        if current_sheet:
            sheet_combo.setCurrentText(current_sheet)
        sheet_combo.blockSignals(False)

        if current_sheet:
            self.load_sheet(current_sheet)

        self._refresh_editable_label()

    # ---------- export ----------
    def save_as_file(self):
        if not self.wb:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        if self.model:
            self.model.apply_dirty_to_sheet()

        save_path, _ = QFileDialog.getSaveFileName(self, "최종 엑셀로 저장", "", "Excel Files (*.xlsx)")
        if not save_path:
            return

        try:
            save_workbook_safe(self.wb, Path(save_path))
            QMessageBox.information(self, "완료", "저장했습니다.")
        except AppError as e:
            QMessageBox.critical(self, "오류", str(e))


