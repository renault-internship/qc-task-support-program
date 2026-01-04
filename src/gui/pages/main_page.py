from __future__ import annotations

from pathlib import Path
from typing import Dict, Any

from PySide6.QtCore import Qt, QSortFilterProxyModel, QRegularExpression, QStringListModel
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFileDialog, QMessageBox,
    QAbstractItemView, QMenu, QSplitter
)

from openpyxl.workbook.workbook import Workbook

from src.utils import load_workbook_safe, save_workbook_safe, AppError
from src.excel_processor import preprocess_inplace
from src.database import (
    get_company_info, get_all_companies, get_all_companies_with_code,
    get_rules_from_table, add_rule_to_table
)
from src.gui.containers import (
    PreviewContainer, InfoPanel, ControlPanel
)
from src.gui.models import ExcelSheetModel
from src.gui.excel_filter import ExcelFilterProxyModel, ColumnFilterDialog
from src.gui.dialogs import AddRuleDialog


class MainPageWidget(QWidget):
    """메인 페이지 - 엑셀 전처리 도구"""

    def __init__(self, parent=None):
        super().__init__(parent)

        self.file_path_domestic: Path | None = None
        self.file_path_overseas: Path | None = None
        self.wb_domestic: Workbook | None = None
        self.wb_overseas: Workbook | None = None
        self.model: ExcelSheetModel | None = None
        self.proxy: QSortFilterProxyModel | None = None
        self.current_company_info: Dict[str, Any] | None = None

        # ================= 컨테이너 생성 =================
        self.control_panel = ControlPanel(self)
        self.preview_container = PreviewContainer(self)
        self.info_panel = InfoPanel(self)

        # ================= 레이아웃 구성 =================
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.control_panel)
        layout.addSpacing(4)

        # QSplitter로 preview와 info_panel 높이 조절 가능하게
        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(self.preview_container)
        splitter.addWidget(self.info_panel)

        # 초기 stretch 비율
        splitter.setStretchFactor(0, 4)  # preview_container가 크게
        splitter.setStretchFactor(1, 1)  # info_panel 작게

        # handle 시각적으로 표시
        splitter.setHandleWidth(6)
        splitter.setStyleSheet("""
        QSplitter::handle {
            background-color: #DDD;
            border-top: 2px solid #AAA;
            border-bottom: 2px solid #AAA;
        }
        QSplitter::handle:hover {
            background-color: #BBB;
        }
        """)

        layout.addWidget(splitter, 1)

        self.setLayout(layout)

        self._connect_signals()
        self._initialize()

    # ================= 초기화 =================
    def _initialize(self):
        self.info_panel.set_remark("-")
        self.load_companies()

    def _connect_signals(self):
        self.control_panel.get_upload_domestic_button().clicked.connect(lambda: self.open_file("domestic"))
        self.control_panel.get_upload_overseas_button().clicked.connect(lambda: self.open_file("overseas"))
        self.control_panel.get_preprocess_button().clicked.connect(self.on_preprocess_clicked)
        # 검색창에서 Enter 키 또는 편집 완료 시 회사 선택
        self.control_panel.get_company_edit().editingFinished.connect(self._on_company_search_finished)
        self.control_panel.get_company_edit().returnPressed.connect(self._on_company_search_finished)
        # 자동완성 목록에서 항목 선택 시
        self.control_panel.get_company_completer().activated.connect(self._on_company_selected_from_completer)
        self.control_panel.get_search_edit().textChanged.connect(self.on_search_changed)
        self.control_panel.get_edit_all_checkbox().stateChanged.connect(self.on_edit_mode_changed)

        self.control_panel.get_sheet_combo().currentTextChanged.connect(self.on_sheet_changed)
        self.control_panel.get_export_final_button().clicked.connect(self.save_as_file)

    # ================= 회사 =================
    def load_companies(self):
        """협력사 목록 로드 및 자동완성 설정 (코드와 이름 모두 포함)"""
        companies_data = get_all_companies_with_code()
        if companies_data:
            # 코드와 이름을 모두 포함한 리스트 생성
            # 형식: "이름 (코드)"만 사용
            company_list = []
            for company in companies_data:
                sap_code = company["sap_code"]
                sap_name = company["sap_name"]
                # "이름 (코드)" 형식만 추가
                company_list.append(f"{sap_name} ({sap_code})")
            
            # QCompleter에 모델 설정
            model = QStringListModel(company_list, self)
            completer = self.control_panel.get_company_completer()
            completer.setModel(model)

    def _on_company_search_finished(self):
        """검색창에서 Enter 키 또는 편집 완료 시 호출"""
        text = self.control_panel.get_company_edit().text().strip()
        # "이름 (코드)" 또는 "코드 - 이름" 형식에서 실제 이름 또는 코드 추출
        name_or_code = self._extract_company_name_or_code(text)
        self._on_company_changed(name_or_code)
    
    def _on_company_selected_from_completer(self, text: str):
        """자동완성 목록에서 항목 선택 시 호출"""
        # "이름 (코드)" 또는 "코드 - 이름" 형식에서 실제 이름 또는 코드 추출
        name_or_code = self._extract_company_name_or_code(text)
        self._on_company_changed(name_or_code)
    
    def _extract_company_name_or_code(self, text: str) -> str:
        """자동완성 텍스트에서 실제 회사 이름 또는 코드 추출"""
        text = text.strip()
        # "이름 (코드)" 형식인 경우
        if " (" in text and text.endswith(")"):
            # 이름 부분만 반환
            return text.split(" (")[0]
        # 그 외의 경우는 그대로 반환 (직접 입력한 경우: 이름 또는 코드)
        return text
    
    def _on_company_changed(self, name: str):
        """회사 선택 시 정보 로드"""
        if not name:
            self.info_panel.set_company_info("", "")
            self.info_panel.set_rules([])
            self.current_company_info = None
            return

        company_info = get_company_info(name)
        if not company_info:
            # 검색 결과가 없으면 경고만 표시하고 초기화하지 않음
            QMessageBox.warning(self, "오류", f"기업정보를 찾을 수 없습니다: {name}")
            return

        self.current_company_info = company_info

        # 회사 정보
        self.info_panel.set_company_info(
            company_info.get("sap_name", ""),
            company_info.get("sap_code", "")
        )

        # Remark
        self.info_panel.set_remark(company_info.get("remark", ""))

        # Rule → InfoPanel에 바로 표시
        rule_table_name = company_info.get("rule_table_name")
        if rule_table_name:
            rules = get_rules_from_table(rule_table_name)
            self.info_panel.set_rules(rules)
        else:
            self.info_panel.set_rules([])

    # ================= Rule 추가 =================
    def add_rule(self):
        if not self.current_company_info:
            QMessageBox.warning(self, "오류", "먼저 기업을 선택해주세요.")
            return

        rule_table_name = self.current_company_info.get("rule_table_name")
        if not rule_table_name:
            QMessageBox.warning(self, "오류", "선택한 기업에 Rule 테이블이 없습니다.")
            return

        dialog = AddRuleDialog(rule_table_name, self)
        if dialog.exec():
            try:
                add_rule_to_table(rule_table_name=rule_table_name, **dialog.get_data())
                QMessageBox.information(self, "완료", "규칙이 추가되었습니다.")
                self._on_company_changed(
                    self.control_panel.get_company_edit().text().strip()
                )
            except Exception as e:
                QMessageBox.critical(self, "오류", str(e))

    # ================= 업로드 =================
    def open_file(self, file_type: str = "domestic"):
        """
        파일 업로드
        Args:
            file_type: "domestic" (국내) 또는 "overseas" (해외)
        """
        title = "국내 청구서 선택" if file_type == "domestic" else "해외 청구서 선택"
        path, _ = QFileDialog.getOpenFileName(
            self, title, "", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        file_path = Path(path)
        try:
            wb = load_workbook_safe(file_path)
        except AppError as e:
            QMessageBox.critical(self, "오류", str(e))
            return

        # 워크북 저장
        if file_type == "domestic":
            self.file_path_domestic = file_path
            self.wb_domestic = wb
        else:
            self.file_path_overseas = file_path
            self.wb_overseas = wb

        # 시트 목록 업데이트
        self._update_sheet_list()

        # 첫 번째 시트 로드
        sheet_combo = self.control_panel.get_sheet_combo()
        if sheet_combo.count() > 0:
            sheet_combo.setCurrentIndex(0)
            self._load_sheet_from_combo()

        remark = f"{'국내' if file_type == 'domestic' else '해외'} 청구서 업로드 완료. 전처리 전 상태"
        self.info_panel.set_remark(remark)
    
    def _update_sheet_list(self):
        """시트 목록 업데이트 (국내/해외 모두 포함)"""
        sheet_combo = self.control_panel.get_sheet_combo()
        sheet_combo.blockSignals(True)
        sheet_combo.clear()
        
        # 국내 청구서 시트 추가
        if self.wb_domestic:
            for sheet_name in self.wb_domestic.sheetnames:
                sheet_combo.addItem(f"국내: {sheet_name}")
        
        # 해외 청구서 시트 추가
        if self.wb_overseas:
            for sheet_name in self.wb_overseas.sheetnames:
                sheet_combo.addItem(f"해외: {sheet_name}")
        
        sheet_combo.blockSignals(False)
    
    def _load_sheet_from_combo(self):
        """시트 콤보박스에서 선택한 시트 로드"""
        sheet_combo = self.control_panel.get_sheet_combo()
        sheet_text = sheet_combo.currentText()
        if not sheet_text:
            return
        self.load_sheet(sheet_text)

    # ================= 시트 =================
    def load_sheet(self, sheet_display_name: str):
        """
        시트 로드
        Args:
            sheet_display_name: "국내: Sheet1" 또는 "해외: Sheet1" 형식
        """
        # "국내: Sheet1" 또는 "해외: Sheet1" 형식에서 파싱
        if sheet_display_name.startswith("국내: "):
            if not self.wb_domestic:
                return
            actual_sheet_name = sheet_display_name.replace("국내: ", "")
            wb = self.wb_domestic
        elif sheet_display_name.startswith("해외: "):
            if not self.wb_overseas:
                return
            actual_sheet_name = sheet_display_name.replace("해외: ", "")
            wb = self.wb_overseas
        else:
            # 기존 형식 호환성 (없을 수도 있음)
            return

        if actual_sheet_name not in wb.sheetnames:
            return

        ws = wb[actual_sheet_name]
        self.model = ExcelSheetModel(ws, parent=self)

        edit_all = self.control_panel.get_edit_all_checkbox().isChecked()
        self.model.set_edit_all(edit_all)

        self.proxy = ExcelFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy.setFilterKeyColumn(-1)

        table = self.preview_container.get_table()
        table.clearSpans()
        table.setModel(self.proxy)

        table.setAlternatingRowColors(True)
        table.setSortingEnabled(False)  # 컬럼 헤더 클릭 정렬 비활성화
        table.setSelectionBehavior(QAbstractItemView.SelectItems)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        header = table.horizontalHeader()
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self._on_header_context_menu)

        table.resizeColumnsToContents()
        self._apply_excel_layout(ws)

        self.on_search_changed(self.control_panel.get_search_edit().text())

    def on_sheet_changed(self, sheet_name: str):
        if self.model:
            self.model.apply_dirty_to_sheet()
        self.load_sheet(sheet_name)

    # ================= 검색 =================
    def on_search_changed(self, text: str):
        if not self.proxy:
            return
        if not text:
            self.proxy.setFilterRegularExpression(QRegularExpression(""))
            return
        rx = QRegularExpression(
            QRegularExpression.escape(text),
            QRegularExpression.CaseInsensitiveOption
        )
        self.proxy.setFilterRegularExpression(rx)

    # ================= 편집 모드 =================
    def on_edit_mode_changed(self):
        if self.model:
            edit_all = self.control_panel.get_edit_all_checkbox().isChecked()
            self.model.set_edit_all(edit_all)
            self.model.layoutChanged.emit()

    # ================= 전처리 =================
    def on_preprocess_clicked(self):
        # 현재 선택된 시트가 있는지 확인
        sheet_combo = self.control_panel.get_sheet_combo()
        current_sheet = sheet_combo.currentText()
        if not current_sheet:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        # 현재 시트의 워크북 찾기
        if current_sheet.startswith("국내: "):
            if not self.wb_domestic:
                QMessageBox.information(self, "안내", "국내 청구서가 없습니다.")
                return
            wb = self.wb_domestic
        elif current_sheet.startswith("해외: "):
            if not self.wb_overseas:
                QMessageBox.information(self, "안내", "해외 청구서가 없습니다.")
                return
            wb = self.wb_overseas
        else:
            QMessageBox.information(self, "안내", "시트를 선택하세요.")
            return

        if self.model:
            self.model.apply_dirty_to_sheet()

        try:
            preprocess_inplace(
                wb,
                company=self.control_panel.get_company_edit().text().strip(),
                keyword=self.control_panel.get_search_edit().text().strip()
            )
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))
            return

        self.info_panel.set_remark("전처리 완료. 미리보기 갱신됨")
        self.load_sheet(current_sheet)

    # ================= 저장 =================
    def save_as_file(self):
        # 현재 선택된 시트의 dirty 데이터 저장
        if self.model:
            self.model.apply_dirty_to_sheet()

        # 저장할 워크북 결정
        if self.wb_domestic and self.wb_overseas:
            # 둘 다 있으면 합쳐서 저장
            from openpyxl import Workbook
            merged_wb = Workbook()
            merged_wb.remove(merged_wb.active)  # 기본 시트 제거
            
            # 국내 시트들 복사
            for sheet_name in self.wb_domestic.sheetnames:
                source_sheet = self.wb_domestic[sheet_name]
                new_sheet = merged_wb.create_sheet(f"국내_{sheet_name}")
                self._copy_sheet(source_sheet, new_sheet)
            
            # 해외 시트들 복사
            for sheet_name in self.wb_overseas.sheetnames:
                source_sheet = self.wb_overseas[sheet_name]
                new_sheet = merged_wb.create_sheet(f"해외_{sheet_name}")
                self._copy_sheet(source_sheet, new_sheet)
            
            wb_to_save = merged_wb
        elif self.wb_domestic:
            wb_to_save = self.wb_domestic
        elif self.wb_overseas:
            wb_to_save = self.wb_overseas
        else:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "최종 엑셀로 저장", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        try:
            save_workbook_safe(wb_to_save, Path(path))
            QMessageBox.information(self, "완료", "저장했습니다.")
        except AppError as e:
            QMessageBox.critical(self, "오류", str(e))
    
    def _copy_sheet(self, source_sheet, target_sheet):
        """시트 내용 복사"""
        for row in source_sheet.iter_rows():
            for cell in row:
                target_cell = target_sheet.cell(row=cell.row, column=cell.column)
                target_cell.value = cell.value
                if cell.has_style:
                    target_cell.font = cell.font
                    target_cell.border = cell.border
                    target_cell.fill = cell.fill
                    target_cell.number_format = cell.number_format
                    target_cell.protection = cell.protection
                    target_cell.alignment = cell.alignment
        
        # 병합 셀 복사
        for merged_range in source_sheet.merged_cells.ranges:
            target_sheet.merge_cells(str(merged_range))
        
        # 열 너비 복사
        for col in source_sheet.column_dimensions:
            target_sheet.column_dimensions[col].width = source_sheet.column_dimensions[col].width
        
        # 행 높이 복사
        for row in source_sheet.row_dimensions:
            target_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height

    # ================= 테이블 헤더 메뉴 =================
    def _on_header_context_menu(self, pos):
        if not self.proxy or not self.model:
            return

        table = self.preview_container.get_table()
        header = table.horizontalHeader()
        col = header.logicalIndexAt(pos)
        if col < 0:
            return

        menu = QMenu(self)
        act_filter = menu.addAction("필터...")
        act_clear = menu.addAction("이 컬럼 필터 해제")
        act_clear_all = menu.addAction("전체 필터 초기화")

        picked = menu.exec(header.mapToGlobal(pos))
        if not picked:
            return

        if picked == act_filter:
            col_name = ExcelSheetModel.excel_col_name(col + 1)
            ColumnFilterDialog(self.model, self.proxy, col, col_name, self).exec()
        elif picked == act_clear:
            self.proxy.clear_column_filter(col)
        elif picked == act_clear_all:
            self.proxy.clear_all_column_filters()

    # ================= 엑셀 레이아웃 =================
    def _apply_excel_layout(self, ws):
        table = self.preview_container.get_table()

        for col_idx in range(1, ws.max_column + 1):
            dim = ws.column_dimensions.get(
                ExcelSheetModel.excel_col_name(col_idx)
            )
            if dim and dim.width:
                table.setColumnWidth(col_idx - 1, int(dim.width * 7 + 12))

        for row_idx in range(1, ws.max_row + 1):
            dim = ws.row_dimensions.get(row_idx)
            if dim and dim.height:
                table.setRowHeight(row_idx - 1, int(dim.height * 1.33))
        
        # 병합 셀 처리: setSpan으로 병합 표시
        for mr in ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = mr.bounds
            # QTableWidget의 인덱스는 0부터 시작
            row = min_row - 1
            col = min_col - 1
            row_span = max_row - min_row + 1
            col_span = max_col - min_col + 1
            table.setSpan(row, col, row_span, col_span)
