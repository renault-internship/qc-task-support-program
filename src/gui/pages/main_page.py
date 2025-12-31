from __future__ import annotations

from pathlib import Path
from typing import Dict, Any

from PySide6.QtCore import Qt, QSortFilterProxyModel, QRegularExpression
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QFileDialog, QMessageBox,
    QAbstractItemView, QMenu, QSplitter
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
from src.gui.excel_filter import ExcelFilterProxyModel, ColumnFilterDialog
from src.gui.dialogs import AddRuleDialog


class MainPageWidget(QWidget):
    """메인 페이지 - 엑셀 전처리 도구"""

    def __init__(self, parent=None):
        super().__init__(parent)

        self.file_path: Path | None = None
        self.wb: Workbook | None = None
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
        self.control_panel.get_upload_button().clicked.connect(self.open_file)
        self.control_panel.get_preprocess_button().clicked.connect(self.on_preprocess_clicked)
        self.control_panel.get_company_combo().currentTextChanged.connect(self._on_company_changed)
        self.control_panel.get_search_edit().textChanged.connect(self.on_search_changed)
        self.control_panel.get_edit_all_checkbox().stateChanged.connect(self.on_edit_mode_changed)

        self.preview_container.get_sheet_combo().currentTextChanged.connect(self.on_sheet_changed)
        self.control_panel.get_export_final_button().clicked.connect(self.save_as_file)

    # ================= 회사 =================
    def load_companies(self):
        combo = self.control_panel.get_company_combo()
        combo.clear()
        companies = get_all_companies()
        combo.addItem("선택")
        if companies:
            combo.addItems(companies)

    def _on_company_changed(self, name: str):
        if not name or name == "선택":
            self.info_panel.set_company_info("", "")
            self.info_panel.set_rules([])
            self.current_company_info = None
            return

        company_info = get_company_info(name)
        if not company_info:
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
                    self.control_panel.get_company_combo().currentText()
                )
            except Exception as e:
                QMessageBox.critical(self, "오류", str(e))

    # ================= 업로드 =================
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 선택", "", "Excel Files (*.xlsx)"
        )
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

        self.info_panel.set_remark("업로드 완료. 전처리 전 상태")

    # ================= 시트 =================
    def load_sheet(self, sheet_name: str):
        if not self.wb:
            return

        ws = self.wb[sheet_name]
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
        table.setSortingEnabled(True)
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
        if not self.wb:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        if self.model:
            self.model.apply_dirty_to_sheet()

        try:
            preprocess_inplace(
                self.wb,
                company=self.control_panel.get_company_combo().currentText(),
                keyword=self.control_panel.get_search_edit().text().strip()
            )
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))
            return

        self.info_panel.set_remark("전처리 완료. 미리보기 갱신됨")
        self.load_sheet(self.preview_container.get_sheet_combo().currentText())

    # ================= 저장 =================
    def save_as_file(self):
        if not self.wb:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        if self.model:
            self.model.apply_dirty_to_sheet()

        path, _ = QFileDialog.getSaveFileName(
            self, "최종 엑셀로 저장", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return

        try:
            save_workbook_safe(self.wb, Path(path))
            QMessageBox.information(self, "완료", "저장했습니다.")
        except AppError as e:
            QMessageBox.critical(self, "오류", str(e))

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
