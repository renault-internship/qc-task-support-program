# gui.py
from __future__ import annotations

from pathlib import Path
from datetime import datetime, date

from PySide6.QtCore import (
    Qt,
    QAbstractTableModel,
    QModelIndex,
    QSortFilterProxyModel,
    QRegularExpression,
)
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QTableView,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QComboBox,
    QLineEdit,
    QLabel,
    QGroupBox,
    QFormLayout,
    QCheckBox,
)

from openpyxl.workbook.workbook import Workbook

from src.utils import load_workbook_safe, save_workbook_safe, AppError
from src.excel_processor import preprocess_inplace


class MainWindow(QWidget):
    """
    - 좌측: 미리보기(QTableView) + 시트 선택 + 정보 패널
    - 우측: 업로드/전처리/기업선택/search/export
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("엑셀 전처리 도구")

        self.file_path: Path | None = None
        self.wb: Workbook | None = None

        self.model: ExcelSheetModel | None = None
        self.proxy: QSortFilterProxyModel | None = None

        # ===== 좌측: 미리보기 =====
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)

        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(False)

        left_top = QHBoxLayout()
        left_top.addWidget(QLabel("시트"))
        left_top.addWidget(self.sheet_combo, 1)

        left_preview_box = QVBoxLayout()
        left_preview_box.addLayout(left_top)
        left_preview_box.addWidget(self.table, 1)

        info_group = QGroupBox("정보")
        form = QFormLayout()

        self.lbl_company = QLabel("-")
        self.lbl_remark = QLabel("-")
        self.lbl_editable = QLabel("-")

        for lbl in (self.lbl_company, self.lbl_remark, self.lbl_editable):
            lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        form.addRow("기업명", self.lbl_company)
        form.addRow("비고(remark)", self.lbl_remark)
        form.addRow("변경가능(rule)", self.lbl_editable)
        info_group.setLayout(form)

        left_preview_box.addWidget(info_group)

        left = QWidget()
        left.setLayout(left_preview_box)

        # ===== 우측: 컨트롤 =====
        self.btn_upload = QPushButton("업로드")
        self.btn_preprocess = QPushButton("전처리")

        self.company_combo = QComboBox()
        self.company_combo.addItems(["선택", "AMS", "건화", "etc"])  # 더미
        self.company_combo.currentTextChanged.connect(self._refresh_company_label)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("search (전체 검색)")
        self.search_edit.textChanged.connect(self.on_search_changed)

        # 편집 제어: 전체 편집 vs 구상율만
        self.chk_edit_all = QCheckBox("전체 셀 편집 허용")
        self.chk_edit_all.setChecked(False)
        self.chk_edit_all.stateChanged.connect(self.on_edit_mode_changed)

        self.btn_export_rule = QPushButton("export (rule)")
        self.btn_export_final = QPushButton("export (최종 엑셀)")

        self.btn_upload.clicked.connect(self.open_file)
        self.btn_preprocess.clicked.connect(self.on_preprocess_clicked)
        self.btn_export_rule.clicked.connect(self.export_rule_stub)
        self.btn_export_final.clicked.connect(self.save_as_file)

        right_box = QVBoxLayout()
        right_box.addWidget(self.btn_upload)
        right_box.addWidget(self.btn_preprocess)
        right_box.addSpacing(8)
        right_box.addWidget(QLabel("기업 선택"))
        right_box.addWidget(self.company_combo)
        right_box.addSpacing(8)
        right_box.addWidget(self.search_edit)
        right_box.addSpacing(8)
        right_box.addWidget(self.chk_edit_all)
        right_box.addStretch(1)
        right_box.addWidget(self.btn_export_rule)
        right_box.addWidget(self.btn_export_final)

        right = QWidget()
        right.setFixedWidth(240)
        right.setLayout(right_box)

        root = QHBoxLayout()
        root.addWidget(left, 1)
        root.addWidget(right)
        self.setLayout(root)

        self._set_info_defaults()

    def _set_info_defaults(self):
        self.lbl_company.setText("-")
        self.lbl_remark.setText("-")
        self.lbl_editable.setText("-")

    def _refresh_company_label(self):
        company = self.company_combo.currentText()
        self.lbl_company.setText(company if company != "선택" else "-")

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

        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(self.wb.sheetnames)
        self.sheet_combo.blockSignals(False)

        if self.wb.sheetnames:
            self.sheet_combo.setCurrentIndex(0)
            self.load_sheet(self.wb.sheetnames[0])

        self._refresh_company_label()
        self.lbl_remark.setText("업로드 완료. 전처리 전 상태")
        self._refresh_editable_label()

    # ---------- 시트 로드/변경 ----------
    def load_sheet(self, sheet_name: str):
        if not self.wb:
            return

        ws = self.wb[sheet_name]
        self.model = ExcelSheetModel(ws, parent=self)

        # 편집 모드 반영
        self.model.set_edit_all(self.chk_edit_all.isChecked())

        self.proxy = QSortFilterProxyModel(self)
        self.proxy.setSourceModel(self.model)
        self.proxy.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.proxy.setFilterKeyColumn(-1)  # 전체 컬럼 대상으로 검색

        self.table.setModel(self.proxy)
        self.table.resizeColumnsToContents()

        # 기존 검색어 유지
        self.on_search_changed(self.search_edit.text())

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
            self.model.set_edit_all(self.chk_edit_all.isChecked())
            # flags만 바뀌는 거라 화면 갱신
            self.model.layoutChanged.emit()
        self._refresh_editable_label()

    def _refresh_editable_label(self):
        if self.chk_edit_all.isChecked():
            self.lbl_editable.setText("현재: 전체 셀 편집 가능")
        else:
            # 구상율 컬럼 찾았는지 표시
            if self.model and self.model.editable_cols:
                cols = ", ".join(self.model.excel_col_name(c) for c in sorted(self.model.editable_cols))
                self.lbl_editable.setText(f"현재: 구상율 컬럼만 편집 가능 ({cols})")
            else:
                self.lbl_editable.setText("현재: 편집 제한(구상율 컬럼 미탐지 또는 없음)")

    # ---------- 전처리 ----------
    def on_preprocess_clicked(self):
        if not self.wb:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        # 미리보기에서 수정해둔 내용이 있으면 먼저 workbook에 반영
        if self.model:
            self.model.apply_dirty_to_sheet()

        company = self.company_combo.currentText()
        keyword = self.search_edit.text().strip()

        try:
            preprocess_inplace(self.wb, company=company, keyword=keyword)
        except AppError as e:
            QMessageBox.critical(self, "오류", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "오류", f"전처리 실패:\n{e}")
            return

        self._refresh_company_label()
        self.lbl_remark.setText("전처리 완료. 미리보기 갱신됨")
        self.refresh_preview_after_processing()

    def refresh_preview_after_processing(self):
        if not self.wb:
            return

        current_sheet = self.sheet_combo.currentText()
        if not current_sheet or current_sheet not in self.wb.sheetnames:
            current_sheet = self.wb.sheetnames[0] if self.wb.sheetnames else ""

        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(self.wb.sheetnames)
        if current_sheet:
            self.sheet_combo.setCurrentText(current_sheet)
        self.sheet_combo.blockSignals(False)

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

    def export_rule_stub(self):
        QMessageBox.information(self, "안내", "rule export는 아직 연결 안 함.")


class ExcelSheetModel(QAbstractTableModel):
    """
    - openpyxl worksheet를 UI로 보여주는 모델
    - dirty: UI에서 수정된 값(메모리상)
    - 표시 포맷: 숫자 콤마, 날짜 yyyy-mm-dd
    - 편집 제어:
        * edit_all=True  -> 헤더(1행) 제외 전부 편집
        * edit_all=False -> '구상율' 헤더 컬럼만 편집
    """
    def __init__(self, ws, parent=None):
        super().__init__(parent)
        self.ws = ws
        self.max_row = ws.max_row
        self.max_col = ws.max_column

        self.dirty: dict[tuple[int, int], object] = {}
        self.edit_all: bool = False
        self.editable_cols: set[int] = self._find_chargeback_rate_cols()

    # ----- Qt 필수 -----
    def rowCount(self, parent=QModelIndex()):
        return self.max_row

    def columnCount(self, parent=QModelIndex()):
        return self.max_col

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        r = index.row() + 1
        c = index.column() + 1

        v = self.dirty.get((r, c), self.ws.cell(row=r, column=c).value)

        if role == Qt.EditRole:
            # 편집기는 타입 보존(가능한 한)
            return "" if v is None else v

        if role == Qt.DisplayRole:
            return self._format_value(v)

        if role == Qt.BackgroundRole:
            # 수정된 셀은 표시(엑셀 느낌)
            if (r, c) in self.dirty:
                return QBrush(QColor(255, 250, 205))  # 연노랑
            return None

        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags

        r = index.row() + 1
        c = index.column() + 1
        base = Qt.ItemIsSelectable | Qt.ItemIsEnabled

        # 보통 헤더 행(1행)은 편집 막는 게 안전
        if r == 1:
            return base

        if self.edit_all:
            return base | Qt.ItemIsEditable

        # 구상율 컬럼만 편집
        if c in self.editable_cols:
            return base | Qt.ItemIsEditable

        return base

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False

        r = index.row() + 1
        c = index.column() + 1

        # 편집 불가 셀 방어
        if r == 1:
            return False
        if (not self.edit_all) and (c not in self.editable_cols):
            return False

        new_val = self._parse_user_input(value)

        self.dirty[(r, c)] = new_val
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole, Qt.BackgroundRole])
        return True

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.excel_col_name(section + 1)
        return str(section + 1)

    # ----- 유틸 -----
    def set_edit_all(self, on: bool):
        self.edit_all = bool(on)

    @staticmethod
    def excel_col_name(n: int) -> str:
        name = ""
        while n:
            n, rem = divmod(n - 1, 26)
            name = chr(65 + rem) + name
        return name

    @staticmethod
    def _format_value(v):
        if v is None:
            return ""
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, int):
            return f"{v:,}"
        if isinstance(v, float):
            s = f"{v:,.2f}"
            return s.rstrip("0").rstrip(".")
        if isinstance(v, (datetime, date)):
            return v.strftime("%Y-%m-%d")
        return str(v)

    @staticmethod
    def _parse_user_input(value):
        text = "" if value is None else str(value).strip()
        if text == "":
            return None

        raw = text.replace(",", "")

        # 숫자 파싱 (int/float)
        try:
            if "." in raw:
                return float(raw)
            return int(raw)
        except ValueError:
            return text

    def _find_chargeback_rate_cols(self) -> set[int]:
        """
        1행(헤더)에서 '구상'+'율' 포함 컬럼을 찾아 편집 가능 컬럼으로 등록
        (파일마다 컬럼 위치가 다를 수 있어 탐지 기반으로 처리)
        """
        editable = set()
        header_row = 1

        for c in range(1, self.max_col + 1):
            hv = self.ws.cell(row=header_row, column=c).value
            if hv and isinstance(hv, str):
                s = hv.replace(" ", "")
                if ("구상" in s and "율" in s) or ("chargeback" in hv.lower() and "rate" in hv.lower()):
                    editable.add(c)
        return editable

    def apply_dirty_to_sheet(self):
        for (r, c), v in self.dirty.items():
            self.ws.cell(row=r, column=c).value = v
        # 반영 후 dirty를 유지할지(되돌리기 위해) 정책 선택 가능
        # 지금은 "화면 표시(수정 강조)"를 위해 유지
