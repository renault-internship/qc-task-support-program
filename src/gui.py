# gui.py
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QTableView,
    QPushButton, QFileDialog, QMessageBox, QComboBox,
    QLineEdit, QLabel, QGroupBox, QFormLayout
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

        # ===== 좌측: 미리보기 =====
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)

        self.table = QTableView()
        self.table.setAlternatingRowColors(True)

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
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("search")

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
        right_box.addStretch(1)
        right_box.addWidget(self.btn_export_rule)
        right_box.addWidget(self.btn_export_final)

        right = QWidget()
        right.setFixedWidth(220)
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

        company = self.company_combo.currentText()
        self.lbl_company.setText(company if company != "선택" else "-")
        self.lbl_remark.setText("업로드 완료. 전처리 전 상태")
        self.lbl_editable.setText("현재: 전체 셀 편집 가능(추후 rule로 제한)")

    # ---------- 시트 로드/변경 ----------
    def load_sheet(self, sheet_name: str):
        if not self.wb:
            return
        ws = self.wb[sheet_name]
        self.model = ExcelSheetModel(ws, self)
        self.table.setModel(self.model)
        self.table.resizeColumnsToContents()

    def on_sheet_changed(self, sheet_name: str):
        if self.model:
            self.model.apply_dirty_to_sheet()
        self.load_sheet(sheet_name)

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

        self.lbl_company.setText(company if company != "선택" else "-")
        self.lbl_remark.setText("전처리 완료. 미리보기 갱신됨")
        self.lbl_editable.setText("전처리 후: 필요한 경우 rule 기반 편집 제한 적용 예정")

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
    def __init__(self, ws, parent=None):
        super().__init__(parent)
        self.ws = ws
        self.max_row = ws.max_row
        self.max_col = ws.max_column
        self.dirty = {}

    def rowCount(self, parent=QModelIndex()):
        return self.max_row

    def columnCount(self, parent=QModelIndex()):
        return self.max_col

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r = index.row() + 1
        c = index.column() + 1
        if role in (Qt.DisplayRole, Qt.EditRole):
            v = self.dirty.get((r, c), self.ws.cell(row=r, column=c).value)
            return "" if v is None else str(v)
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False
        r = index.row() + 1
        c = index.column() + 1
        text = (value or "").strip()

        new_val = None
        if text != "":
            raw = text.replace(",", "")
            try:
                if "." in raw:
                    new_val = float(raw)
                else:
                    new_val = int(raw)
            except ValueError:
                new_val = text

        self.dirty[(r, c)] = new_val
        self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
        return True

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self._excel_col_name(section + 1)
        return str(section + 1)

    @staticmethod
    def _excel_col_name(n: int) -> str:
        name = ""
        while n:
            n, rem = divmod(n - 1, 26)
            name = chr(65 + rem) + name
        return name

    def apply_dirty_to_sheet(self):
        for (r, c), v in self.dirty.items():
            self.ws.cell(row=r, column=c).value = v
