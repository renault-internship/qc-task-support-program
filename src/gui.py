from __future__ import annotations

import sys
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QTableView,
    QPushButton, QFileDialog, QMessageBox, QComboBox,
    QLineEdit, QLabel, QGroupBox, QFormLayout,
    QDialog, QSpinBox, QFrame, QApplication
)

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from src.constants import ALLOWED_EXT
from src.database import init_database, get_company_info, get_all_companies, upsert_company
from src.excel_processor import process_all

# app 모듈 (선택적 - 없어도 동작하도록)
try:
    from app.models.excel_table_model import ExcelSheetModel
    from app.services.excel_io import load_workbook_safe, save_workbook_safe
    from app.services.preprocess import preprocess_inplace
    from app.utils.errors import AppError
    HAS_APP_MODULE = True
except ImportError:
    ExcelSheetModel = None
    load_workbook_safe = None
    save_workbook_safe = None
    preprocess_inplace = None
    AppError = Exception
    HAS_APP_MODULE = False


class AddCompanyDialog(QDialog):
    """업체 추가 다이얼로그"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("업체 추가")
        self.setFixedSize(400, 250)
        
        layout = QFormLayout()
        
        # SAP 코드
        self.sap_code_edit = QLineEdit()
        self.sap_code_edit.setPlaceholderText("예: B907")
        layout.addRow("SAP 코드 *:", self.sap_code_edit)
        
        # SAP 기업명
        self.sap_name_edit = QLineEdit()
        self.sap_name_edit.setPlaceholderText("예: AMS")
        layout.addRow("SAP 기업명 *:", self.sap_name_edit)
        
        # 규칙 테이블명
        self.rule_table_edit = QLineEdit()
        self.rule_table_edit.setPlaceholderText("예: rule_B907")
        layout.addRow("규칙 테이블명 *:", self.rule_table_edit)
        
        # 보증 주행거리
        self.warranty_mileage_spin = QSpinBox()
        self.warranty_mileage_spin.setRange(0, 1000000)
        self.warranty_mileage_spin.setValue(50000)
        self.warranty_mileage_spin.setSuffix(" km")
        layout.addRow("보증 주행거리:", self.warranty_mileage_spin)
        
        # 보증 기간 (년)
        self.warranty_period_spin = QSpinBox()
        self.warranty_period_spin.setRange(0, 100)
        self.warranty_period_spin.setValue(3)
        self.warranty_period_spin.setSuffix(" 년")
        layout.addRow("보증 기간:", self.warranty_period_spin)
        
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
    
    def get_data(self):
        """입력된 데이터 반환"""
        return {
            "sap_code": self.sap_code_edit.text().strip(),
            "sap_name": self.sap_name_edit.text().strip(),
            "rule_table_name": self.rule_table_edit.text().strip(),
            "warranty_mileage": self.warranty_mileage_spin.value(),
            "warranty_period": self.warranty_period_spin.value() * 365  # 년을 일로 변환
        }


class MainWindow(QWidget):
    """
    - 좌측: 미리보기(QTableView) + 시트 선택 + 정보 패널
    - 우측: 업로드/전처리/기업선택/search/export
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("엑셀 전처리 도구")
        self.resize(1024, 576)  # 기본 윈도우 크기 (조절 가능)

        self.file_path: Path | None = None
        self.wb: Workbook | None = None
        self.model: ExcelSheetModel | None = None
        
        # 데이터베이스 초기화
        init_database()

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
        
        # 기업 선택
        company_label = QLabel("기업 선택")
        self.company_combo = QComboBox()
        self.load_companies()
        
        # 업체 추가 버튼
        company_row = QHBoxLayout()
        company_row.addWidget(self.company_combo, 1)
        self.btn_add_company = QPushButton("+")
        self.btn_add_company.setFixedWidth(30)
        self.btn_add_company.setToolTip("업체 추가")
        self.btn_add_company.clicked.connect(self.add_company)
        company_row.addWidget(self.btn_add_company)
        
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
        right_box.addWidget(company_label)
        right_box.addLayout(company_row)
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
        self.company_combo.currentTextChanged.connect(self._on_company_changed)

    def _set_info_defaults(self):
        self.lbl_company.setText("-")
        self.lbl_remark.setText("-")
        self.lbl_editable.setText("-")
    
    def _on_company_changed(self, name: str):
        """기업 선택 변경 시 정보 업데이트"""
        if name and name != "(기업 정보 없음)" and name != "선택":
            self.lbl_company.setText(name)
        else:
            self.lbl_company.setText("-")
    
    def load_companies(self):
        """기업 목록 로드 (DB에서)"""
        self.company_combo.clear()
        companies = get_all_companies()
        if companies:
            self.company_combo.addItems(companies)
        else:
            self.company_combo.addItem("(기업 정보 없음)")
    
    def add_company(self):
        """업체 추가 다이얼로그 열기"""
        dialog = AddCompanyDialog(self)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.get_data()
            
            # 필수 필드 검증
            if not data["sap_code"]:
                QMessageBox.warning(self, "오류", "SAP 코드를 입력해주세요.")
                return
            if not data["sap_name"]:
                QMessageBox.warning(self, "오류", "SAP 기업명을 입력해주세요.")
                return
            if not data["rule_table_name"]:
                QMessageBox.warning(self, "오류", "규칙 테이블명을 입력해주세요.")
                return
            
            try:
                upsert_company(
                    sap_code=data["sap_code"],
                    sap_name=data["sap_name"],
                    rule_table_name=data["rule_table_name"],
                    warranty_mileage=data["warranty_mileage"],
                    warranty_period=data["warranty_period"]
                )
                QMessageBox.information(self, "완료", "업체가 추가되었습니다.")
                self.load_companies()  # 목록 새로고침
            except Exception as e:
                QMessageBox.critical(self, "오류", f"업체 추가 실패: {str(e)}")

    # ---------- 업로드 ----------
    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 선택", "", "Excel Files (*.xlsx)")
        if not path:
            return

        self.file_path = Path(path)

        try:
            if HAS_APP_MODULE and load_workbook_safe:
                self.wb = load_workbook_safe(self.file_path)
            else:
                self.wb = load_workbook(self.file_path)
        except Exception as e:
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
        if company and company != "선택" and company != "(기업 정보 없음)":
            self.lbl_company.setText(company)
        else:
            self.lbl_company.setText("-")
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
        if self.model and hasattr(self.model, 'apply_dirty_to_sheet'):
            self.model.apply_dirty_to_sheet()
        self.load_sheet(sheet_name)

    # ---------- 전처리 ----------
    def on_preprocess_clicked(self):
        if not self.wb:
            QMessageBox.information(self, "안내", "먼저 파일을 업로드하세요.")
            return

        # 기업 정보 확인
        company_name = self.company_combo.currentText()
        if company_name == "(기업 정보 없음)" or company_name == "선택":
            QMessageBox.warning(self, "오류", "기업을 선택해주세요.")
            return

        company_info = get_company_info(company_name)
        if not company_info:
            QMessageBox.warning(self, "오류", f"기업정보를 찾을 수 없습니다: {company_name}")
            return

        # 미리보기에서 수정해둔 내용이 있으면 먼저 workbook에 반영
        if self.model and hasattr(self.model, 'apply_dirty_to_sheet'):
            self.model.apply_dirty_to_sheet()

        keyword = self.search_edit.text().strip()

        try:
            # 기존 preprocess_inplace 사용 (app 모듈이 있는 경우)
            if HAS_APP_MODULE and preprocess_inplace:
                preprocess_inplace(self.wb, company=company_name, keyword=keyword)
            else:
                # app 모듈이 없으면 src.excel_processor 사용
                # 임시 파일로 저장 후 처리
                import tempfile
                import os
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                    tmp_path = tmp.name
                self.wb.save(tmp_path)
                process_all(tmp_path, tmp_path, company_info)
                self.wb = load_workbook(tmp_path)
                os.unlink(tmp_path)
        except Exception as e:
            QMessageBox.critical(self, "오류", f"전처리 실패:\n{e}")
            return

        self.lbl_company.setText(company_name)
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

        if self.model and hasattr(self.model, 'apply_dirty_to_sheet'):
            self.model.apply_dirty_to_sheet()

        save_path, _ = QFileDialog.getSaveFileName(self, "최종 엑셀로 저장", "", "Excel Files (*.xlsx)")
        if not save_path:
            return

        try:
            if HAS_APP_MODULE and save_workbook_safe:
                save_workbook_safe(self.wb, Path(save_path))
            else:
                self.wb.save(save_path)
            QMessageBox.information(self, "완료", "저장했습니다.")
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))

    def export_rule_stub(self):
        QMessageBox.information(self, "안내", "rule export는 아직 연결 안 함.")


def main():
    """애플리케이션 진입점"""
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
