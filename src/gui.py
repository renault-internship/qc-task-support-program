"""
GUI 컴포넌트 모듈
"""
import sys
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QFrame, QComboBox,
    QDialog, QLineEdit, QFormLayout, QSpinBox
)

from src.constants import ALLOWED_EXT
from src.database import init_database, get_company_info, get_all_companies, upsert_company
from src.excel_processor import process_all


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


class DropZone(QFrame):
    """드래그 앤 드롭 영역"""
    def __init__(self, on_file_dropped):
        super().__init__()
        self.on_file_dropped = on_file_dropped
        self.setAcceptDrops(True)
        self.setFixedHeight(80)
        self.setStyleSheet("QFrame { border: 1px dashed #888; border-radius: 6px; }")
        lay = QVBoxLayout()
        self.lbl = QLabel("엑셀 드래그 (.xlsx)")
        self.lbl.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.lbl)
        self.setLayout(lay)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            p = Path(event.mimeData().urls()[0].toLocalFile())
            if p.suffix.lower() in ALLOWED_EXT:
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            self.on_file_dropped(Path(urls[0].toLocalFile()))


class App(QWidget):
    """메인 애플리케이션 윈도우"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AMS")
        self.setFixedSize(420, 280)

        self.in_path: Path | None = None

        # 데이터베이스 초기화
        init_database()

        root = QVBoxLayout()

        # 기업 선택 콤보박스
        company_row = QHBoxLayout()
        company_row.addWidget(QLabel("기업:"))
        self.company_combo = QComboBox()
        self.load_companies()
        company_row.addWidget(self.company_combo, 1)
        
        # 업체 추가 버튼
        self.btn_add_company = QPushButton("+")
        self.btn_add_company.setFixedWidth(30)
        self.btn_add_company.setToolTip("업체 추가")
        self.btn_add_company.clicked.connect(self.add_company)
        company_row.addWidget(self.btn_add_company)
        
        root.addLayout(company_row)

        # 드래그 앤 드롭 영역
        self.drop = DropZone(self.set_file)
        root.addWidget(self.drop)

        # 파일 선택
        row = QHBoxLayout()
        self.lbl = QLabel("파일: (없음)")
        self.lbl.setWordWrap(True)
        btn = QPushButton("선택")
        btn.clicked.connect(self.pick_file)
        row.addWidget(self.lbl, 1)
        row.addWidget(btn)
        root.addLayout(row)

        # 저장 버튼
        self.btn_export = QPushButton("저장")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self.export_processed)
        root.addWidget(self.btn_export)

        self.setLayout(root)

    def load_companies(self):
        """기업 목록 로드"""
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

    def set_file(self, p: Path):
        """파일 설정"""
        if p.suffix.lower() not in ALLOWED_EXT:
            QMessageBox.warning(self, "확장자", "현재는 .xlsx만 지원합니다.")
            return
        self.in_path = p
        self.lbl.setText(f"파일: {p.name}")
        self.btn_export.setEnabled(True)

    def pick_file(self):
        """파일 선택 다이얼로그"""
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 선택", "", "Excel Files (*.xlsx)")
        if path:
            self.set_file(Path(path))

    def export_processed(self):
        """처리 후 저장"""
        if not self.in_path:
            return

        # 기업 정보 가져오기
        company_name = self.company_combo.currentText()
        if company_name == "(기업 정보 없음)":
            QMessageBox.warning(self, "오류", "기업을 선택해주세요.")
            return

        company_info = get_company_info(company_name)
        if not company_info:
            QMessageBox.warning(self, "오류", f"기업정보를 찾을 수 없습니다: {company_name}")
            return

        # 저장 경로 선택
        ts = datetime.now().strftime("%H%M%S")
        default = f"{self.in_path.stem}_{ts}{self.in_path.suffix}"

        save_path, _ = QFileDialog.getSaveFileName(self, "저장", default, "Excel Files (*.xlsx)")
        if not save_path:
            return

        try:
            process_all(str(self.in_path), save_path, company_info)
            QMessageBox.information(self, "완료", "처리 후 저장됨")
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))


def main():
    """애플리케이션 진입점"""
    app = QApplication(sys.argv)
    w = App()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

