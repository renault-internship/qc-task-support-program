"""
GUI 컴포넌트 모듈
"""
import sys
from pathlib import Path
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QFrame, QComboBox
)

from src.constants import ALLOWED_EXT
from src.database import init_database, get_company_info, get_all_companies
from src.excel_processor import process_all


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
        self.setFixedSize(420, 250)

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

