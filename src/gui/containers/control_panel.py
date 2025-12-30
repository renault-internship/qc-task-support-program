"""
상단 컨트롤 패널 - 업로드, 전처리, 기업 선택, 검색, 편집 모드
"""
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout,
    QPushButton, QComboBox, QLineEdit, QLabel, QCheckBox
)


class ControlPanel(QWidget):
    """상단 컨트롤 패널"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # 업로드, 전처리, Export 버튼
        self.btn_upload = QPushButton("업로드")
        self.btn_preprocess = QPushButton("전처리")
        self.btn_export_final = QPushButton("export (최종 엑셀)")
        layout.addWidget(self.btn_upload)
        layout.addWidget(self.btn_preprocess)
        layout.addWidget(self.btn_export_final)
        layout.addSpacing(16)
        
        # 기업 선택
        layout.addWidget(QLabel("기업:"))
        self.company_combo = QComboBox()
        layout.addWidget(self.company_combo)
        layout.addSpacing(16)
        
        # 검색
        layout.addWidget(QLabel("검색:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("search (전체 검색)")
        layout.addWidget(self.search_edit, 1)
        layout.addSpacing(16)
        
        # 편집 제어
        self.chk_edit_all = QCheckBox("전체 셀 편집 허용")
        self.chk_edit_all.setChecked(False)
        layout.addWidget(self.chk_edit_all)
        layout.addStretch()
        
        self.setLayout(layout)
    
    def get_upload_button(self) -> QPushButton:
        """업로드 버튼 반환"""
        return self.btn_upload
    
    def get_preprocess_button(self) -> QPushButton:
        """전처리 버튼 반환"""
        return self.btn_preprocess
    
    def get_company_combo(self) -> QComboBox:
        """기업 선택 콤보박스 반환"""
        return self.company_combo
    
    def get_search_edit(self) -> QLineEdit:
        """검색 입력창 반환"""
        return self.search_edit
    
    def get_edit_all_checkbox(self) -> QCheckBox:
        """편집 모드 체크박스 반환"""
        return self.chk_edit_all
    
    def get_export_final_button(self) -> QPushButton:
        """export (최종 엑셀) 버튼 반환"""
        return self.btn_export_final

