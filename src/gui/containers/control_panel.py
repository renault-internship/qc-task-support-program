"""
상단 컨트롤 패널 - 업로드, 전처리, 기업 선택, 검색, 편집 모드
"""
from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QComboBox, QLineEdit, QLabel, QCheckBox, QCompleter, QSizePolicy
)
from PySide6.QtCore import Qt


class ControlPanel(QWidget):
    """상단 컨트롤 패널"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(8, 0, 8, 0)
        main_layout.setSpacing(4)
        
        # ================= 윗줄 =================
        top_row = QHBoxLayout()
        top_row.setSpacing(8)
        
        # 국내 청구서 불러오기
        self.btn_upload_domestic = QPushButton("국내 청구서 불러오기")
        top_row.addWidget(self.btn_upload_domestic)
        
        # 해외 청구서 불러오기
        self.btn_upload_overseas = QPushButton("해외 청구서 불러오기")
        top_row.addWidget(self.btn_upload_overseas)
        
        # 시트 선택
        top_row.addWidget(QLabel("시트:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        top_row.addWidget(self.sheet_combo, 1)  # stretch factor 1
        
        # 코멕스 불러오기
        top_row.addWidget(QLabel("코멕스 불러오기:"))
        self.company_edit = QLineEdit()
        self.company_edit.setPlaceholderText("협력사 검색 (이름 또는 코드)")
        # QCompleter 설정
        self.company_completer = QCompleter()
        self.company_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.company_completer.setFilterMode(Qt.MatchContains)
        self.company_edit.setCompleter(self.company_completer)
        self.company_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        top_row.addWidget(self.company_edit, 1)  # stretch factor 1 (시트와 동일)
        
        # 전처리
        self.btn_preprocess = QPushButton("전처리")
        top_row.addWidget(self.btn_preprocess)
        
        # export
        self.btn_export_final = QPushButton("EXCEL로 내려받기")
        top_row.addWidget(self.btn_export_final)
        
        top_row.addStretch()
        main_layout.addLayout(top_row)
        
        # ================= 아래줄 =================
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(8)
        
        # 검색
        bottom_row.addWidget(QLabel("검색:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("search (시트 내 검색)")
        bottom_row.addWidget(self.search_edit, 1)
        
        # 편집 제어
        self.chk_edit_all = QCheckBox("전체 셀 편집 허용")
        self.chk_edit_all.setChecked(False)
        bottom_row.addWidget(self.chk_edit_all)
        
        bottom_row.addStretch()
        main_layout.addLayout(bottom_row)
        
        self.setLayout(main_layout)
    
    def get_upload_domestic_button(self) -> QPushButton:
        """국내 청구서 불러오기 버튼 반환"""
        return self.btn_upload_domestic
    
    def get_upload_overseas_button(self) -> QPushButton:
        """해외 청구서 불러오기 버튼 반환"""
        return self.btn_upload_overseas
    
    def get_preprocess_button(self) -> QPushButton:
        """전처리 버튼 반환"""
        return self.btn_preprocess
    
    def get_company_edit(self) -> QLineEdit:
        """기업 선택 검색창 반환"""
        return self.company_edit
    
    def get_company_completer(self) -> QCompleter:
        """기업 선택 자동완성 반환"""
        return self.company_completer
    
    def get_sheet_combo(self) -> QComboBox:
        """시트 선택 콤보박스 반환"""
        return self.sheet_combo
    
    def get_search_edit(self) -> QLineEdit:
        """검색 입력창 반환"""
        return self.search_edit
    
    def get_edit_all_checkbox(self) -> QCheckBox:
        """편집 모드 체크박스 반환"""
        return self.chk_edit_all
    
    def get_export_final_button(self) -> QPushButton:
        """export (최종 엑셀) 버튼 반환"""
        return self.btn_export_final

