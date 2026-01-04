"""
메인 윈도우 - 페이지네이션 컨테이너
"""
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QStackedWidget
)

from src.database import init_database
from src.gui.pages.main_page import MainPageWidget


class MainWindow(QWidget):
    """
    메인 윈도우 - 페이지네이션 컨테이너
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("엑셀 전처리 도구")
        
        # 데이터베이스 초기화
        init_database()
        
        # 페이지네이션 (탭 스타일)
        page_nav = QHBoxLayout()
        page_nav.setContentsMargins(0, 0, 0, 0)
        page_nav.setSpacing(0)
        
        self.btn_main_page = QPushButton("메인")
        self.btn_comex_page = QPushButton("comex 관리")
        self.btn_main_page.setCheckable(True)
        self.btn_comex_page.setCheckable(True)
        self.btn_main_page.setChecked(True)
        
        # 탭 스타일 적용
        self._apply_tab_style()
        
        page_nav.addWidget(self.btn_main_page)
        page_nav.addWidget(self.btn_comex_page)
        page_nav.addStretch()
        
        # 페이지 스택
        self.stacked = QStackedWidget()
        
        # 페이지 1: 메인 페이지
        self.main_page = MainPageWidget()
        self.stacked.addWidget(self.main_page)
        
        # 페이지 2: comex 관리 페이지
        from src.gui.pages.comex_management_page import ComExManagementPageWidget
        self.comex_page = ComExManagementPageWidget()
        self.stacked.addWidget(self.comex_page)
        
        # 레이아웃
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addLayout(page_nav)
        layout.addWidget(self.stacked)
        self.setLayout(layout)
        
        # 페이지 전환 연결
        self.btn_main_page.clicked.connect(lambda: self.switch_page(0))
        self.btn_comex_page.clicked.connect(lambda: self.switch_page(1))
    
    def _apply_tab_style(self):
        """탭 스타일 적용 (Qt Material Theme 스타일)"""
        # light_blue 테마의 primary color 사용 (#1976D2 또는 #2196F3)
        tab_style = """
        QPushButton {
            background-color: #F5F5F5;
            border: none;
            border-radius: 0px;
            padding: 6px 20px;
            text-align: center;
            font-family: 맑은 고딕;
            font-size: 10pt;
            color: #000000;
            min-width: 80px;
        }
        QPushButton:hover {
            background-color: #E0E0E0;
        }
        QPushButton:checked {
            background-color: #E0E0E0;
            color: #1976D2;
            border-right: 3px solid #1976D2;
            border-radius: 0px;
            font-family: 맑은 고딕;
            font-size: 10pt;
        }
        QPushButton:checked:hover {
            background-color: #D0D0D0;
        }
        """
        self.btn_main_page.setStyleSheet(tab_style)
        self.btn_comex_page.setStyleSheet(tab_style)
    
    def switch_page(self, index: int):
        """페이지 전환"""
        self.stacked.setCurrentIndex(index)
        self.btn_main_page.setChecked(index == 0)
        self.btn_comex_page.setChecked(index == 1)

