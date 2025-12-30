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
        
        # 페이지네이션 (좌측 상단)
        page_nav = QHBoxLayout()
        self.btn_main_page = QPushButton("메인")
        self.btn_comex_page = QPushButton("comex 관리")
        self.btn_main_page.setCheckable(True)
        self.btn_comex_page.setCheckable(True)
        self.btn_main_page.setChecked(True)
        
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
        layout.addLayout(page_nav)
        layout.addWidget(self.stacked)
        self.setLayout(layout)
        
        # 페이지 전환 연결
        self.btn_main_page.clicked.connect(lambda: self.switch_page(0))
        self.btn_comex_page.clicked.connect(lambda: self.switch_page(1))
    
    def switch_page(self, index: int):
        """페이지 전환"""
        self.stacked.setCurrentIndex(index)
        self.btn_main_page.setChecked(index == 0)
        self.btn_comex_page.setChecked(index == 1)

