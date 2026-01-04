import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont
from qt_material import apply_stylesheet
from src.gui import MainWindow

def main():
    app = QApplication(sys.argv)
    
    # 전체 앱에 맑은 고딕 Regular 적용
    font = QFont("맑은 고딕", 10)
    font.setWeight(QFont.Weight.Normal)  # Regular (Normal weight)
    font.setItalic(False)  # Regular (not italic)
    app.setFont(font)
    
    # Qt-Material light_blue 테마 적용 (추가 폰트 설정)
    apply_stylesheet(
        app, 
        theme='light_blue.xml', 
        invert_secondary=True,
        extra={
            # 전역 폰트 설정
            '*': {
                'font-family': '맑은 고딕',
            },
            # 특정 위젯별 폰트 설정
            'QPushButton': {
                'font-family': '맑은 고딕',
                'font-size': '10pt',
                'font-weight': 'normal',  # Regular
                'border-radius': '0px',  # 탭 버튼의 둥근 모서리 제거
            },
            'QLabel': {
                'font-family': '맑은 고딕',
                'font-size': '10pt',
                'font-weight': 'normal',  # Regular
            },
            'QLineEdit': {
                'font-family': '맑은 고딕',
                'font-size': '10pt',
                'font-weight': 'normal',  # Regular
            },
            'QComboBox': {
                'font-family': '맑은 고딕',
                'font-size': '10pt',
                'font-weight': 'normal',  # Regular
            },
            'QTableWidget': {
                'font-family': '맑은 고딕',
                'font-size': '10pt',
                'font-weight': 'normal',  # Regular
            },
            'QTextEdit': {
                'font-family': '맑은 고딕',
                'font-size': '10pt',
                'font-weight': 'normal',  # Regular
            },
            'QGroupBox': {
                'font-family': '맑은 고딕',
            },
            'QHeaderView': {
                'font-family': '맑은 고딕',
            },
            'QAbstractItemView': {
                'font-family': '맑은 고딕',
            },
        }
    )
    
    # 전역 폰트 강제 적용 (Qt-Material 테마가 덮어쓸 수 있는 경우 대비)
    app.setStyleSheet(app.styleSheet() + """
        * {
            font-family: 맑은 고딕 !important;
            font-size: 10pt !important;
            font-weight: normal !important;
        }
    """)
    
    w = MainWindow()
    w.resize(1200, 720)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
