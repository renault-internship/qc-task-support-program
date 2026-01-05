"""
색상 팔레트 - 엑셀 스타일 드롭다운 메뉴
"""
from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QFrame
)
from PySide6.QtGui import QColor, QPainter, QPen


# 엑셀 테마 색상 (10개 테마 × 6개 색상) - 세로로 배치
EXCEL_THEME_COLORS = [
    # 테마 1: 흰색~회색
    [(255, 255, 255), (242, 242, 242), (216, 216, 216), (191, 191, 191), (165, 165, 165), (127, 127, 127)],
    # 테마 2: 검정~회색
    [(0, 0, 0), (127, 127, 127), (89, 89, 89), (63, 63, 63), (38, 38, 38), (12, 12, 12)],
    # 테마 3: 베이지
    [(238, 236, 225), (221, 217, 195), (196, 189, 151), (147, 137, 83), (73, 68, 41), (29, 27, 16)],
    # 테마 4: 파란색 1
    [(31, 73, 125), (198, 217, 240), (141, 179, 226), (84, 141, 212), (23, 54, 93), (15, 36, 62)],
    # 테마 5: 파란색 2
    [(79, 129, 189), (219, 229, 241), (184, 204, 228), (149, 179, 215), (54, 96, 146), (36, 64, 97)],
    # 테마 6: 빨간색
    [(192, 80, 77), (242, 220, 219), (229, 185, 183), (217, 150, 148), (149, 55, 52), (99, 36, 35)],
    # 테마 7: 초록색
    [(155, 187, 89), (235, 241, 221), (215, 227, 188), (195, 214, 155), (118, 146, 60), (79, 97, 40)],
    # 테마 8: 보라색
    [(128, 100, 162), (229, 224, 236), (204, 193, 217), (178, 162, 199), (95, 73, 122), (63, 49, 81)],
    # 테마 9: 청록색
    [(75, 172, 198), (219, 238, 243), (183, 221, 232), (146, 205, 220), (49, 133, 155), (32, 88, 103)],
    # 테마 10: 주황색
    [(247, 150, 70), (253, 234, 218), (251, 213, 181), (250, 192, 143), (227, 108, 9), (151, 72, 6)],
]

# 엑셀 표준 색상 (가로 1줄) - 10개
EXCEL_STANDARD_COLORS = [
    (112, 48, 160),   # 7030A0
    (0, 32, 96),      # 002060
    (0, 112, 192),    # 0070C0
    (0, 176, 240),    # 00B0F0
    (0, 176, 80),     # 00B050
    (146, 208, 80),   # 92D050
    (255, 255, 0),    # FFFF00
    (255, 192, 0),    # FFC000
    (255, 0, 0),      # FF0000
    (192, 0, 0),      # C00000
]


class ColorButton(QPushButton):
    """색상 선택 버튼"""
    def __init__(self, color: QColor, parent=None):
        super().__init__(parent)
        self.color = color
        self.setFixedSize(24, 24)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: rgb({color.red()}, {color.green()}, {color.blue()});
                border: 1px solid #CCC;
                border-radius: 2px;
            }}
            QPushButton:hover {{
                border: 2px solid #2196F3;
            }}
        """)


class ColorPaletteWidget(QWidget):
    """색상 팔레트 위젯 - QMenu에 들어갈 커스텀 위젯"""
    color_selected = Signal(QColor)
    color_cleared = Signal()
    
    def __init__(self, show_clear_button=True, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, False)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        
        # "색상 없음" 버튼 (선택적)
        if show_clear_button:
            clear_btn = QPushButton("🚫 색상 없음")
            clear_btn.setStyleSheet("""
                QPushButton {
                    text-align: left;
                    padding: 6px 12px;
                    border: 1px solid #CCC;
                    background-color: white;
                }
                QPushButton:hover {
                    background-color: #F0F0F0;
                    border: 1px solid #2196F3;
                }
            """)
            clear_btn.clicked.connect(lambda: self.color_cleared.emit())
            clear_btn.clicked.connect(self.close)
            layout.addWidget(clear_btn)
            
            # 구분선
            line = QFrame()
            line.setFrameShape(QFrame.HLine)
            line.setFrameShadow(QFrame.Sunken)
            layout.addWidget(line)
        
        # 제목 - 테마 색상
        title_label = QLabel("테마 색상")
        title_label.setStyleSheet("font-weight: bold; color: #555; font-size: 10pt;")
        layout.addWidget(title_label)
        
        # 테마 색상 그리드 (10개 테마 × 6개 색상, 세로 배치)
        theme_grid_layout = QGridLayout()
        theme_grid_layout.setSpacing(3)
        theme_grid_layout.setContentsMargins(0, 0, 0, 0)
        
        for col, theme_colors in enumerate(EXCEL_THEME_COLORS):
            for row, (r, g, b) in enumerate(theme_colors):
                color = QColor(r, g, b)
                color_btn = ColorButton(color, self)
                color_btn.clicked.connect(lambda checked, c=color: self._on_color_clicked(c))
                theme_grid_layout.addWidget(color_btn, row, col)
        
        layout.addLayout(theme_grid_layout)
        
        # 구분선
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)
        
        # 제목 - 표준 색상
        standard_title_label = QLabel("표준 색상")
        standard_title_label.setStyleSheet("font-weight: bold; color: #555; font-size: 10pt;")
        layout.addWidget(standard_title_label)
        
        # 표준 색상 그리드 (10개, 가로 1줄)
        standard_grid_layout = QGridLayout()
        standard_grid_layout.setSpacing(3)
        standard_grid_layout.setContentsMargins(0, 0, 0, 0)
        
        for col, (r, g, b) in enumerate(EXCEL_STANDARD_COLORS):
            color = QColor(r, g, b)
            color_btn = ColorButton(color, self)
            color_btn.clicked.connect(lambda checked, c=color: self._on_color_clicked(c))
            standard_grid_layout.addWidget(color_btn, 0, col)
        
        layout.addLayout(standard_grid_layout)
        self.setLayout(layout)
        
        # 스타일 설정
        self.setStyleSheet("""
            QWidget {
                background-color: white;
                border: 1px solid #CCC;
                border-radius: 4px;
            }
        """)
    
    def _on_color_clicked(self, color: QColor):
        """색상 버튼 클릭 시"""
        self.color_selected.emit(color)
        self.close()


class ColorToolButton(QPushButton):
    """색상 도구 버튼 - 현재 색상 표시"""
    color_selected = Signal(QColor)
    color_cleared = Signal()
    
    def __init__(self, button_type="fill", parent=None):
        super().__init__(parent)
        self.button_type = button_type  # "fill" or "font"
        self.current_color = None
        self.palette_widget = None
        
        # 버튼 텍스트 및 아이콘
        if button_type == "fill":
            self.setText("배경색")
        else:
            self.setText("글자색")
        
        self.setFixedHeight(39)        
        self.setMinimumWidth(60)
        self.clicked.connect(self._show_palette)
        self._update_style()
    
    def _update_style(self):
        """버튼 스타일 업데이트 (현재 색상 표시)"""
        if self.current_color:
            # 선택된 색상의 hex 값 (테두리용)
            border_color = f"#{self.current_color.red():02X}{self.current_color.green():02X}{self.current_color.blue():02X}"
            # 배경색/글자색 버튼 공통: 텍스트는 진한 회색, 배경은 기본, 테두리는 선택된 색상
            self.setStyleSheet(f"""
                QPushButton {{
                    text-align: center;
                    padding-top: 3px;
                    padding-bottom: 3px;
                    color: #555;
                    border: 2px solid {border_color};
                }}
                QPushButton:hover {{
                    border: 2px solid #2196F3;
                }}
            """)
        else:
            # 색상 미선택 시
            self.setStyleSheet("""
                QPushButton {
                    text-align: center;
                    padding-top: 4px;
                    padding-bottom: 4px;
                    color: #555;
                    border: 2px solid #CCC;
                }
                QPushButton:hover {
                    border: 2px solid #2196F3;
                }
            """)
    
    def _show_palette(self):
        """색상 팔레트 표시"""
        if self.palette_widget and self.palette_widget.isVisible():
            self.palette_widget.close()
            return
        
        self.palette_widget = ColorPaletteWidget(show_clear_button=True, parent=self)
        self.palette_widget.color_selected.connect(self._on_color_selected)
        self.palette_widget.color_cleared.connect(self._on_color_cleared)
        
        # 버튼 아래에 팔레트 표시
        button_pos = self.mapToGlobal(self.rect().bottomLeft())
        self.palette_widget.move(button_pos)
        self.palette_widget.show()
    
    def _on_color_selected(self, color: QColor):
        """색상 선택 시"""
        self.current_color = color
        self._update_style()
        self.color_selected.emit(color)
    
    def _on_color_cleared(self):
        """색상 제거 시"""
        self.current_color = None
        self._update_style()
        self.color_cleared.emit()
    
    def get_current_color(self):
        """현재 선택된 색상 반환"""
        return self.current_color
    
    def set_current_color(self, color: QColor):
        """현재 색상 설정"""
        self.current_color = color
        self._update_style()
