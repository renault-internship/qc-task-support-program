"""
미리보기 컨테이너 - 엑셀 테이블
"""
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout,
    QTableView, QStyledItemDelegate, QStyleOptionViewItem, QStyle, QLabel, QApplication
)
from PySide6.QtGui import QPen, QPainter, QColor, QBrush


class NoElideDelegate(QStyledItemDelegate):
    """말줄임표 없이 전체 텍스트 표시하는 Delegate - 텍스트 직접 그리기"""
    def paint(self, painter, option, index):
        # 스타일 옵션 초기화
        self.initStyleOption(option, index)
        
        # 텍스트를 제거하여 배경만 그리기
        option.text = ""  # 텍스트를 비워서 배경만 그림
        
        # 배경과 테두리 그리기 (선택 상태, hover 등)
        style = option.widget.style() if option.widget else QStyle()
        style.drawControl(QStyle.CE_ItemViewItem, option, painter, option.widget)
        
        # 텍스트 가져오기
        text = index.data(Qt.DisplayRole)
        if text:
            painter.save()
            
            # 텍스트 색상 설정 (선택된 경우와 일반 경우 구분)
            if option.state & QStyle.State_Selected:
                painter.setPen(QPen(option.palette.highlightedText().color()))
            else:
                painter.setPen(QPen(option.palette.text().color()))
            
            # 텍스트를 말줄임 없이 직접 그리기
            text_rect = option.rect.adjusted(4, 0, -4, 0)  # 좌우 패딩
            painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, str(text))
            
            painter.restore()


class SpinnerWidget(QWidget):
    """회전 스피너 위젯"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(50, 50)
        self._angle = 0
        
        # QTimer를 사용한 회전 애니메이션
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._rotate)
        self.timer.setInterval(50)  # 50ms마다 업데이트
    
    def _rotate(self):
        """회전 업데이트"""
        self._angle = (self._angle + 10) % 360  # 10도씩 회전
        self.update()  # 화면 갱신
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # 중앙 좌표
        center_x = self.width() / 2
        center_y = self.height() / 2
        radius = min(self.width(), self.height()) / 2 - 8
        
        # 원형 스피너 그리기 (8개의 점)
        import math
        for i in range(8):
            # 각 점의 각도 (현재 회전 각도 + 점의 기본 각도)
            angle_deg = self._angle + (i * 45)  # 45도씩 떨어진 8개의 점
            angle_rad = math.radians(angle_deg)
            
            # 점의 위치 계산
            x = center_x + radius * math.cos(angle_rad)
            y = center_y + radius * math.sin(angle_rad)
            
            # 투명도 계산 (첫 번째 점이 가장 진하고, 마지막이 가장 투명)
            alpha = int(255 - (i * 30))  # 255, 225, 195, ... , 45
            if alpha < 50:
                alpha = 50  # 최소 투명도 보장
            
            color = QColor("#2196F3")
            color.setAlpha(alpha)
            
            painter.setBrush(QBrush(color))
            painter.setPen(Qt.NoPen)
            
            # 작은 원 그리기
            dot_size = 5
            painter.drawEllipse(int(x - dot_size/2), int(y - dot_size/2), dot_size, dot_size)
    
    def start(self):
        """애니메이션 시작"""
        self._angle = 0
        self.timer.start()
        self.show()  # 위젯 표시 확인
    
    def stop(self):
        """애니메이션 중지"""
        self.timer.stop()


class PreviewContainer(QWidget):
    """엑셀 미리보기 컨테이너"""
    def __init__(self, parent=None):
        super().__init__(parent)
        
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # 엑셀 테이블
        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(False)
        
        # 말줄임표 없이 전체 텍스트 표시
        self.table.setItemDelegate(NoElideDelegate(self.table))
        
        layout.addWidget(self.table, 1)
        
        self.setLayout(layout)
        
        # 로딩 오버레이 위젯
        self.loading_overlay = QWidget(self)
        self.loading_overlay.setStyleSheet("""
            QWidget {
                background-color: rgba(255, 255, 255, 240);
            }
        """)
        self.loading_overlay.hide()
        self.loading_overlay.setAttribute(Qt.WA_TransparentForMouseEvents, False)  # 클릭 이벤트 차단
        self.loading_overlay.setAttribute(Qt.WA_NoSystemBackground, False)
        self.loading_overlay.raise_()  # 항상 최상위에 표시
        
        overlay_layout = QVBoxLayout(self.loading_overlay)
        overlay_layout.setAlignment(Qt.AlignCenter)
        overlay_layout.setSpacing(15)
        overlay_layout.setContentsMargins(0, 0, 0, 0)
        
        # 회전 스피너
        self.spinner = SpinnerWidget()
        overlay_layout.addWidget(self.spinner, 0, Qt.AlignCenter)
        
        # 로딩 메시지
        self.loading_message = QLabel()
        self.loading_message.setAlignment(Qt.AlignCenter)
        self.loading_message.setStyleSheet("""
            QLabel {
                font-size: 12pt;
                color: #2196F3;
                font-weight: bold;
            }
        """)
        overlay_layout.addWidget(self.loading_message)
    
    def get_table(self) -> QTableView:
        """엑셀 테이블 반환"""
        return self.table
    
    def show_loading(self, message: str = "처리 중"):
        """로딩 애니메이션 표시"""
        self.loading_message.setText(message)
        # 오버레이 크기 설정 (위젯 크기가 0이면 기본값 사용)
        width = self.width() if self.width() > 0 else 800
        height = self.height() if self.height() > 0 else 600
        self.loading_overlay.setGeometry(0, 0, width, height)
        self.loading_overlay.show()
        self.loading_overlay.raise_()  # 테이블 위에 표시
        self.loading_overlay.update()  # 즉시 업데이트
        self.spinner.start()
        # UI 업데이트 강제
        QApplication.processEvents()
    
    def hide_loading(self):
        """로딩 애니메이션 숨김"""
        self.spinner.stop()
        self.loading_overlay.hide()
    
    def resizeEvent(self, event):
        """위젯 크기 변경 시 오버레이 크기도 조정"""
        super().resizeEvent(event)
        if self.loading_overlay.isVisible():
            self.loading_overlay.setGeometry(0, 0, self.width(), self.height())
    
    def showEvent(self, event):
        """위젯이 표시될 때 오버레이 위치 업데이트"""
        super().showEvent(event)
        if self.loading_overlay.isVisible():
            self.loading_overlay.setGeometry(0, 0, self.width(), self.height())

