"""실제 적용된 폰트 확인 스크립트"""
import sys
from PySide6.QtWidgets import QApplication, QPushButton, QWidget, QVBoxLayout
from PySide6.QtGui import QFont, QFontDatabase, QFontInfo

app = QApplication(sys.argv)

# main.py와 동일한 설정
font = QFont("맑은 고딕", 9)
font.setWeight(QFont.Weight.Normal)
font.setItalic(False)
app.setFont(font)

print("=" * 60)
print("시스템에 등록된 맑은 고딕 관련 폰트:")
print("=" * 60)
db = QFontDatabase()
families = db.families()
malgun_fonts = [f for f in families if '맑은' in f or 'Malgun' in f or 'malgun' in f.lower()]
if malgun_fonts:
    for font_name in malgun_fonts:
        print(f"  - {font_name}")
else:
    print("  (맑은 고딕 폰트를 찾을 수 없습니다)")

print("\n" + "=" * 60)
print("테스트 버튼 생성 및 폰트 확인:")
print("=" * 60)

# 테스트 윈도우
window = QWidget()
layout = QVBoxLayout()

# 1. 기본 QPushButton (app.setFont 적용)
btn1 = QPushButton("기본 버튼 (app.setFont)")
layout.addWidget(btn1)

# 2. 스타일시트로 "맑은 고딕" (따옴표) 적용
btn2 = QPushButton("맑은 고딕 버튼 (따옴표)")
btn2.setStyleSheet("font-family: '맑은 고딕'; font-size: 10pt;")
layout.addWidget(btn2)

# 3. 스타일시트로 "맑은 고딕" 적용
btn3 = QPushButton("맑은 고딕 버튼")
btn3.setStyleSheet("font-family: 맑은 고딕; font-size: 10pt;")
layout.addWidget(btn3)

# 4. 스타일시트로 "Malgun Gothic" (영문 이름) 적용
btn4 = QPushButton("Malgun Gothic 버튼")
btn4.setStyleSheet("font-family: Malgun Gothic; font-size: 10pt;")
layout.addWidget(btn4)

window.setLayout(layout)
window.setWindowTitle("폰트 확인")
window.resize(400, 300)
window.show()

# 이벤트 루프 실행 후 폰트 정보 확인
QApplication.processEvents()

def print_font_info(btn, name):
    font = btn.font()
    fi = QFontInfo(font)
    print(f"\n{name}:")
    print(f"  요청한 폰트 패밀리: {font.family()}")
    print(f"  실제 렌더링 폰트: {fi.family()}")
    print(f"  요청한 폰트 크기: {font.pointSize()}pt")
    print(f"  실제 폰트 크기: {fi.pointSize()}pt")
    print(f"  폰트 스타일: {font.style()}")
    print(f"  폰트 굵기: {font.weight()}")
    if font.family() != fi.family():
        print(f"  ⚠️ 경고: 요청한 폰트와 실제 폰트가 다릅니다!")

print("\n" + "=" * 60)
print("실제 적용된 폰트 정보:")
print("=" * 60)
print_font_info(btn1, "1. 기본 버튼 (app.setFont)")
print_font_info(btn2, "2. 맑은 고딕 (따옴표)")
print_font_info(btn3, "3. 맑은 고딕")
print_font_info(btn4, "4. Malgun Gothic (영문 이름)")

print("\n" + "=" * 60)
print("윈도우가 열렸습니다. 폰트를 확인한 후 창을 닫으세요.")
print("=" * 60)

sys.exit(app.exec())
