import sys
from PySide6.QtWidgets import QApplication
from src.gui import MainWindow

def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(1200, 720)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
