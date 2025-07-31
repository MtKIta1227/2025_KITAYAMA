import sys
from PyQt6.QtWidgets import QApplication
from app.main_window import OceanDirectApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = OceanDirectApp()
    win.show()
    sys.exit(app.exec())