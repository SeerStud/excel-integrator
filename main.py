import sys
import logging
from PySide6.QtWidgets import QApplication
from app.ui.main_window import MainWindow

logging.basicConfig(
    filename='app.log',
    filemode='w',
    level=logging.DEBUG,
    format='%(asctime)s %(levelname)s %(message)s'
)

def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(800, 600)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
