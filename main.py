# main.py
import sys
from PyQt5.QtWidgets import QApplication
from file_selector import FileSelectorWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    selector = FileSelectorWindow()
    selector.show()
    sys.exit(app.exec_())
