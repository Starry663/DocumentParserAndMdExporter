# file_selector.py
from PyQt5.QtWidgets import QMainWindow, QLabel, QLineEdit, QPushButton, QFileDialog, QHBoxLayout, QVBoxLayout, QWidget
from main_window import MainWindow

class FileSelectorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("选择文档文件")
        self.resize(400, 100)
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 文件选择区域
        file_layout = QHBoxLayout()
        file_label = QLabel("选择文件:")
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        browse_btn = QPushButton("浏览...")
        file_layout.addWidget(file_label)
        file_layout.addWidget(self.file_path_edit, 1)
        file_layout.addWidget(browse_btn)
        layout.addLayout(file_layout)

        # 下一步按钮
        self.next_btn = QPushButton("下一步")
        self.next_btn.setEnabled(False)
        layout.addWidget(self.next_btn)

        # 信号连接
        browse_btn.clicked.connect(self.open_file_dialog)
        self.next_btn.clicked.connect(self.open_main_window)
        self.selected_file = None

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "打开文件", "", "Documents (*.docx *.pdf)")
        if file_path:
            self.selected_file = file_path
            self.file_path_edit.setText(file_path)
            self.next_btn.setEnabled(True)

    def open_main_window(self):
        if not self.selected_file:
            return
        self.main_window = MainWindow(self.selected_file)
        self.main_window.show()
        self.close()
