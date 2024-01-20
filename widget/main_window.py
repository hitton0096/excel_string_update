
import openpyxl as op

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow, 
    QPushButton,
    QVBoxLayout,
    QWidget
)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Super simple sheet update")
        # self.resize(800, 400)
        # self.work_book = op.load_workbook("test.xlsx")
        # active_sheet = self.work_book.active
        # print(active_sheet.cell(column=1, row=1).value)
        # print(active_sheet['A1'].value)

        # top_widget = QWidget()
        # top_widget_layout = QVBoxLayout(top_widget)
        base_layout = QVBoxLayout()
        select_sheet_label = QLabel("更新するシートファイルを選択してください。")


        self.line_edit = QLineEdit()
        self.line_edit.setReadOnly(True)
        button = QPushButton("参照")
        button.clicked.connect(self.the_button_was_clicked)

        select_layout = QHBoxLayout()
        select_layout.addWidget(self.line_edit)
        select_layout.addWidget(button)

        base_layout.addWidget(select_sheet_label)
        base_layout.addLayout(select_layout)
        # addStretchを入れることで空白分を埋めるというか伸ばす
        base_layout.addStretch(1)
        # top_widget.setFixedHeight(20)

        central_widget = QWidget()
        central_widget.setLayout(base_layout)
        self.setCentralWidget(central_widget)

    def the_button_was_clicked(self):
        print("Clicked!")
        caption = "Select Directory"
        dir = ""
        filter = "Excel (*.xlsx)"
        selectedFilter = ""
        options = QFileDialog.Options()
        fileNames = QFileDialog.getExistingDirectory(
            None,
            caption,
            dir,
            # filter,
            # selectedFilter,
            options
        )
        self.line_edit.setText(fileNames)
        print(fileNames)

    def the_button_was_toggled(self, checked):
        print("Checked?", checked)
