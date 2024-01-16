
import openpyxl as op
from PySide6.QtWidgets import (
    QMainWindow, 
    QPushButton,
    QFileDialog,
    QLineEdit,
    QVBoxLayout,
    QHBoxLayout,
    QWidget
)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel String Update")
        # self.work_book = op.load_workbook("test.xlsx")
        # active_sheet = self.work_book.active
        # print(active_sheet.cell(column=1, row=1).value)
        # print(active_sheet['A1'].value)
        h_layout = QHBoxLayout()
        central_widget = QWidget()
        central_widget.setLayout(h_layout)

        self.line_edit = QLineEdit()
        self.line_edit.setReadOnly(True)
        h_layout.addWidget(self.line_edit)

        button = QPushButton("Press Me!")
        button.clicked.connect(self.the_button_was_clicked)
        # button.setCheckable(True)
        # button.clicked.connect(self.the_button_was_toggled)
        h_layout.addWidget(button)


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
