
import openpyxl as op
from PySide6.QtWidgets import QMainWindow

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel String Update")
        self.work_book = op.load_workbook("test.xlsx")
        active_sheet = self.work_book.active
        print(active_sheet.cell(column=1, row=1).value)
        print(active_sheet['A1'].value)
