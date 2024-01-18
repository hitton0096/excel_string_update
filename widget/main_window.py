
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

        # button1 = QPushButton('1')
        # button2 = QPushButton('2')
        # button3 = QPushButton('3')
        
        # layout = QGridLayout()
        # # レイアウトにウィジェットを追加
        # layout.addWidget(button1, 0, 0, 1, 2)
        # layout.addWidget(button2, 1, 0)
        # layout.addWidget(button3, 1, 1)

        # central_widget = QWidget()
        # central_widget.setLayout(layout)
        # self.setCentralWidget(central_widget)

        g_layout = QGridLayout()
        dir_select_label = QLabel("更新するシートファイルを作成してください。")
        # style = 'background-color: rgb(255, 50, 50);'
        # dir_select_label.setStyleSheet(style)
        # 列結合する場合に行結合しなくても最小値の1を設定する必要がある
        g_layout.addWidget(dir_select_label, 0, 0, 1, 2)

        self.line_edit = QLineEdit()
        self.line_edit.setReadOnly(True)
        g_layout.addWidget(self.line_edit, 1, 0, alignment=Qt.AlignmentFlag.AlignTop)

        button = QPushButton("参照")
        button.clicked.connect(self.the_button_was_clicked)
        g_layout.addWidget(button, 1, 1, alignment=Qt.AlignmentFlag.AlignTop)

        # g_layout.setRowStretch(0, 0)
        # g_layout.setRowStretch(1, 3)
        # g_layout.setRowMinimumHeight(0, 10)
        # これを入れないと0行目の高さが大きくなるため無理やり高さを出して確保
        g_layout.setRowMinimumHeight(1, 60)

        central_widget = QWidget()
        central_widget.setLayout(g_layout)
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
