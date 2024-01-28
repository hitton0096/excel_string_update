
import openpyxl as op

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow, 
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget
)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Super simple sheet updater")
        self.resize(350, 200)
        # 画面レイアウトは以下のサイトを参考に作成
        # https://stackoverflow.com/questions/12007807/create-qt-layout-with-fixed-height
        base_layout = QVBoxLayout()
        select_dir_label = QLabel("更新するファイルが入っているディレクトリを選択してください。")

        self.select_dir_edit = QLineEdit()
        self.select_dir_edit.setReadOnly(True)
        reference_button = QPushButton("参照")
        reference_button.clicked.connect(self.reference_button_click)

        select_layout = QHBoxLayout()
        select_layout.addWidget(self.select_dir_edit)
        select_layout.addWidget(reference_button)

        cell_address_label = QLabel("更新するセル番地を入力してください。")
        cell_address_edit = QLineEdit()
        cell_address_edit.setText("")
        cell_address_edit.setFixedWidth(40)
        cell_address_edit.setPlaceholderText("A1")
        # 英数字で大文字に自動変換
        cell_address_edit.setInputMask(">NNNNN")

        update_value_label = QLabel("更新する値を入力してください。")
        update_value_edit = QLineEdit()
        update_value_edit.setText("")

        exe_button = QPushButton("実行")
        exe_button.setFixedWidth(50)
        exe_button.clicked.connect(self.exe_button_click)

        exe_layout = QHBoxLayout()
        exe_layout.addStretch(1)
        exe_layout.addWidget(exe_button)
        exe_layout.addStretch(1)

        base_layout.addWidget(select_dir_label)
        base_layout.addLayout(select_layout)
        base_layout.addWidget(cell_address_label)
        base_layout.addWidget(cell_address_edit)
        base_layout.addWidget(update_value_label)
        base_layout.addWidget(update_value_edit)
        base_layout.addLayout(exe_layout)
        # addStretchを入れることで空白分を埋めるというか伸ばす
        base_layout.addStretch(1)
        # top_widget.setFixedHeight(20)

        central_widget = QWidget()
        central_widget.setLayout(base_layout)
        self.setCentralWidget(central_widget)

    def reference_button_click(self):
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
        self.select_dir_edit.setText(fileNames)
        # print(fileNames)

    def exe_button_click(self):
        print("exe_button_click")
        if error_message := self.validator():
            self.show_error_dialog(error_message) 
            return
        # self.work_book = op.load_workbook("test.xlsx")
        # active_sheet = self.work_book.active
        # print(active_sheet.cell(column=1, row=1).value)
        # print(active_sheet['A1'].value)

        # top_widget = QWidget()
        # top_widget_layout = QVBoxLayout(top_widget)

    def validator(self):
        if not self.select_dir_edit.text():
            return "ディレクトリパスが入力されていません"
    
    def show_error_dialog(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle("入力エラー")
        msg_box.setText(message)
        # msg_box.setInformativeText(message)
        msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg_box.exec()