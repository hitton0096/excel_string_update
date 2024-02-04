
import glob
import openpyxl as op
import os
import re
import shutil

from datetime import datetime
from PySide6.QtGui import QRegularExpressionValidator
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
        self.reg_cell_addr = r"^[A-Z][A-Z1-9]*[1-9]$"
        self.setWindowTitle("Super simple sheet updater")
        self.resize(350, 200)
        self.set_widget()

    def set_widget(self):
        # 画面レイアウトは以下のサイトを参考に作成
        # https://stackoverflow.com/questions/12007807/create-qt-layout-with-fixed-height
        base_layout = QVBoxLayout()
        select_dir_label = QLabel("更新するファイルが入っているディレクトリを選択してください。")

        self.select_dir_edit = QLineEdit()
        self.select_dir_edit.setPlaceholderText("C:\\Users\\user_name\\Documents")
        self.select_dir_edit.setReadOnly(True)
        reference_button = QPushButton("参照")
        reference_button.clicked.connect(self.reference_button_click)

        select_layout = QHBoxLayout()
        select_layout.addWidget(self.select_dir_edit)
        select_layout.addWidget(reference_button)

        sheet_name_label = QLabel("更新対象のシート名を入力してください。")
        self.sheet_name_edit = QLineEdit()
        self.sheet_name_edit.setText("")
        self.sheet_name_edit.setPlaceholderText("シート名")

        cell_address_label = QLabel("更新するセル番地を入力してください。")
        self.cell_address_edit = QLineEdit()
        self.cell_address_edit.setText("")
        self.cell_address_edit.setFixedWidth(40)
        self.cell_address_edit.setPlaceholderText("A1")
        # 英数字で大文字に自動変換
        # cell_address_edit.setInputMask(">NNNNN")
        cell_addr_validator = QRegularExpressionValidator(self.reg_cell_addr)
        self.cell_address_edit.setValidator(cell_addr_validator)

        update_value_label = QLabel("更新する値を入力してください。")
        self.update_value_edit = QLineEdit()
        self.update_value_edit.setText("")

        exe_button = QPushButton("実行")
        exe_button.setFixedWidth(50)
        exe_button.clicked.connect(self.exe_button_click)

        exe_layout = QHBoxLayout()
        exe_layout.addStretch(1)
        exe_layout.addWidget(exe_button)
        exe_layout.addStretch(1)

        base_layout.addWidget(select_dir_label)
        base_layout.addLayout(select_layout)
        base_layout.addWidget(sheet_name_label)
        base_layout.addWidget(self.sheet_name_edit)
        base_layout.addWidget(cell_address_label)
        base_layout.addWidget(self.cell_address_edit)
        base_layout.addWidget(update_value_label)
        base_layout.addWidget(self.update_value_edit)
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

    def exe_button_click(self):
        if error_message := self.validator():
            self.show_msg_dialog(error_message) 
            return
        try:
            self.update_file()
        except PermissionError as pe:
            self.show_msg_dialog("更新ファイルを開いている可能性がありますので閉じてから再度実行してください。")
    
    def update_file(self):
        select_dir_value = self.select_dir_edit.text()
        for file in glob.glob(select_dir_value + "/*.xlsx"):
            # back_で始まるバックアップファイル名を対象外
            if re.search(r"^back_[0-9]{14}_.*\.xlsx$", os.path.basename(file)):
                continue

            work_book = op.load_workbook(file)
            input_sheet_name = self.sheet_name_edit.text()
            if not input_sheet_name in work_book.sheetnames:
                self.show_msg_dialog(f"{file}\n入力されたシート名の{input_sheet_name}が見つかりません。\n処理を終了します。") 
                return
            # バックアップファイル名作成
            date_s = datetime.now().strftime("%Y%m%d%H%M%S")
            back_up_file_name = "back_" + date_s + "_" + os.path.basename(file)
            shutil.copy(file, select_dir_value + "/" + back_up_file_name)
            work_sheet = work_book[input_sheet_name]
            cell_address = self.cell_address_edit.text()
            print(work_sheet[cell_address].value)
            work_sheet[cell_address].value = self.update_value_edit.text()
            print(work_sheet[cell_address].value)
            work_book.save(file)
        self.show_msg_dialog(f"処理が終了しました。", "info") 

    def validator(self):
        select_dir_value = self.select_dir_edit.text()
        if not select_dir_value:
            return "ディレクトリパスが入力されていません。"
        if not os.path.isdir(select_dir_value):
            return f"選択された{select_dir_value}のディレクトリは存在しません。"

        if not self.sheet_name_edit.text():
            return "シート名が入力されていません。"

        cell_address_value = self.cell_address_edit.text()
        if not cell_address_value:
            return "セル番号が入力されていません。"
        
        if not re.search(self.reg_cell_addr, cell_address_value):
            return "入力されたセル番号が正しくありません。"

        if not self.update_value_edit.text():
            return "更新する値が入力されていません。"
    
    def show_msg_dialog(self, message, msg_type="error"):
        msg_box = QMessageBox()
        if msg_type == "error":
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setWindowTitle("入力エラー")
        if msg_type == "info":
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setWindowTitle("情報")
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        # msg_box.setInformativeText(message)
        # msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        msg_box.exec()