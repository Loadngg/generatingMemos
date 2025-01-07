import os
import sys
from threading import Thread

from PyQt5 import QtWidgets
from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QFileDialog

from src.core.generator import generate
from src.core.models.date_type import DateType
from src.core.parser import parse
from src.core.uploader import Uploader
from src.ui.design import Ui_MainWindow


class App(QtWidgets.QMainWindow):
    output_path: str = "./"
    generating_date: str = None
    template_path: str = None

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.uploader = Uploader()

        self.init_ui()

    def init_ui(self) -> None:
        self.ui.btn_upload.clicked.connect(self.btn_upload_handler)
        self.ui.btn_pick_dir.clicked.connect(self.btn_pick_dir_handler)
        self.ui.btn_generate.clicked.connect(self.btn_generate_handler)
        self.ui.btn_pick_template.clicked.connect(self.btn_pick_template_handler)

        current_date = QDate.currentDate()
        self.ui.edit_date.setDate(current_date)

        self.ui.date_type_combo.addItems([DateType.numbers.value, DateType.words.value])

        self.ui.output_dir_label.setText("Папка вывода: " + os.getcwd())

        self.enable_generating(False)

    def enable_generating(self, flag: bool = True) -> None:
        self.ui.btn_generate.setEnabled(flag)

    def btn_pick_template_handler(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(self, filter="Docx (*.docx)")
        if not filename:
            return

        self.template_path = filename
        self.ui.template_label.setText(f"Шаблон: {filename.split('/')[-1]}")

    def btn_upload_handler(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(self, filter="Excel (*.xls *.xlsx)")
        if not filename:
            return

        self.uploader.clear()
        self.enable_generating(False)
        self.ui.data_label.setText(f"Данные: {filename.split('/')[-1]}")

        Thread(
            target=self.uploader.upload,
            args=(
                filename,
                self.enable_generating
            ),
        ).start()

    def btn_pick_dir_handler(self) -> None:
        path = QFileDialog.getExistingDirectory(self)
        if not path:
            return

        self.output_path = path + "/"
        self.ui.output_dir_label.setText("Папка вывода: " + self.output_path)

    def update_info_label(self, text):
        self.ui.info_label.setText(f"Статус: {text}")

    def btn_generate_handler(self) -> None:
        generating_date = self.ui.edit_date.text() + "г."

        if not self.template_path:
            self.update_info_label("Не выбран шаблон")
            return

        date_type = self.ui.date_type_combo.currentText()
        records = parse(self.uploader.get_book(), generating_date, date_type)

        self.enable_generating(False)
        self.update_info_label("В процессе...")

        Thread(
            target=generate,
            args=(
                self.template_path,
                records,
                self.output_path,
                self.update_info_label,
                self.enable_generating
            ),
        ).start()


if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    application = App()
    application.show()

    sys.exit(app.exec())
