import os
import platform
import subprocess

from PyQt5.QtWidgets import QMainWindow, QMessageBox
from docx2pdf import convert
from transliterate import translit


def show_error(parent: QMainWindow, msg: str, is_error: bool = False):
    if is_error:
        QMessageBox.critical(parent, 'Произошла ошибка', msg)
        return

    QMessageBox.warning(parent, 'Предупреждение', msg)


def _convert_docx_to_pdf(input_filepath: str, output_dir: str, output_filename: str) -> None:
    system = platform.system()

    if system == "Windows":
        convert(input_filepath, f"{output_dir}/{output_filename}.pdf")
        return

    command = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_dir,
        input_filepath
    ]
    subprocess.run(command, check=True)


def save_pdf(output_file, pdf_path, output_name) -> None:
    if not os.path.exists(pdf_path):
        os.makedirs(pdf_path)

    _convert_docx_to_pdf(output_file, pdf_path, output_name)


def get_translit(value: str) -> str:
    return translit(value, language_code='ru', reversed=True)
