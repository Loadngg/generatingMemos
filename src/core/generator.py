import os
import platform
import subprocess
from typing import List, Callable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2pdf import convert
from transliterate import translit

from src.core.models.record import Record


def _convert_docx_to_pdf(system: str, input_filepath: str, output_dir: str, output_filename: str):
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


def generate(template_path: str, records: List[Record], output_path: str, update_info_label: Callable,
             enable_generating: Callable) -> None:
    last_generated_name: str = ''
    system = platform.system()

    for item in records:
        item_fields = item.get_fields()

        second_name = item.full_name[1].split(' ')[0]

        if last_generated_name != item.full_name[1]:
            last_generated_name = item.full_name[1]

        doc = Document(template_path)

        table_text: str = doc.tables[0].rows[0].cells[1].text
        doc.tables[0].rows[0].cells[1].text = table_text.replace(item.institute[0], item.institute[1])
        doc.tables[0].rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        for p in doc.paragraphs:
            for field in item_fields:
                key = field[0]
                value = field[1]
                p.text = p.text.replace(key, value)

        output_name = (translit(second_name, language_code='ru', reversed=True)
                       + ' '
                       + translit(' '.join(item.title[1].split(' ')[:3]), language_code='ru', reversed=True)
                       + ' '
                       + item.date[1])

        output_file = output_path + output_name + '.docx'
        pdf_path = output_path + second_name

        doc.save(output_file)

        if not os.path.exists(pdf_path):
            os.makedirs(pdf_path)

        _convert_docx_to_pdf(system, output_file, pdf_path, output_name)

        update_info_label(f"сгенерировано для {second_name}")

    update_info_label(f'файлы были успешно сгенерированы - {len(records)}')
    enable_generating()
