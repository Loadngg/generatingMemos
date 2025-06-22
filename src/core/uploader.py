from collections.abc import Callable
from functools import lru_cache
from typing import List

import openpyxl
from xls2xlsx import XLS2XLSX


def _xls2xlsx(file_path: str) -> openpyxl.Workbook:
    x2x = XLS2XLSX(file_path)
    return x2x.to_xlsx()


class Uploader:
    _book: openpyxl.Workbook = None
    _already_uploaded_books: List[str] = []

    def clear(self) -> None:
        self._book = None
        self._already_uploaded_books.clear()

    @lru_cache(maxsize=None)
    def upload(self, filename: str, update_info_callback: Callable, enable_generating: Callable) -> None:
        if filename in self._already_uploaded_books:
            self._finish_uploading(update_info_callback, enable_generating)
            return

        self._already_uploaded_books.append(filename)
        book = _xls2xlsx(filename) if filename.endswith(".xls") else openpyxl.load_workbook(filename)
        self._book = book
        self._finish_uploading(update_info_callback, enable_generating)

    @staticmethod
    def _finish_uploading(update_info_callback: Callable, enable_generating: Callable) -> None:
        update_info_callback("Подготовка данных завершена")
        enable_generating()

    def get_book(self) -> openpyxl.Workbook:
        return self._book
