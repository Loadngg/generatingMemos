from collections.abc import Callable
from functools import lru_cache

import openpyxl
from openpyxl import Workbook
from xls2xlsx import XLS2XLSX


def _xls2xlsx(file_path: str) -> openpyxl.Workbook:
    x2x = XLS2XLSX(file_path)
    return x2x.to_xlsx()


class Uploader:
    _book: openpyxl.Workbook = None
    _already_uploaded_books: [str] = []

    def clear(self) -> None:
        self._book = None
        self._already_uploaded_books.clear()

    @lru_cache(maxsize=None)
    def upload(
            self,
            filename: str,
            enable_generating: Callable,
    ) -> None:
        if filename in self._already_uploaded_books:
            return

        self._already_uploaded_books.append(filename)
        book = _xls2xlsx(filename) if filename.endswith(".xls") else openpyxl.load_workbook(filename)
        self._book = book

        enable_generating()

    def get_book(self) -> Workbook:
        return self._book
