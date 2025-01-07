import calendar
from typing import List

import pymorphy2
from openpyxl.workbook import Workbook

from src.core.models.date_type import DateType
from src.core.models.record import Record


def _get_month_name(month: int) -> str:
    name = calendar.month_name[month].lower()

    morph = pymorphy2.MorphAnalyzer()

    month_name = morph.parse(name)[0]
    month_name_loct = month_name.inflect({'loct'})
    return month_name_loct.word


def parse(book: Workbook, generating_date: str, date_type: str) -> List[Record]:
    result: List[Record] = []
    sheet = book.worksheets[0]

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[1].value is None:
            continue

        record = Record()

        record.generating_date = ["[generating_date]", generating_date]

        record.title = ["[title]", str(row[1].value).strip()]
        record.full_name = ["[full_name]", str(row[3].value).strip()]
        record.group = ["[group]", str(row[4].value).strip()]
        record.institute = ["[institute]", str(row[5].value).strip()]
        record.type = ["[type]", str(row[6].value).strip()]

        date = str(row[2].value).strip().split(' ')[0]
        date = '.'.join(date.split('-')[::-1])

        if date_type == DateType.numbers.value:
            date = date + "г."
        else:
            day, month, year = date.split('.')
            month = _get_month_name(int(month))
            date = month

        record.date = ["[date]", date]

        print(record.generating_date)
        print(record.title)
        print(record.full_name)
        print(record.group)
        print(record.institute)
        print(record.type)
        print(record.date)
        print("----------------------------------")

        result.append(record)

    return result
