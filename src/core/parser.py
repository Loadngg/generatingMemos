import calendar
from typing import Dict, List

import pymorphy2
from openpyxl import Workbook

from src.core.column_config import ColumnConfig
from src.core.models.date_type import DateType
from src.core.models.field import Field
from src.core.models.record import Record


def _get_month_name(month: int) -> str:
    name = calendar.month_name[month].lower()

    morph = pymorphy2.MorphAnalyzer()

    month_name = morph.parse(name)[0]
    month_name_loct = month_name.inflect({'loct'})
    return month_name_loct.word


def _get_date(value: str, date_type: str) -> str:
    date = value.strip().split(' ')[0]
    date = '.'.join(date.split('-')[::-1])

    if date_type == DateType.numbers.value:
        date = date + "г."
    else:
        _, month, _ = date.split('.')
        month = _get_month_name(int(month))
        date = month

    return date


def _get_record(column_config: ColumnConfig, row, generating_date: str, date_type: str) -> Record:
    generating_date_field = Field("[generating_date]", generating_date)
    full_name = Field("[full_name]", str(row[column_config.full_name_index].value).strip())
    group = Field("[group]", str(row[column_config.group_index].value).strip())
    institute = Field("[institute]", str(row[column_config.institute_index].value).strip())
    type_field = Field("[type]", str(row[column_config.type_field_index].value).strip())

    event_date: Field = None
    event_title: Field = None
    events: Dict[int, List[Field]] = {}

    date_index = column_config.event_date_index
    if date_index:
        date_str = _get_date(str(row[column_config.event_date_index].value), date_type)
        event_date = Field("[event_date]", date_str)

    name_index = column_config.event_name_index
    if name_index:
        event_title = Field("[event_title]", str(row[column_config.event_name_index].value).strip())

    events_config = column_config.events
    for num, data in events_config.items():
        date_str, title_str = None, None
        if data[0]:
            date_str = _get_date(str(row[data[0]].value), date_type)
        if data[1]:
            title_str = str(row[data[1]].value)

        events[num] = [
            Field(f"[event_date_{num}]", date_str),
            Field(f"[event_title_{num}]", title_str)
        ]

    record = Record(
        generating_date=generating_date_field,
        full_name=full_name,
        group=group,
        institute=institute,
        type_field=type_field,
        event_date=event_date,
        event_title=event_title,
        events=events
    )

    return record


def parse(book: Workbook, generating_date: str, date_type: str) -> List[Record]:
    result: List[Record] = []
    column_config = ColumnConfig(book)
    sheet = book.worksheets[0]

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[1].value is None:
            continue

        record = _get_record(column_config, row, generating_date, date_type)
        result.append(record)

    return result
