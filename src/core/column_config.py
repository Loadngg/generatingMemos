from typing import Any, Dict, List

from openpyxl import Workbook


class ColumnConfig:
    _sheet: Workbook

    institute_index: int = None
    full_name_index: int = None
    group_index: int = None
    type_field_index: int = None

    events: Dict[int, List[int]] = {}
    event_date_index: int = None
    event_name_index: int = None

    def __init__(self, book: Workbook) -> None:
        self._book = book
        self._manage()

    def _manage(self) -> None:
        sheet = self._book.worksheets[0]
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))

        temp_events: Dict[int, Dict[str, int]] = {}
        for idx, value in enumerate(header_row):
            if value is None:
                continue

            self._set_index("ФИО", "full_name_index", idx, value)
            self._set_index("Группа", "group_index", idx, value)
            self._set_index("Институт", "institute_index", idx, value)
            self._set_index("Тип", "type_field_index", idx, value)

            s_value = str(value).strip()
            self._process_event_column(s_value, idx, temp_events, "Дата", "date", "event_date_index")
            self._process_event_column(s_value, idx, temp_events, "Название", "name", "event_name_index")

        for num, data in temp_events.items():
            date, name = None, None
            if 'date' in data:
                date = data['date']
            if 'name' in data:
                name = data['name']

            self.events[num] = [date, name]

    def _process_event_column(self, s_value: str, idx: int, temp_events: Dict[int, Dict[str, int]],
                              prefix: str, event_key: str, attr_name: str) -> None:
        if s_value.startswith(prefix):
            parts = s_value.split()
            if len(parts) > 1 and parts[1].isdigit():
                num = int(parts[1])
                if num not in temp_events:
                    temp_events[num] = {}
                temp_events[num][event_key] = idx
            elif len(parts) == 1:
                setattr(self, attr_name, idx)

    def _set_index(self, substring: str, attr_name: str, index: int, value: Any) -> None:
        if substring in str(value):
            setattr(self, attr_name, index)

    def __str__(self) -> str:
        sections = [
            f"Институт: {self.institute_index}",
            f"ФИО: {self.full_name_index}",
            f"Группа: {self.group_index}",
            f"Тип мероприятия: {self.type_field_index}",
            f"Ненумерованное мероприятие: дата: {self.event_date_index}, название: {self.event_name_index}"
        ]

        if self.events:
            sections.append("Нумерованные мероприятия:")
            for num, data in sorted(self.events.items()):
                sections.append(f"  Мероприятие {num}: дата: {data[0]}, название: {data[1]}")
        else:
            sections.append("Нумерованные мероприятия: отсутствуют")

        return "\n".join(sections) + '\n'
