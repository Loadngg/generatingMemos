from typing import Dict, List, Union

from src.core.models.field import Field


class Record:
    institute: Field
    generating_date: Field
    full_name: Field
    group: Field
    type_field: Field
    events: Dict[int, List[Field]]
    event_title: Field
    event_date: Field

    def __init__(
            self,
            institute: Field,
            generating_date: Field,
            full_name: Field,
            group: Field,
            type_field: Field,
            events: Dict[int, List[Field]],
            event_title: Field,
            event_date: Field
    ) -> None:
        self.institute = institute
        self.generating_date = generating_date
        self.full_name = full_name
        self.group = group
        self.type_field = type_field
        self.events = events
        self.event_title = event_title
        self.event_date = event_date

    def get_fields(self) -> List[Union[Field, Dict[int, List[Field]]]]:
        fields = [
            self.institute,
            self.generating_date,
            self.full_name,
            self.group,
            self.type_field,
            self.event_date,
            self.event_title,
            self.events
        ]
        return fields

    def __str__(self) -> str:
        sections = [
            f"Институт: {self.institute}",
            f"Дата генерации: {self.generating_date}",
            f"ФИО: {self.full_name}",
            f"Группа: {self.group}",
            f"Тип: {self.type_field}",
        ]

        if self.event_date or self.event_title:
            sections.append("Одиночное мероприятие:")
            sections.append(f"  Название: {self.event_title}")
            sections.append(f"  Дата: {self.event_date}")

        if self.events:
            sections.append("Нумерованные мероприятия:")
            for event_id, fields in sorted(self.events.items()):
                event_date = fields[0] if len(fields) > 0 else "Нет даты"
                event_title = fields[1] if len(fields) > 1 else "Нет названия"

                sections.append(f"  Мероприятие {event_id}:")
                sections.append(f"    Дата: {event_date}")
                sections.append(f"    Название: {event_title}")

        return "\n" + "\n".join(sections)
