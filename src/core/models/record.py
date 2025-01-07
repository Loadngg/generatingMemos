from typing import List


class Record:
    institute: List[str] = []
    generating_date: List[str] = []
    full_name: List[str] = []
    group: List[str] = []
    title: List[str] = []
    date: List[str] = []
    type: List[str] = ["[type]"]

    def get_record(self) -> str:
        result = " ".join(
            [
                self.title[1],
                self.date[1],
                self.full_name[1],
                self.group[1],
                self.institute[1],
                self.type[1],
                self.generating_date[1]
            ]
        )
        return result

    def get_fields(self) -> List[List[str]]:
        return [
            self.title,
            self.date,
            self.full_name,
            self.group,
            self.institute,
            self.type,
            self.generating_date,
        ]
