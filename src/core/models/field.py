class Field:
    tag: str
    text: str

    def __init__(self, tag: str, text: str) -> None:
        self.tag = tag
        self.text = text

    def __str__(self) -> str:
        return f"{self.tag}: {self.text}"
