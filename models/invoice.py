
class Invoice:
    def __init__(self, number: str, year: int, month: str, day: int, value: str):
        self.number = number
        self.year = year
        self.month = month
        self.day = day
        self.date = f'{self.month} {self.day}, {self.year}'
        self.value = value
