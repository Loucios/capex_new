from dataclasses import InitVar, dataclass, field, fields

from titles import Titles


@dataclass
class SummaryTable:
    number: int
    event_title: str
    events: InitVar[list]
    directions: str | None = None
    event_years: str = '-'
    mw: float = 0.0
    length: float = 0.0
    diameter: float = 0.0
    gh: float = 0.0
    th: float = 0.0
    total_cost: float = 0.0
    year_2020: float = 0.0
    year_2021: float = 0.0
    year_2022: float = 0.0
    year_2023: float = 0.0
    year_2024: float = 0.0
    year_2025: float = 0.0
    year_2026: float = 0.0
    year_2027: float = 0.0
    year_2028: float = 0.0
    year_2029: float = 0.0
    year_2030: float = 0.0
    year_2031: float = 0.0
    year_2032: float = 0.0
    year_2033: float = 0.0
    year_2034: float = 0.0
    year_2035: float = 0.0
    year_2036: float = 0.0
    year_2037: float = 0.0
    year_2038: float = 0.0
    year_2039: float = 0.0
    year_2040: float = 0.0
    year_2041: float = 0.0
    year_2042: float = 0.0
    year_2043: float = 0.0
    year_2044: float = 0.0
    year_2045: float = 0.0
    year_2046: float = 0.0
    year_2047: float = 0.0
    year_2048: float = 0.0
    year_2049: float = 0.0
    year_2050: float = 0.0
    total: float = 0.0
    amount: int = field(init=False, default=0)

    def __post_init__(self, events: list[object]) -> None:
        exclude_names = [
            'number', 'directions', 'event_title', 'event_years',
            'diameter', 'amount',
        ]
        diameter = 0
        length = 0
        amount = 0
        for event in events:
            # Если мероприятие относится к направлению
            # или мы создаем суммарную строку
            if (self.directions is None or
               getattr(event, self.directions) == self.number):
                amount += 1
                # Пробегаем все поля датакласса
                for attribute in fields(self):
                    name = attribute.name
                    # Если у нас мероприятия на сетях, то
                    # аккуратно обрабатываем средний диаметр
                    # if dict_summary_table[directions] == 'network':
                    if name == 'length':
                        length = getattr(event, name)
                    if name == 'diameter':
                        diameter = getattr(event, name)
                        self.diameter += diameter * length
                        diameter = 0
                        length = 0
                    # Исключаем поля, которые не суммируем
                    # или которые получаем извне
                    # или которые определяем особым способом
                    if name not in exclude_names:
                        # суммируем все показатели мероприятий
                        # по направлению
                        value = getattr(self, name)
                        value += getattr(event, name)
                        setattr(self, name, value)
        # определяем особым способом
        self.event_years = self.get_event_years()
        self.diameter = self.get_diameter()
        self.amount = amount

    def get_diameter(self) -> float:
        if self.length:
            return self.diameter / self.length
        return 0.0

    def get_event_years(self) -> str:
        '''Create terms like 2024-2026'''
        first_year = Titles.begining_year
        last_year = Titles.ending_year
        start_year = 0
        end_year = first_year
        for year in range(first_year, last_year + 1):
            if getattr(self, 'year_' + str(year)):
                if not start_year:
                    start_year = year
                end_year = year
        if not start_year:
            return '-'
        if start_year == end_year:
            return start_year
        return f'{start_year}-{end_year}'
