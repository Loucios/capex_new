from dataclasses import dataclass


@dataclass(frozen=True)
class SheetNames:
    begining = (
        'Планируемые капитальные вложения в реализацию мероприятий '
        'по новому строительству, реконструкции, техническому перевооружению '
        'и/или модернизации'
    )
    types = {'title_name': 0, 'sheet_name': 1}
    source = [' в части источников тепловой энергии', 'И']
    network = [' в части тепловых сетей и объектов на них', 'ТС']
    chapter_7_directions = [' в части источников тепловой энергии', 'И7']
    chapter_8_directions = [' в части тепловых сетей и объектов на них', 'ТС8']
    source_ch12_directions = [
        ' в части источников тепловой энергии '
        'в прогнозных ценах (тыс. руб. с НДС)',
        'И12'
    ]
    network_ch12_directions = [
        ' в части тепловых сетей и объектов на них '
        'в прогнозных ценах (тыс. руб. с НДС)',
        'ТС12'
    ]
    ending = ''
    tso_begining = ' в зоне деятельности '


@dataclass(frozen=True)
class SourceTable:
    number: int
    eto_name: str
    tso_name: str
    district: str
    source_name: str
    event_title: str
    event_years: str
    mw: float | None
    gh: float | None
    th: float | None
    total_cost: float | None
    year_2020: float | None
    year_2021: float | None
    year_2022: float | None
    year_2023: float | None
    year_2024: float | None
    year_2025: float | None
    year_2026: float | None
    year_2027: float | None
    year_2028: float | None
    year_2029: float | None
    year_2030: float | None
    year_2031: float | None
    year_2032: float | None
    year_2033: float | None
    year_2034: float | None
    year_2035: float | None
    year_2036: float | None
    year_2037: float | None
    year_2038: float | None
    year_2039: float | None
    year_2040: float | None
    year_2041: float | None
    year_2042: float | None
    year_2043: float | None
    year_2044: float | None
    year_2045: float | None
    year_2046: float | None
    year_2047: float | None
    year_2048: float | None
    year_2049: float | None
    year_2050: float | None
    total: float | None


@dataclass(frozen=True)
class NetworkTable:
    number: int
    eto_name: str
    tso_name: str
    district: str
    source_name: str
    event_title: str
    event_years: str
    laying_type: str
    length: float
    diameter: float
    gh: float | None
    total_cost: float | None
    year_2020: float | None
    year_2021: float | None
    year_2022: float | None
    year_2023: float | None
    year_2024: float | None
    year_2025: float | None
    year_2026: float | None
    year_2027: float | None
    year_2028: float | None
    year_2029: float | None
    year_2030: float | None
    year_2031: float | None
    year_2032: float | None
    year_2033: float | None
    year_2034: float | None
    year_2035: float | None
    year_2036: float | None
    year_2037: float | None
    year_2038: float | None
    year_2039: float | None
    year_2040: float | None
    year_2041: float | None
    year_2042: float | None
    year_2043: float | None
    year_2044: float | None
    year_2045: float | None
    year_2046: float | None
    year_2047: float | None
    year_2048: float | None
    year_2049: float | None
    year_2050: float | None
    total: float | None


@dataclass(frozen=True)
class Chapter7Directions:
    number: int
    eto_name: str
    tso_name: str
    district: str
    source_name: str
    event_title: str
    event_years: str
    mw: float | None
    gh: float | None
    th: float | None
    total_cost: float | None
    year_2020: float | None
    year_2021: float | None
    year_2022: float | None
    year_2023: float | None
    year_2024: float | None
    year_2025: float | None
    year_2026: float | None
    year_2027: float | None
    year_2028: float | None
    year_2029: float | None
    year_2030: float | None
    year_2031: float | None
    year_2032: float | None
    year_2033: float | None
    year_2034: float | None
    year_2035: float | None
    year_2036: float | None
    year_2037: float | None
    year_2038: float | None
    year_2039: float | None
    year_2040: float | None
    year_2041: float | None
    year_2042: float | None
    year_2043: float | None
    year_2044: float | None
    year_2045: float | None
    year_2046: float | None
    year_2047: float | None
    year_2048: float | None
    year_2049: float | None
    year_2050: float | None
    total: float | None


@dataclass(frozen=True)
class Chapter8directions:
    number: int
    eto_name: str
    tso_name: str
    district: str
    source_name: str
    event_title: str
    event_years: str
    laying_type: str
    length: float
    diameter: float
    gh: float | None
    total_cost: float | None
    year_2020: float | None
    year_2021: float | None
    year_2022: float | None
    year_2023: float | None
    year_2024: float | None
    year_2025: float | None
    year_2026: float | None
    year_2027: float | None
    year_2028: float | None
    year_2029: float | None
    year_2030: float | None
    year_2031: float | None
    year_2032: float | None
    year_2033: float | None
    year_2034: float | None
    year_2035: float | None
    year_2036: float | None
    year_2037: float | None
    year_2038: float | None
    year_2039: float | None
    year_2040: float | None
    year_2041: float | None
    year_2042: float | None
    year_2043: float | None
    year_2044: float | None
    year_2045: float | None
    year_2046: float | None
    year_2047: float | None
    year_2048: float | None
    year_2049: float | None
    year_2050: float | None
    total: float | None


@dataclass(frozen=True)
class Chapter12Directions:
    # number: int
    # event_title: str
    # source_name: str
    # mw: float | None
    # gh: float | None
    # th: float | None
    # event_years: str
    row_name: str = ''
    year_2020: float | None = 0
    year_2021: float | None = 0
    year_2022: float | None = 0
    year_2023: float | None = 0
    year_2024: float | None = 0
    year_2025: float | None = 0
    year_2026: float | None = 0
    year_2027: float | None = 0
    year_2028: float | None = 0
    year_2029: float | None = 0
    year_2030: float | None = 0
    year_2031: float | None = 0
    year_2032: float | None = 0
    year_2033: float | None = 0
    year_2034: float | None = 0
    year_2035: float | None = 0
    year_2036: float | None = 0
    year_2037: float | None = 0
    year_2038: float | None = 0
    year_2039: float | None = 0
    year_2040: float | None = 0
    year_2041: float | None = 0
    year_2042: float | None = 0
    year_2043: float | None = 0
    year_2044: float | None = 0
    year_2045: float | None = 0
    year_2046: float | None = 0
    year_2047: float | None = 0
    year_2048: float | None = 0
    year_2049: float | None = 0
    year_2050: float | None = 0
    # total: float | None
