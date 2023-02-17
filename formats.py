from dataclasses import dataclass


@dataclass(frozen=True)
class SheetNames:
    source = 'Источники тепловой энергии'
    network = 'Тепловые сети'
    chapter_7_directions = 'Источники тепловой энергии по направлениям Главы 7'
    chapter_8_directions = 'Тепловые сети по направлениям Главы 8'
    source_ch12_directions = 'Источники тепловой энергии по направлениям Главы 12'
    network_ch12_directions = 'Тепловые сети по направлениям Главы 12'


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
