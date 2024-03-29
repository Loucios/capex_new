from dataclasses import InitVar, dataclass, field, fields

from help import HelpClass
from titles import Titles


@dataclass
class ChapterDirect:
    number: int
    name: str


@dataclass
class NetworkCostIndicator:
    number: int
    diameter: int
    const_chanal: float
    const_chanalless: float
    const_elevated: float
    const_basement: float
    const_casing: float
    const_mixed: float
    reconst_chanal: float
    reconst_chanalless: float
    reconst_elevated: float
    reconst_basement: float
    reconst_casing: float
    reconst_mixed: float


@dataclass
class CTPCostIndicator:
    number: int
    power: float
    const_indicators: float
    reconst_indicators: float


@dataclass
class SourceCostIndicator:
    number: int
    power_range: str
    const_indicators: float
    reconst_indicators: float
    reconst2_indicators: float
    power: float
    units: str


@dataclass
class TfuUnitCost:
    number: int
    const_indicators: float
    reconst_indicators: float
    reconst2_indicators: float


@dataclass
class Stage:
    number: int
    stage: str
    source: int
    network: int


@dataclass
class DeflatorIndex:
    number: int
    year: int
    index: float
    cumulative: float = 1.0


@dataclass
class Terms:
    number: int
    cost_year: int
    start_year: int
    end_year: int


@dataclass
class NDS:
    number: int
    nds: float


@dataclass
class TSOList:
    number: int
    name: str
    short_name: str


@dataclass
class BaseEvent:
    number: int
    eto_name: str
    tso_name: str
    district: str
    source_type: str
    source_id: int
    source_name: str
    event_title: str
    event_years: str
    mw: float | None
    laying_type: str
    length: float
    diameter: float
    gh: float | None
    th: float | None
    design_work: str | None
    inv_program: str | None
    base_directions: str
    event_type: str
    chapter_7_directions: int
    chapter_8_directions: int
    source_ch12_directions: int
    network_ch12_directions: int
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

    stages: InitVar[list]
    deflator_indices: InitVar[list]
    terms: InitVar[list]
    nds: InitVar[list]
    total: float = field(init=False)
    obj_type: str = field(init=False, repr=False)
    row_name = Titles.row_name_first

    def set_capex(self, *args):
        pass

    def set_capex_flow(self, stages, deflator_indices, terms, nds) -> None:
        total = 0
        self.total = 0
        start_year, end_year = self.get_first_and_last_year()
        time = end_year - start_year + 1

        for year in range(Titles.begining_year, Titles.ending_year + 1):
            value = getattr(self, 'year_' + str(year))
            if not isinstance(value, (float, int)):
                setattr(self, 'year_' + str(year), 0)
                value = 0
            else:
                value *= (1 + nds[0].nds)
            if start_year <= year <= end_year:
                total += value

        if not total:
            design_rate = getattr(stages[0], self.obj_type) / 100
            nds = (1 + nds[0].nds)
            for year in range(start_year, end_year + 1):
                value = 0
                deflator = deflator_indices[
                        year - Titles.begining_year
                    ].cumulative
                if time == 1:
                    value = self.total_cost * deflator * nds
                else:
                    if year == start_year:
                        value = self.total_cost * design_rate * deflator * nds
                    else:
                        value = (
                            self.total_cost * (1 - design_rate) * deflator /
                            (time - 1) * nds
                        )
                setattr(self, 'year_' + str(year), value)
                if terms[0].start_year <= year <= terms[0].end_year:
                    self.total += value

    def get_first_and_last_year(self) -> tuple[int]:
        years = str(self.event_years)
        if years[:4] == years[-4:]:
            first = last = int(years)
        else:
            first = int(years[:4])
            last = int(years[-4:])
        return first, last

    def set_none_to_zero(self) -> None:
        for attribute in fields(self):
            if attribute.init and getattr(self, attribute.name) is None:
                setattr(self, attribute.name, 0)


@dataclass
class SourceEvent(BaseEvent):
    laying_type: str = field(init=False, default='')
    length: float = field(init=False, default=0.0)
    diameter: float = field(init=False, default=0.0)
    chapter_8_directions: int = field(init=False, default='')
    network_ch12_directions: int = field(init=False, default='')

    source_unit_costs: InitVar[list]
    tfu_unit_costs: InitVar[list]
    obj_type: str = field(init=False, default='source')

    def __post_init__(self, stages: list[Stage],
                      deflator_indices: list[DeflatorIndex],
                      terms: list[Terms], nds: list[NDS],
                      source_unit_costs: list[SourceCostIndicator],
                      tfu_unit_costs: list[TfuUnitCost]):
        if self.total_cost is None:
            self.set_capex(source_unit_costs, tfu_unit_costs)
        self.set_none_to_zero()
        self.set_capex_flow(stages, deflator_indices, terms, nds)

    def set_capex(self, source_unit_costs, tfu_unit_costs) -> None:
        unit_types = {
            Titles.const_indicators: 'const_indicators',
            Titles.reconst_indicators: 'reconst_indicators',
            Titles.reconst2_indicators: 'reconst2_indicators'
        }
        if self.mw is not None:
            unit_cost = getattr(
                tfu_unit_costs[0], unit_types[self.event_type]
            )
            power = self.mw
        else:
            if self.th is not None:
                power = self.th * 0.6
            else:
                power = self.gh
            item = HelpClass.binary_search(power, source_unit_costs, 'power')
            unit_cost = getattr(item, unit_types[self.event_type])
        self.total_cost = power * unit_cost


@dataclass
class NetworkEvent(BaseEvent):
    mw: float = field(init=False, default=0.0)
    th: float = field(init=False, default=0.0)
    chapter_7_directions: int = field(init=False, default='')
    source_ch12_directions: int = field(init=False, default='')

    network_unit_costs: InitVar[list]
    ctp_unit_costs: InitVar[list]
    obj_type: str = field(init=False, default='network')

    def __post_init__(self, stages: list[Stage],
                      deflator_indices: list[DeflatorIndex],
                      terms: list[Terms], nds: list[NDS],
                      network_unit_costs: list[SourceCostIndicator],
                      ctp_unit_costs: list[TfuUnitCost]):
        if self.total_cost is None:
            self.set_capex(network_unit_costs, ctp_unit_costs)
        self.set_none_to_zero()
        self.set_capex_flow(stages, deflator_indices, terms, nds)

    def set_capex(self, network_unit_costs, ctp_unit_costs) -> None:
        ctp_event_types = {
            Titles.const_indicators: 'const_indicators',
            Titles.reconst_indicators: 'reconst_indicators',
        }
        network_event_types = {
            Titles.const_indicators: 'const',
            Titles.reconst_indicators: 'reconst',
        }
        laying_types = {
            Titles.chanal: 'chanal',
            Titles.chanalless: 'chanalless',
            Titles.elevated: 'elevated',
            Titles.basement: 'basement',
            Titles.casing: 'casing',
            Titles.mixed: 'mixed',
        }
        if self.diameter is not None:
            item = HelpClass.binary_search(self.diameter,
                                           network_unit_costs,
                                           'diameter')
            laying_type = laying_types[self.laying_type]
            event_type = network_event_types[self.event_type]
            unit_cost = getattr(item, f'{event_type}_{laying_type}')
            self.total_cost = unit_cost * self.length / 1000
        else:
            item = HelpClass.binary_search(self.gh, ctp_unit_costs, 'power')
            unit_cost = getattr(item, ctp_event_types[self.event_type])
            self.total_cost = unit_cost * self.gh
