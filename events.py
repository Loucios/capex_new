from base_datas import (NDS, ChapterDirect, CTPCostIndicator, DeflatorIndex,
                        NetworkCostIndicator, NetworkEvent,
                        SourceCostIndicator, SourceEvent, Stage, Terms,
                        TfuUnitCost)
from openpyxl import load_workbook
from summary_datas import SummaryTable
from titles import Titles


class BaseMixin:
    '''Наименование таблицы в Excel -> Наименование датакласса'''
    classes = {
        Titles.stages: Stage,
        Titles.deflator_indices: DeflatorIndex,
        Titles.event_years: Terms,
        Titles.nds: NDS,
        Titles.source_unit_costs: SourceCostIndicator,
        Titles.tfu_unit_costs: TfuUnitCost,
        Titles.source_events: SourceEvent,
        Titles.network_unit_costs: NetworkCostIndicator,
        Titles.ctp_unit_costs: CTPCostIndicator,
        Titles.network_events: NetworkEvent,
        Titles.chapter_7_directions: ChapterDirect,
        Titles.chapter_8_directions: ChapterDirect,
        Titles.source_ch12_directions: ChapterDirect,
        Titles.network_ch12_directions: ChapterDirect,
    }

    '''Группа мероприятий по структуре направлений'''
    group_event = {
        Titles.chapter_7_directions: 'source_events',
        Titles.chapter_8_directions: 'network_events',
        Titles.source_ch12_directions: 'source_events',
        Titles.network_ch12_directions: 'network_events',
    }

    '''Наименование атрибута'''
    atributes = {
        Titles.chapter_7_directions: 'chapter_7_directions',
        Titles.chapter_8_directions: 'chapter_8_directions',
        Titles.source_ch12_directions: 'source_ch12_directions',
        Titles.network_ch12_directions: 'network_ch12_directions',
    }


class Events(BaseMixin):
    def __init__(self, filename: str) -> None:
        self.wb, self.table_names = self._get_wb_and_table_names(filename)

        # import secondary tables
        source_unit_costs = self._get_table(Titles.source_unit_costs)
        tfu_unit_costs = self._get_table(Titles.tfu_unit_costs)
        network_unit_costs = self._get_table(Titles.network_unit_costs)
        ctp_unit_costs = self._get_table(Titles.ctp_unit_costs)
        stages = self._get_table(Titles.stages)
        deflator_indices = self._get_table(Titles.deflator_indices)
        terms = self._get_table(Titles.event_years)

        cum_deflator = 1
        for index in deflator_indices:
            if index.year >= terms[0].cost_year:
                cum_deflator *= index.index
                index.cumulative = cum_deflator

        nds = self._get_table(Titles.nds)

        # import and saving main tables
        base_params = (stages, deflator_indices, terms, nds)
        source_params = (source_unit_costs, tfu_unit_costs)
        network_params = (network_unit_costs, ctp_unit_costs)
        self.source_events = self._get_table(Titles.source_events,
                                             *base_params, *source_params)
        self.network_events = self._get_table(Titles.network_events,
                                              *base_params, *network_params)

        # creating and saving summary tables
        self.chapter7_summary = self._get_summary_table(
                                                Titles.chapter_7_directions)
        self.source_ch12_sum = self._get_summary_table(
                                                Titles.chapter_8_directions)
        self.chapter8_summary = self._get_summary_table(
                                                Titles.source_ch12_directions)
        self.network_ch12_sum = self._get_summary_table(
                                                Titles.network_ch12_directions)

        self.wb, self.table_names = None, None

    def _get_wb_and_table_names(self, filename: str):
        wb = load_workbook(filename=filename, data_only=True)

        table_names = {}
        for worksheet in wb.sheetnames:
            for table in wb[worksheet].tables:
                table_names[table] = worksheet
        return wb, table_names

    def _get_table(self, table_name: str, *args) -> list[object]:
        worksheet = self.wb[self.table_names[table_name]]
        range = worksheet.tables[table_name].ref
        origin_data = worksheet[range][1:]
        import_data = []
        for row in origin_data:
            values = (cell.value for cell in row)
            import_data.append(self.classes[table_name](*values, *args))
        # Create a "Total" row
        if table_name in (Titles.source_events, Titles.network_events):
            values = (len(import_data) + 1, Titles.total, import_data)
            import_data.append(SummaryTable(*values))
        return import_data

    def _get_summary_table(self, direction_name):
        directions = self._get_table(direction_name)
        summary_tables = []
        events = getattr(self, self.group_event[direction_name])[:-1]
        atribute_name = self.atributes[direction_name]
        for direct in directions:
            values = (direct.number, direct.name, events, atribute_name)
            summary_tables.append(SummaryTable(*values))
        return summary_tables
