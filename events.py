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
        Titles.source_ch12_directions: ChapterDirect,
        Titles.chapter_8_directions: ChapterDirect,
        Titles.network_ch12_directions: ChapterDirect,
    }


class Events(BaseMixin):
    def __init__(self, filename) -> None:
        wb = load_workbook(filename=filename, data_only=True)

        table_names = {}
        for worksheet in wb.sheetnames:
            for table in wb[worksheet].tables:
                table_names[table] = worksheet

        # import secondary tables
        source_unit_costs = self.__init_table(wb,
                                              table_names,
                                              Titles.source_unit_costs)
        tfu_unit_costs = self.__init_table(wb,
                                           table_names,
                                           Titles.tfu_unit_costs)
        network_unit_costs = self.__init_table(wb,
                                               table_names,
                                               Titles.network_unit_costs)
        ctp_unit_costs = self.__init_table(wb,
                                           table_names,
                                           Titles.ctp_unit_costs)
        stages = self.__init_table(wb,
                                   table_names,
                                   Titles.stages)
        deflator_indices = self.__init_table(wb,
                                             table_names,
                                             Titles.deflator_indices)
        terms = self.__init_table(wb, table_names, Titles.event_years)

        cum_deflator = 1
        for index in deflator_indices:
            if index.year >= terms[0].cost_year:
                cum_deflator *= index.index
                index.cumulative = cum_deflator

        nds = self.__init_table(wb, table_names, Titles.nds)

        chapter7_directions = self.__init_table(wb, table_names,
                                                Titles.chapter_7_directions)
        chapter8_directions = self.__init_table(wb, table_names,
                                                Titles.chapter_8_directions)
        source_ch12_directions = self.__init_table(
                                                wb,
                                                table_names,
                                                Titles.source_ch12_directions)
        network_ch12_directions = self.__init_table(
            wb, table_names, Titles.network_ch12_directions
        )

        # import and saving main tables
        self.source_events = self.__init_table(wb, table_names,
                                               Titles.source_events,
                                               stages, deflator_indices,
                                               terms, nds,
                                               source_unit_costs,
                                               tfu_unit_costs,)
        self.network_events = self.__init_table(wb, table_names,
                                                Titles.network_events,
                                                stages, deflator_indices,
                                                terms, nds,
                                                network_unit_costs,
                                                ctp_unit_costs,)
        # creating and saving summary tables
        self.chapter7_summary = self.__get_summary_tables(
                                                    chapter7_directions,
                                                    self.source_events,
                                                    'chapter_7_directions')
        self.source_ch12_sum = self.__get_summary_tables(
                                                    source_ch12_directions,
                                                    self.source_events,
                                                    'source_ch12_directions')
        self.chapter8_summary = self.__get_summary_tables(
                                                    chapter8_directions,
                                                    self.network_events,
                                                    'chapter_8_directions')
        self.network_ch12_sum = self.__get_summary_tables(
                                                    network_ch12_directions,
                                                    self.network_events,
                                                    'network_ch12_directions')

    def __init_table(self, wb, table_names, table_name, *args):
        worksheet = wb[table_names[table_name]]
        range = worksheet.tables[table_name].ref
        origin_data = worksheet[range][1:]
        import_data = []
        for row in origin_data:
            values = [cell.value for cell in row] + list(args)
            import_data.append(self.classes[table_name](*values))
        return import_data

    def __get_summary_tables(self, directions, events, direction_name):
        summary_tables = []
        for direct in directions:
            values = [
                direct.number, direct.name, events, False,
                direction_name
            ]
            summary_tables.append(SummaryTable(*values))
        values = [
            len(summary_tables) + 1, Titles.total, summary_tables, True,
            direction_name
        ]
        summary_tables.append(SummaryTable(*values))
        return summary_tables
