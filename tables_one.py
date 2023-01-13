from dataclasses import fields

from base_datas import BaseEvent
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from styles import Styles
from summary_datas import SummaryTable
from titles import Titles
from widths import Formats, Widths


class BaseTable:

    def __init__(self, events: list[BaseEvent], sums: list[SummaryTable],
                 workbook: Workbook, name: str) -> None:
        self.events = events
        self.sums = sums
        self.wb = workbook
        self.styles = {
            Titles.base_style: Styles.style_1,
            Titles.title_style: Styles.style_4,
            Titles.header_style: Styles.style_3,
            Titles.footer_style: Styles.style_2,
        }
        self._add_style()
        self.name = name

    def _add_style(self) -> None:
        for title, style in self.styles.items():
            if title not in self.wb.named_styles:
                self.wb.add_named_style(style)

    def create_table(self) -> None:
        '''Create the table'''
        sheet_name = ''.join(w[0].upper() for w in self.name.split())
        if sheet_name not in self.wb.sheetnames:
            self.wb.create_sheet(title=sheet_name)
            self._create_title(sheet_name)
            self._create_header(sheet_name)
        self._create_content(sheet_name)
        self._create_footer(sheet_name)

    def _create_title(self, sheet_name):
        '''Create title of our table'''
        ws = self.wb[sheet_name]
        ws.cell(row=1, column=1).value = self.name
        ws.cell(row=1, column=1).style = Titles.title_style

    def _create_header(self, sheet_name: str) -> None:
        '''Create header of our table'''
        column = 1
        row = 3
        style = Titles.header_style
        height = 70
        ws = self.wb[sheet_name]
        obj = self.events[0]

        for attribute in fields(obj):
            title = attribute.name
            title_year = title[-4:]
            if title_year != str(Titles.begining_year):
                # Base header cell
                ws.merge_cells(
                    start_row=row, start_column=column,
                    end_row=row + 1, end_column=column
                )
                ws.cell(row=row, column=column).value = getattr(Titles, title)
                ws.cell(row=row, column=column).style = style
                ws.cell(row=row + 1, column=column).style = style
                self._set_column_width__(sheet_name, title, column)
                column += 1
            else:
                # We must merge first row
                first_year = Titles.begining_year
                last_year = Titles.ending_year
                period = last_year - first_year + 1
                ws.merge_cells(
                    start_row=row, start_column=column,
                    end_row=row, end_column=column + period
                )
                ws.cell(row=row, column=column).value = Titles.capex_flow
                # And draw head_capex_flow cells with dates in second row
                for year in range(first_year, last_year + 1):
                    ws.cell(row=row, column=column).style = style
                    ws.cell(row=row + 1, column=column).value = year
                    ws.cell(row=row + 1, column=column).style = style
                    self._set_column_width__(sheet_name, 'year_' + str(year),
                                             column)
                    column += 1
                ws.cell(row=row, column=column).style = style
                ws.cell(row=row + 1, column=column).style = style
                ws.cell(row=row + 1, column=column).value = Titles.total
                self._set_column_width__(sheet_name, 'total', column)
                break
        ws.row_dimensions[row + 1].height = height

    def _create_content(self, sheet_name: str):
        '''Create event data rows of our table'''
        column = 1
        row = 5
        style = Titles.base_style

        for event in self.events:
            column = 1
            self._set_row_data(event, sheet_name, row, column, style)
            row += 1

    def _create_footer(self, sheet_name: str):
        '''Create footer of our table'''
        column = 1
        row = len(self.events) + 5
        style = Titles.footer_style

        self._set_row_data(self.sums[-1], sheet_name, row, column, style)

    def _set_cell_value(self, sheet_name, row, column, value, style, title):
        ws = self.wb[sheet_name]
        ws.cell(row=row, column=column).value = value
        ws.cell(row=row, column=column).style = style
        form = getattr(Formats, title)
        ws.cell(row=row, column=column).number_format = form

    def _set_column_width__(self, sheet_name, title, column):
        ws = self.wb[sheet_name]
        width = getattr(Widths, title)
        ws.column_dimensions[get_column_letter(column)].width = width

    def _set_row_data(self, obj, sheet_name, row, column, style):
        for attribute in fields(self.events[0]):
            title = attribute.name
            if title != 'obj_type':
                value = getattr(obj, title, '')
                self._set_cell_value(sheet_name, row, column, value, style,
                                     title)
                column += 1


class Chapter7Table(BaseTable):
    def _create_content(self, sheet_name: str):
        self.events.sort(key=lambda x: x.chapter_7_directions, reverse=False)
        super()._create_content(sheet_name)

        column = 1
        style = Titles.footer_style
        row = 5
        for direction in self.sums[:-1]:
            ws = self.wb[sheet_name]
            ws.insert_rows(row)
            self._set_row_data(direction, sheet_name, row, column, style)
            row += direction.amount
