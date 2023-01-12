from dataclasses import fields

from base_datas import BaseEvent
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from styles import Styles
from summary_datas import SummaryTable
from titles import Titles
from widths import Formats, Widths


class TableMixin:

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
        self.__add_styles__()
        self.name = name

    def __add_styles__(self) -> None:
        for title, style in self.styles.items():
            if title not in self.wb.named_styles:
                self.wb.add_named_style(style)

    def create_table(self) -> None:
        '''Create the table'''
        sheet_name = ''.join(w[0].upper() for w in self.name.split())
        if sheet_name not in self.wb.sheetnames:
            self.wb.create_sheet(title=sheet_name)
            self.create_title(sheet_name)
            self.create_header(sheet_name)
        self.create_content(sheet_name)
        self.create_footer(sheet_name)

    def create_title(self, sheet_name):
        '''Create title of our table'''
        ws = self.wb[sheet_name]
        ws.cell(row=1, column=1).value = self.name
        ws.cell(row=1, column=1).style = Titles.title_style

    def create_header(self, sheet_name: str) -> None:
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
                self.__set_column_width__(sheet_name, title, column)
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
                    self.__set_column_width__(sheet_name, 'year_' + str(year),
                                              column)
                    column += 1
                ws.cell(row=row, column=column).style = style
                ws.cell(row=row + 1, column=column).style = style
                ws.cell(row=row + 1, column=column).value = Titles.total
                self.__set_column_width__(sheet_name, 'total', column)
                break
        ws.row_dimensions[row + 1].height = height

    def create_content(self, sheet_name: str):
        '''Create event data rows of our table'''
        column = 1
        row = 5
        style = Titles.base_style

        for event in self.events:
            column = 1
            self.__set_row_data__(event, sheet_name, row, column, style)
            row += 1

    def create_footer(self, sheet_name: str):
        '''Create footer of our table'''
        column = 1
        row = len(self.events) + 5
        style = Titles.footer_style

        self.__set_row_data__(self.sums[-1], sheet_name, row, column, style)

    def __set_cell_value__(self, sheet_name, row, column, value, style, title):
        ws = self.wb[sheet_name]
        ws.cell(row=row, column=column).value = value
        ws.cell(row=row, column=column).style = style
        form = getattr(Formats, title)
        ws.cell(row=row, column=column).number_format = form

    def __set_column_width__(self, sheet_name, title, column):
        ws = self.wb[sheet_name]
        width = getattr(Widths, title)
        ws.column_dimensions[get_column_letter(column)].width = width

    def __set_row_data__(self, obj, sheet_name, row, column, style):
        for attribute in fields(self.events[0]):
            title = attribute.name
            if title != 'obj_type':
                value = getattr(obj, title, '')
                self.__set_cell_value__(sheet_name, row, column, value, style,
                                        title)
                column += 1
