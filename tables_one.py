import re
from dataclasses import fields

from base_datas import BaseEvent, Terms
from formats import (Chapter7Directions, Chapter8directions, NetworkTable,
                     SheetNames, SourceTable)
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from styles import Styles
from summary_datas import SummaryTable
from titles import Titles
from widths import Formats, Widths


class BaseTable:
    column = 1
    title_row = 1
    header_row = 3
    content_row = 5
    header_height = 70
    tables = {
        'source': SourceTable,
        'network': NetworkTable,
        'chapter_7_directions': Chapter7Directions,
        'chapter_8_directions': Chapter8directions,
    }
    styles = {
        Titles.base_style: Styles.style_1,
        Titles.title_style: Styles.style_4,
        Titles.header_style: Styles.style_3,
        Titles.footer_style: Styles.style_2,
        Titles.direction_style: Styles.style_5,
    }

    def __init__(self, workbook: Workbook, events: list[BaseEvent],
                 total: SummaryTable, terms: list[Terms],
                 sums: list[SummaryTable] = None, tso: str = None) -> None:
        self.wb = workbook
        self.events = events
        self.total = total
        self.start = terms[0].start_year
        self.end = terms[0].end_year
        self.sums = sums
        self.tso = tso
        self.table = self._get_table()
        self.name = self._get_name()
        self.sheet_name = self._get_sheet_name()

        self._add_style()
        self._sort_events()

    def _get_table(self):
        return self.tables[self.events[0].obj_type]

    def _get_name(self):
        return getattr(SheetNames, self.events[0].obj_type)

    def _sort_events(self):
        pass

    def _add_style(self) -> None:
        for title, style in self.styles.items():
            if title not in self.wb.named_styles:
                self.wb.add_named_style(style)

    def create_table(self) -> None:
        '''Create the table'''
        if self.sheet_name not in self.wb.sheetnames:
            self.wb.create_sheet(title=self.sheet_name)
            self._create_title()
            self._create_header()
        self._create_content()

    def _get_sheet_name(self):
        return ''.join(w[0].upper() for w in self.name.split())

    def _create_title(self):
        '''Create title of our table'''
        self._set_cell_value(
            self.title_row, self.column, self.name, Titles.title_style, 'title'
        )

    def _create_header(self) -> None:
        '''Create header of our table'''
        column = self.column
        row = self.header_row
        height = self.header_height
        style = Titles.header_style
        ws = self.wb[self.sheet_name]

        for attribute in fields(self.table):
            title = attribute.name
            if 'year_' not in title:
                # Base header cell
                ws.merge_cells(
                    start_row=row, start_column=column,
                    end_row=row + 1, end_column=column
                )
                ws.cell(row=row, column=column).value = getattr(Titles, title)
                ws.cell(row=row, column=column).style = style
                ws.cell(row=row + 1, column=column).style = style
                self._set_column_width(title, column)
                column += 1
            else:
                # We must merge first row
                period = self.end - self.start + 1
                ws.merge_cells(
                    start_row=row, start_column=column,
                    end_row=row, end_column=column + period
                )
                ws.cell(row=row, column=column).value = Titles.capex_flow
                # And draw head_capex_flow cells with dates in second row
                for year in range(self.start, self.end + 1):
                    ws.cell(row=row, column=column).style = style
                    ws.cell(row=row + 1, column=column).value = year
                    ws.cell(row=row + 1, column=column).style = style
                    self._set_column_width('year_' + str(year), column)
                    column += 1
                ws.cell(row=row, column=column).style = style
                ws.cell(row=row + 1, column=column).style = style
                ws.cell(row=row + 1, column=column).value = Titles.total
                self._set_column_width('total', column)
                break
        ws.row_dimensions[row + 1].height = height

    def _create_content(self):
        '''Create event data rows of our table'''
        row = self.content_row
        # content table
        for event in self.events:
            self._set_row_data(event, row, self.column, Titles.base_style)
            row += 1
        # table footer
        self._set_row_data(self.total, row, self.column, Titles.footer_style)

    def _set_cell_value(self, row, column, value, style, title):
        ws = self.wb[self.sheet_name]
        ws.cell(row=row, column=column).value = value
        ws.cell(row=row, column=column).style = style
        form = getattr(Formats, title)
        ws.cell(row=row, column=column).number_format = form

    def _set_column_width(self, title, column):
        ws = self.wb[self.sheet_name]
        width = getattr(Widths, title)
        ws.column_dimensions[get_column_letter(column)].width = width

    def _set_row_data(self, obj, row, column, style):
        for attribute in fields(self.table):
            title = attribute.name
            index = column
            chosen_style = style
            # Create table between start and end year
            if ('year_' not in title or (
                    self.start <= int(title[-4:]) <= self.end
               )):
                value = self._get_value(row, obj, title)
                # Merge cells around direction title
                if title == 'event_title' and isinstance(obj, SummaryTable):
                    self._set_merge(row, column)
                    chosen_style = Titles.direction_style
                    column = self.column + 1
                self._set_cell_value(row, column, value, chosen_style, title)
                column = index + 1
        # Set addition height with direction row
        if isinstance(obj, SummaryTable):
            self._set_height(row)

    def _set_height(self, row):
        ws = self.wb[self.sheet_name]
        ws.row_dimensions[row].height = self.header_height

    def _set_merge(self, row, column):
        ws = self.wb[self.sheet_name]
        ws.merge_cells(start_row=row, start_column=self.column + 1,
                       end_row=row, end_column=column)

    def _get_value(self, row, obj, title):
        if title == 'number':
            return self._get_number(row, obj)
        else:
            return getattr(obj, title, '')

    def _get_number(self, row, obj):
        value = row - self.content_row + 1
        if self.sums is not None:
            if isinstance(obj, SummaryTable):
                if value == len(self.events) + 1:
                    return len(self.sums) + 1
                return obj.number
            left = 0
            for index, direction in enumerate(self.sums):
                if index:
                    left += self.sums[index - 1].amount
                right = left + direction.amount
                if left < value <= right:
                    number = value - left
                    return f'{direction.number}.{number}'
        return value


class DirectionsTable(BaseTable):
    def _get_table(self):
        return self.tables[self.sums[0].directions]

    def _get_name(self):
        return getattr(SheetNames, self.sums[0].directions)

    def _sort_events(self):
        directions = self.sums[0].directions
        self.events.sort(
            key=lambda x: getattr(x, directions), reverse=False
        )

    def _create_directions(self) -> None:
        '''Insert to existing table "Summary by directions" rows'''

        column = self.column
        style = Titles.footer_style
        row = self.content_row
        sheet_name = self.sheet_name
        ws = self.wb[sheet_name]
        index = 0
        for direction in self.sums:
            # Calculate the amount rows to shift merged cells
            index += 1

            ws.insert_rows(row)
            self._set_row_data(direction, row, column, style)
            row += (direction.amount + 1)
        # We shift merged cells after inserting rows because
        # inserting doesn't affect the merged cells
        # row = row - direction.amount - 1
        # self._shift_merge_cells(index)

    def _shift_merge_cells(self, amount: int) -> None:
        '''Shift the merged cells

        Shift the merged cells wich place after "row" by "amount" rows
        in "sheet_name" list
        '''
        ws = self.wb[self.sheet_name]
        merged_cells_range = ws.merged_cells.ranges
        for merged_cell in merged_cells_range:
            if merged_cell.min_row == len(self.events) + self.column + 1:
                print(len(self.events) + self.column + 1)
                print(self.sheet_name)
                merged_cell.shift(0, amount)

    def _create_content(self):
        super()._create_content()
        self._create_directions()


class ByTSOTable(DirectionsTable):

    def _get_sheet_name(self):
        sheet_name = super()._get_sheet_name()
        short_name = self._get_short_name()
        return f'{sheet_name} {short_name}'

    def _get_short_name(self):
        pattern = re.compile(r'[^\w ]+')
        return re.sub(pattern, '', self.tso)[:10]

    def _get_name(self):
        name = super()._get_name()
        return f'{name} {self.tso}'
