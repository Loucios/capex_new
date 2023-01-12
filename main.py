from events import Events
from openpyxl import load_workbook
from tables_one import TableMixin
from titles import Titles


def main():
    filename = 'capex.xlsm'

    # Расчитываем и формируем данные используя базовый Excel
    events = Events(filename)

    # Вставляем таблицу по источникам в Excel
    wb = load_workbook(filename=filename, keep_vba=True)
    source_table = TableMixin(events.source_events,
                              events.chapter7_summary, wb, Titles.sources)
    source_table.create_table()
    source_table.wb.save('new_' + filename)

    # Вставляем таблицу по тепловым сетям в Excel
    wb = load_workbook(filename='new_' + filename, keep_vba=True)
    network_table = TableMixin(events.network_events,
                               events.chapter8_summary, wb, Titles.networks)
    network_table.create_table()
    network_table.wb.save('new_' + filename)


if __name__ == '__main__':
    main()
