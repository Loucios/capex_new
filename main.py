from events import Events
from openpyxl import load_workbook
from tables_one import BaseTable, Chapter7Table, Chapter8Table
from titles import Titles


def main():
    filename = 'capex.xlsm'

    # Расчитываем и формируем данные используя базовый Excel
    events = Events(filename)

    # Вставляем таблицу по источникам в Excel
    wb = load_workbook(filename=filename, keep_vba=True)
    table = BaseTable(events.source_events,
                      events.chapter7_summary, wb, Titles.sources)
    table.create_table()
    table.wb.save('new_' + filename)

    # Вставляем таблицу по источникам в Excel
    wb = load_workbook(filename='new_' + filename, keep_vba=True)
    table = Chapter7Table(events.source_events,
                          events.chapter7_summary, wb,
                          'Источники тепловой энергии по направлениям')
    table.create_table()
    table.wb.save('new_' + filename)

    # Вставляем таблицу по тепловым сетям в Excel
    wb = load_workbook(filename='new_' + filename, keep_vba=True)
    table = BaseTable(events.network_events,
                      events.chapter8_summary, wb, Titles.networks)
    table.create_table()
    table.wb.save('new_' + filename)

    # Вставляем таблицу по тепловым сетям в Excel
    wb = load_workbook(filename='new_' + filename, keep_vba=True)
    table = Chapter8Table(events.network_events,
                          events.chapter8_summary, wb,
                          'Тепловые сети по нарпавлениям')
    table.create_table()
    table.wb.save('new_' + filename)


if __name__ == '__main__':
    main()
