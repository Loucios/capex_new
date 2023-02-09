from events import Events
from openpyxl import load_workbook
from tables_one import BaseTable, DirectionsTable


def main():
    filename = 'capex.xlsm'

    # Расчитываем и формируем данные используя базовый Excel
    events = Events(filename)

    # Открываем файл
    wb = load_workbook(filename=filename, keep_vba=True)

    # Вставляем таблицу по источникам в Excel
    table = BaseTable(events.source_events, wb, events.terms)
    table.create_table()
    # Вставляем таблицу по тепловым сетям в Excel
    table = BaseTable(events.network_events, wb, events.terms)
    table.create_table()

    # Вставляем таблицу по источникам в Excel
    table = DirectionsTable(events.source_events, wb, events.terms,
                            events.chapter7_summary)
    table.create_table()
    # Вставляем таблицу по тепловым сетям в Excel
    table = DirectionsTable(events.network_events, wb, events.terms,
                            events.chapter8_summary)
    table.create_table()

    # Сохраняем сделанное
    wb.save('new_' + filename)


if __name__ == '__main__':
    main()
