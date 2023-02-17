from events import Events
from openpyxl import load_workbook
from tables_one import BaseTable, ByTSOTable, DirectionsTable


def main():
    filename = 'capex.xlsm'

    # Расчитываем и формируем данные используя базовый Excel
    events = Events(filename)

    # Открываем файл
    wb = load_workbook(filename=filename, keep_vba=True)

    # Вставляем таблицу по источникам в Excel
    table = BaseTable(wb, events.source_events, events.source_total,
                      events.terms)
    table.create_table()
    # Вставляем таблицу по тепловым сетям в Excel
    table = BaseTable(wb, events.network_events, events.source_total,
                      events.terms)
    table.create_table()

    # Вставляем таблицу по источникам в Excel
    table = DirectionsTable(wb, events.source_events, events.source_total,
                            events.terms, events.chapter7_summary)
    table.create_table()
    # Вставляем таблицу по тепловым сетям в Excel
    table = DirectionsTable(wb, events.network_events, events.source_total,
                            events.terms, events.chapter8_summary)
    table.create_table()

    events_by_tso = events.split_events_by_tso('source', 'tso_name')
    for tso_name, tso_events in events_by_tso.items():
        table = ByTSOTable(wb, tso_events, events.source_total,
                           events.terms, events.chapter7_summary, tso_name)
        table.create_table()

    events_by_tso = events.split_events_by_tso('network', 'tso_name')
    for tso_name, tso_events in events_by_tso.items():
        table = ByTSOTable(wb, tso_events, events.network_total,
                           events.terms, events.chapter8_summary, tso_name)
        table.create_table()

    # Сохраняем сделанное
    wb.save('new_' + filename)


if __name__ == '__main__':
    main()
