from events import Events
from openpyxl import load_workbook
from tables_one import (BaseTable, ByTSOTable, DirectionsTable,
                        DirectionsTableSum)
from titles import Titles


def main():
    filename = 'capex.xlsx'

    # Расчитываем и формируем данные используя базовый Excel
    events = Events(filename)

    # Открываем файл
    wb = load_workbook(filename=filename)
    '''
    # Вставляем таблицу по источникам в Excel
    table = BaseTable(wb, events.source_events, events.source_total,
                      events.terms)
    table.create_table()
    # Вставляем таблицу по тепловым сетям в Excel
    table = BaseTable(wb, events.network_events, events.source_total,
                      events.terms)
    table.create_table()

    # Вставляем таблицу по источникам в разере направлений в Excel
    table = DirectionsTable(wb, events.source_events, events.source_total,
                            events.terms, events.chapter7_summary)
    table.create_table()
    # Вставляем таблицу по тепловым сетям в разрезе направленийв в Excel
    table = DirectionsTable(wb, events.network_events, events.source_total,
                            events.terms, events.chapter8_summary)
    table.create_table()

    # Вставляем таблицу по источникам в разере направлений в Excel
    table = DirectionsTable(wb, events.source_events, events.source_total,
                            events.terms, events.source_ch12_sum)
    table.create_table()
    # Вставляем таблицу по тепловым сетям в разрезе направленийв в Excel
    table = DirectionsTable(wb, events.network_events, events.source_total,
                            events.terms, events.network_ch12_sum)
    table.create_table()

    # Вставляем таблицы по каждой ТСО по источникам в разрезе направлений
    # в Excel'''
    tso_list = events.get_tso_name_list()
    '''
    events_by_tso, totals, directions_by_tso = events.split_events_by_tso(
        'source', 'tso_name', Titles.chapter_7_directions
    )
    for tso_name, tso_events in events_by_tso.items():
        tso_short_name = tso_list.get(tso_name)
        table = ByTSOTable(wb, tso_events, totals[tso_name], events.terms,
                           directions_by_tso[tso_name], tso_name,
                           tso_short_name)
        table.create_table()

    # Вставляем таблицы по каждой ТСО по тепловым сетям
    # в разрезе направлений в Excel
    events_by_tso, totals, directions_by_tso = events.split_events_by_tso(
        'network', 'tso_name', Titles.chapter_8_directions
    )
    for tso_name, tso_events in events_by_tso.items():
        tso_short_name = tso_list.get(tso_name)
        table = ByTSOTable(wb, tso_events, totals[tso_name], events.terms,
                           directions_by_tso[tso_name], tso_name,
                           tso_short_name)
        table.create_table()'''

    # -------------------------------------------------------------------
    # Вставляем таблицы по каждой ТСО по источникам в разрезе направлений
    # в Excel
    (events_by_tso, totals, directions_by_tso, sum_events_by_tso, sum_totals,
     sum_directions_by_tso) = events.split_events_by_tso(
        'source', 'tso_name', Titles.source_ch12_directions, True
    )
    for tso_name, tso_events in events_by_tso.items():
        vars1 = (
            wb, tso_events, totals[tso_name], events.terms,
            directions_by_tso[tso_name]
        )
        vars2 = (
            sum_events_by_tso[tso_name], sum_totals[tso_name],
            sum_directions_by_tso[tso_name]
        )
        vars3 = (tso_name, tso_list.get(tso_name))
        table = DirectionsTableSum(*vars1, *vars2, *vars3)
        table.create_table()

    # -------------------------------------------------------------------
    # Вставляем таблицы по каждой ТСО по тепловым сетям
    # в разрезе направлений в Excel
    (events_by_tso, totals, directions_by_tso, sum_events_by_tso, sum_totals,
     sum_directions_by_tso) = events.split_events_by_tso(
        'network', 'tso_name', Titles.network_ch12_directions, True
    )
    for tso_name, tso_events in events_by_tso.items():
        vars1 = (
            wb, tso_events, totals[tso_name], events.terms,
            directions_by_tso[tso_name]
        )
        vars2 = (
            sum_events_by_tso[tso_name], sum_totals[tso_name],
            sum_directions_by_tso[tso_name]
        )
        vars3 = (tso_name, tso_list.get(tso_name))
        table = DirectionsTableSum(*vars1, *vars2, *vars3)
        table.create_table()

    # Сохраняем сделанное
    wb.save('new_' + filename)

    print('Сделано!')


if __name__ == '__main__':
    main()
