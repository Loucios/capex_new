from dataclasses import dataclass


@dataclass(frozen=True)
class Titles:
    base_direction = 'Основное направление'
    base_style = 'base_style'
    begining_year = 2020
    capex_flow = 'Капитальные затраты в прогнозных ценах с учетом НДС, млн руб.'
    chapter_12_directions = 'Направления мероприятий из Главы 12'
    chapter_8_directions = 'СтруктураГлава8'
    ctp_unit_costs = 'ЦТПУдельники'
    const_indicators = 'Строительство'
    deflator = 'Индекс'
    deflator_indices = 'Индексы'
    deflators = 'Индексы'
    design = 'ПИР'
    diameter = 'Диаметр, мм'
    district = 'Административный район'
    ending_year = 2050
    eto_name = 'Наименование EТО'
    event_title = 'Наименование мероприятия'
    filename = 'capex.xlsm'
    first_year = 'Год начала'
    footer_style = 'footer_style'
    gh = 'Гкал/ч'
    header_style = 'header_style'
    inv_pro = 'ИП'
    last_year = 'Год окончания'
    laying_type = 'Типа прокладки'
    length = 'Протяженность, м'
    mw = 'МВт'
    nds = 'НДС'
    network_events = 'МероприятияСети'
    network_unit_costs = 'УдельникиСети'
    number = '№ п/п'
    power = 'Гкал/ч'
    power_range = 'Диапазон мощности'
    reconst2_indicators = 'Реконструкция2'
    reconst_indicators = 'Реконструкция'
    source = 'Наименование источника'
    source_events = 'МероприятияИсточники'
    source_id = 'id источника'
    source_name = 'Источники тепловой энергии'
    source_type = 'ТЭЦ/ котельная'
    source_unit_costs = 'УдельникиИсточники'
    stages = 'Этапы'
    terms = 'Сроки'
    tfu_unit_costs = 'ТФУ'
    th = 'т/ч'
    title_style = 'title_style'
    total = 'Итого'
    total_cost = 'Общая стоимость мероприятий, млн руб. без НДС'
    tso_list = 'СписокТСО'
    tso_name = 'Наименование ТСО'
    unit_type = 'Тип мероприятия'
    year = 'Год'
    year_cost = 'Цены, год'
    chanal = 'канальная'
    chanalless = 'бесканальная'
    elevated = 'надземная'
    basement = 'подвальная'
    casing = 'футляр'
    mixed = 'смешанная'
