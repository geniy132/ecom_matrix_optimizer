import calendar                                                       # Модуль для вычисления дней в месяцах
from config_loader import load_business_config                        # Импорт общего загрузчика конфига JSON


def get_next_3_months_info(last_date):
    """Вычисляет названия и количество дней для 3 будущих месяцев."""
    months_info = []                                                  # Список кортежей (название, дней)
    current_year = last_date.year                                     # Извлекаем текущий год данных матрицы
    current_month = last_date.month                                   # Извлекаем текущий месяц данных матрицы

    ru_months = {                                                     # Справочник имен месяцев на русском
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",            # Срез 1-4 месяцев года
        5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",                  # Срез 5-8 месяцев года
        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"     # Срез 9-12 месяцев года
    }                                                                 # Конец структуры справочника имен

    for _ in range(3):                                                # Цикл генерации на 3 месяца вперед
        current_month += 1                                            # Переходим к следующему календарному месяцу
        if current_month > 12:                                        # Если вышли за рамки декабря —
            current_month = 1                                         # сбрасываем счетчик на январь
            current_year += 1                                         # увеличиваем счетчик текущего года
        
        _, days_in_month = calendar.monthrange(current_year, current_month) # Находим длину месяца в днях
        label = f"Прогноз {ru_months[current_month]}, шт"             # Формируем красивое имя колонки отчета
        months_info.append((label, days_in_month))                    # Добавляем данные в список метаданных
    return months_info                                                # Возвращаем метаданные для прогноза


class DemandForecastOperation:
    """Операция 1: Прогнозирование товарного спроса и расчет маржинальности."""
    def execute(self, df, xl_file, context):                          # Главный метод запуска расчетов формул
        date_cols = context["date_cols"]                              # Извлекаем список столбцов дат продаж
        count_days = context["count_days"]                            # Извлекаем общую длину периода продаж
        last_date = context["last_date"]                              # Извлекаем крайнюю дату временного ряда

        df["Всего_продано_шт"] = df[date_cols].sum(axis=1)             # ФОРМУЛА 1: Расчет суммы продаж по строке
        df["Скорость_продаж_день"] = df["Всего_продано_шт"] / count_days # ФОРМУЛА 1: Расчет посуточного темпа спроса

        next_months = get_next_3_months_info(last_date)               # Генерируем календарную сетку месяцев весна
        for col_label, days in next_months:                           # Перебор месяцев прогнозирования потребности
            df[col_label] = (df["Скорость_продаж_день"] * days).round().astype(int) # ФОРМУЛА 2: Объем потребности закупки

        df["Маржинальная прибыль за период, руб"] = (                 # ФОРМУЛА 3: Расчет прибыли за период
            (df["_calc_price"] * df["_calc_margin_pct"]) * df["Всего_продано_шт"]
        ).round(2)                                                    # Округление прибыли до копеек float

        median_speed = df["Скорость_продаж_день"].median()            # Находим медиану скорости по всей матрице
        if median_speed == 0:                                         # Страховка от деления на ноль матрицы
            median_speed = 1.0                                        # Принудительный страховочный дефолт

        cfg = load_business_config()                                  # Читаем пороговые лимиты и правила из JSON
        
        def get_recommendation(row):                                  # Вложенный метод построчных вердиктов робота
            m_pct = row["_calc_margin_pct"] * 100                     # Переводим долю маржи в проценты
            spd = row["Скорость_продаж_день"]                         # Извлекаем скорость текущего SKU матрицы
            high_m = cfg["MARGIN_THRESHOLDS"]["HIGH"]                 # Фиксируем верхнюю планку маржи бизнеса
            low_m = cfg["MARGIN_THRESHOLDS"]["LOW"]                   # Фиксируем нижнюю планку маржи бизнеса
            factor = cfg["SPEED_MULTIPLIERS"]["DRIVER_FACTOR"]        # Коэффициент успешности темпа продаж
            recs = cfg["RECOMMENDATIONS"]                             # Шаблоны выводов отчета из файла настроек

            if m_pct >= high_m and spd >= median_speed * factor:      # Фильтр для категории Драйверов прибыли
                return recs["DRIVER"]                                 # Вывод шаблона драйвера из JSON
            elif m_pct >= high_m and spd < median_speed:              # Фильтр для скрытых Звезд карточек e-com
                return recs["STAR"]                                   # Вывод шаблона звезды из JSON
            elif m_pct < low_m and spd >= median_speed:               # Фильтр для Оборотных позиций матрицы
                return recs["TURNOVER"]                               # Вывод шаблона оборотного SKU из JSON
            elif m_pct < low_m and spd < median_speed:                # Фильтр для балластовой группы матрицы
                return recs["WASTE"]                                  # Вывод шаблона балласта из JSON
            else:                                                     # Для стабильной группы товаров матрицы
                return recs["STABLE"]                                 # Вывод дефолтного вердикта из JSON

        df["Рекомендация по оптимизации"] = df.apply(get_recommendation, axis=1) # Запуск построчной разметки матрицы
        context["forecast_cols"] = next_months                        # Сохраняем новые колонки в контекст
        return df                                                     # Возвращаем рассчитанный датафрейм
