import calendar                                                       # Модуль для вычисления дней в месяцах
import json                                                           # Модуль для чтения файла конфигурации


def load_business_config():
    """Загружает параметры маржи, лимиты скорости и тексты из JSON."""
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:                                                 # Фолбэк-заглушка на случай удаления файла
        return {
            "MARGIN_THRESHOLDS": {"HIGH": 15.0, "LOW": 12.0},
            "SPEED_MULTIPLIERS": {"DRIVER_FACTOR": 1.5},
            "RECOMMENDATIONS": {
                "DRIVER": "Драйвер прибыли.", "STAR": "Потенциальная звезда.",
                "TURNOVER": "Оборотный товар.", "WASTE": "Балласт.",
                "STABLE": "Стабильная позиция."
            }
        }


def get_next_3_months_info(last_date):
    """Вычисляет названия и количество дней для 3 будущих месяцев."""
    months_info = []                                                  # Список кортежей (название, дней)
    current_year = last_date.year                                     # Извлекаем текущий год данных
    current_month = last_date.month                                   # Извлекаем текущий месяц данных

    ru_months = {
        1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
        5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
        9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }

    for _ in range(3):                                                # Цикл генерации на 3 месяца вперед
        current_month += 1                                            # Переходим к следующему месяцу
        if current_month > 12:                                        # Если вышли за рамки декабря —
            current_month = 1                                         # сбрасываем на январь
            current_year += 1                                         # увеличиваем счетчик года
        
        _, days_in_month = calendar.monthrange(current_year, current_month)
        label = f"Прогноз {ru_months[current_month]}, шт"             # Формируем красивое имя колонки
        months_info.append((label, days_in_month))                    # Добавляем данные в список
        
    return months_info                                                # Возвращаем метаданные для прогноза


def calculate_demand_forecast(df, date_cols, count_days, last_date):
    """ФОРМУЛА 1: Расчет спроса, скорости продаж и прогнозов закупки."""
    df["Всего_продано_шт"] = df[date_cols].sum(axis=1)                 # ФОРМУЛА: Горизонтальная сумма строк
    df["Скорость_продаж_день"] = df["Всего_продано_шт"] / count_days   # ФОРМУЛА: Общие продажи / Кол-во дней

    next_months = get_next_3_months_info(last_date)                   # Получаем метаданные будущих месяцев

    for col_label, days in next_months:                               # ФОРМУЛА 2: Чистый расчет без ошибок :=
        df[col_label] = (df["Скорость_продаж_день"] * days).round().astype(int)

    return df, next_months                                            # Возвращаем датафрейм и новые заголовки


def calculate_marginal_profit(df):
    """ФОРМУЛА 3: Расчет маржинальной прибыли товара за весь период."""
    # ФОРМУЛА: (Цена розничная * Доля маржи) * Общее количество проданных штук
    df["Маржинальная прибыль за период, руб"] = (
        (df["_calc_price"] * df["_calc_margin_pct"]) * df["Всего_продано_шт"]
    ).round(2)
    return df                                                         # Возврат фрейма с расчетной прибылью


def generate_business_recommendations(df):
    """ФОРМУЛА 4: Расчет медианы матрицы и применение вердиктов из конфига."""
    median_speed = df["Скорость_продаж_день"].median()                # ФОРМУЛА: Медиана скорости спроса матрицы
    if median_speed == 0:                                              # Защита от деления на ноль
        median_speed = 1.0                                             # Страховочное присвоение единицы

    cfg = load_business_config()                                      # Загружаем внешние пороги и тексты из JSON

    def get_recommendation(row):
        m_pct = row["_calc_margin_pct"] * 100                         # Перевод маржи в проценты
        spd = row["Скорость_продаж_день"]                              # Темп спроса на один SKU

        high_m = cfg["MARGIN_THRESHOLDS"]["HIGH"]                     # Извлекаем верхний порог маржи
        low_m = cfg["MARGIN_THRESHOLDS"]["LOW"]                       # Извлекаем нижний порог маржи
        factor = cfg["SPEED_MULTIPLIERS"]["DRIVER_FACTOR"]            # Извлекаем коэффициент скорости

        recs = cfg["RECOMMENDATIONS"]                                 # Извлекаем словарь текстовых шаблонов

        if m_pct >= high_m and spd >= median_speed * factor:          # Логика деления на Драйверы
            return recs["DRIVER"]
        elif m_pct >= high_m and spd < median_speed:                  # Логика деления на Потенциальные звезды
            return recs["STAR"]
        elif m_pct < low_m and spd >= median_speed:                   # Логика деления на Оборотные SKU
            return recs["TURNOVER"]
        elif m_pct < low_m and spd < median_speed:                    # Логика деления на Балласт
            return recs["WASTE"]
        else:                                                         # Логика деления на Стабильные позиции
            return recs["STABLE"]

    df["Рекомендация по оптимизации"] = df.apply(
        get_recommendation, axis=1
    )                                                                 # Построчный запуск функции матрицы
    return df                                                         # Возврат фрейма с выводами
