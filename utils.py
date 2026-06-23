import datetime                                                       # Модуль для работы с датами и календарем
import re                                                             # Модуль регулярных выражений для очистки
import pandas as pd                                                   # Библиотека манипулирования таблицами


def clean_numeric_string(val):
    """Очищает ячейки от валют, процентов и преобразует в float."""
    if pd.isna(val):                                                  # Проверка на пустую ячейку Excel (NaN)
        return 0.0                                                    # Возвращаем дефолтное вещественное значение
    val_str = str(val).replace("%", "").replace(",", ".").strip()     # Удаляем проценты и нормализуем точки
    val_cleaned = re.sub(r"[^\d\.]", "", val_str)                     # Очищаем регулярным выражением текст ячейки
    try:                                                              # Защита от непредвиденного текста ячеек
        return float(val_cleaned) if val_cleaned else 0.0             # Приведение очищенной строки к float
    except ValueError:                                                # В случае системного сбоя конвертации
        return 0.0                                                    # возвращаем страховой дефолтный нуль


def find_table_header_in_sheet(df_sheet):
    """Ищет строку, в которой находится заголовок таблицы (шапка)."""
    preview = df_sheet.head(50)                                       # Сканируем первые 50 строк переданного листа
    for idx, row in preview.iterrows():                               # Построчный обход выборки переданного листа
        row_str = [str(s).lower().strip() for s in row.tolist()]      # Приводим всю строку к нижнему регистру
        has_name = any(x in s for x in ["название", "товар", "наименование"] for s in row_str) # Маркер Названия
        has_sku = any(x in s for x in ["sku", "артикул"] for s in row_str) # Маркер SKU / Артикула номенклатуры
        has_price = any("цена" in s for s in row_str)                 # Ключевой фильтр: наличие слова "цена"
        has_margin = any("маржа" in s for s in row_str)               # Ключевой фильтр: наличие слова "маржа"

        if has_name and has_sku and has_price and has_margin:         # Шапка валидна только при совпадении всех 4
            return idx                                                # Возвращаем точный индекс реальной шапки
    return None                                                       # Если лист не является базовой матрицей


def parse_to_datetime(col_name):
    """Преобразует название колонки в datetime, если это дата."""
    if isinstance(col_name, (datetime.datetime, datetime.date, pd.Timestamp)): # Если Pandas уже прочитал как дату
        return pd.to_datetime(col_name)                               # Возвращаем готовый объект Timestamp
    col_str = str(col_name).strip()                                   # Срезаем боковые пробелы текста заголовка
    if "unnamed" in col_str.lower() or col_str == "":                 # Игнорируем технические пустые столбцы
        return None                                                   # Возвращаем пустой маркер фильтрации
    if col_str.isdigit():                                             # Если имя колонки — это код даты Excel
        num = int(col_str)                                            # Преобразуем строковый код в целое число
        if 40000 <= num <= 60000:                                     # Фильтр диапазона кодов дат e-com периода
            try:                                                      # Защита от системных ошибок конвертации
                return pd.to_datetime(num, unit='D', origin='1899-12-30') # Конвертируем код Excel в дату
            except Exception:                                         # В случае непредвиденного сбоя ядра
                pass                                                  # пропускаем шаг обработки кода даты
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y.%m.%d", "%d.%m.%y", "%d/%m/%y"): # Массив e-com масок дат
        try:                                                          # Попытка парсинга по строгой маске
            return pd.to_datetime(col_str, format=fmt)                # Возвращаем валидный объект даты/времени
        except (ValueError, TypeError):                               # Если маска не подошла — продолжаем
            continue                                                  # Переходим к следующему элементу цикла
    if any(char.isdigit() for char in col_str):                       # Фолбэк-проверка: если в строке есть цифры
        try:                                                          # Попытка гибкого разбора силами Pandas
            return pd.to_datetime(col_str, errors='coerce')           # Возвращаем объект или безопасный NaT
        except Exception:                                             # Если гибкий разбор провалился —
            pass                                                      # игнорируем исключение
    return None                                                       # Это текстовый бизнес-столбец (не дата)


def extract_date_columns(df, business_cols):
    """Ищет и возвращает список кортежей [(имя_колонки, объект_datetime)]."""
    date_list = []                                                    # Список для хранения валидных micro-дат
    for col in df.columns:                                            # Перебор колонок текущего датафрейма
        if col in business_cols:                                      # Если это точно коммерческий столбец —
            continue                                                  # пропускаем его, это не динамика продаж
        dt = parse_to_datetime(col)                                   # Пробуем распарсить текущий заголовок
        if dt is not None and not pd.isna(dt):                        # Если парсинг успешный и это не NaN —
            date_list.append((col, dt))                               # сохраняем пару (оригинальное имя, дата)
    date_list.sort(key=lambda x: x[1])                                # Сортируем список хронологически по датам
    return date_list                                                  # Возвращаем упорядоченный список пар дат
