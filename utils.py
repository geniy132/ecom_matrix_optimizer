import datetime                                                       # Модуль для работы с датами и календарем
import re                                                             # Модуль регулярных выражений
import pandas as pd                                                   # Библиотека манипулирования таблицами


def clean_numeric_string(val):
    """Очищает ячейки от валют, процентов и преобразует в float."""
    if pd.isna(val):                                                  # Проверка на пустую ячейку Excel (NaN)
        return 0.0                                                    # Возвращаем дефолтное вещественное значение
    val_str = str(val).replace("%", "").replace(",", ".").strip()
    val_cleaned = re.sub(r"[^\d\.]", "", val_str)
    try:
        return (
            float(val_cleaned) if val_cleaned else 0.0
        )                                                             # Приведение к числу с плавающей точкой
    except ValueError:                                                # Защита от непредвиденного текста в ячейке
        return 0.0                                                    # Возврат нуля в случае системной ошибки


def find_table_header(file_path):
    """Ищет строку, в которой находится заголовок таблицы (шапка)."""
    # Читаем первые 50 строк файла целиком, чтобы точно не пропустить шапку
    preview = pd.read_excel(file_path, header=None, nrows=50)

    for idx, row in preview.iterrows():                               # Построчный обход выборки
        row_str = [str(s).lower().strip() for s in row.tolist()]      # Приводим всю строку к нижнему регистру

        # СТРОГАЯ ПРОВЕРКА: Игнорируем вводные тексты, требуя 4 маркера в одной строке
        has_name = any("название" in s or "товар" in s or "наименование" in s for s in row_str)
        has_sku = any("sku" in s or "артикул" in s for s in row_str)
        has_price = any("цена" in s for s in row_str)                 # Наличие слова "цена" в шапке
        has_margin = any("маржа" in s for s in row_str)               # Наличие слова "маржа" в шапке

        if has_name and has_sku and has_price and has_margin:         # Если все 4 условия совпали —
            return idx                                                # возвращаем индекс реальной шапки

    return 10                                                         # Фолбэк-значение по умолчанию


def parse_to_datetime(col_name):
    """Преобразует название колонки в datetime, если это дата (любого типа)."""
    # 1. Если Pandas уже прочитал заголовок как дату (Timestamp, datetime, date)
    if isinstance(col_name, (datetime.datetime, datetime.date, pd.Timestamp)):
        return pd.to_datetime(col_name)

    col_str = str(col_name).strip()

    # Игнорируем технические пустые столбцы Pandas
    if "unnamed" in col_str.lower() or col_str == "":
        return None

    # 2. Если имя колонки — это числовой код даты Excel (обычно от 40000 до 60000)
    if col_str.isdigit():
        num = int(col_str)
        if 40000 <= num <= 60000:
            try:
                return pd.to_datetime(num, unit='D', origin='1899-12-30')
            except Exception:
                pass

    # 3. Если это строка с датой, перебираем любые форматы
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y.%m.%d", "%d.%m.%y", "%d/%m/%y"):
        try:
            return pd.to_datetime(col_str, format=fmt)
        except (ValueError, TypeError):
            continue

    # 4. ФОЛБЭК: Если в названии колонки есть цифры (например, "01.12") и это не бизнес-столбец
    if any(char.isdigit() for char in col_str):
        try:
            # Пробуем распарсить стандартным гибким методом Pandas
            return pd.to_datetime(col_str, errors='coerce')
        except Exception:
            pass

    return None                                                       # Это текстовый бизнес-столбец (не дата)


def extract_date_columns(df, business_cols):
    """Ищет и возвращает список кортежей [(имя_колонки, объект_datetime)]."""
    date_list = []                                                    # Список для хранения валидных колонок-дат
    for col in df.columns:
        if col in business_cols:                                      # Если это точно столбец Названия, SKU, Цены или Маржи —
            continue                                                  # пропускаем его, это не динамика продаж

        dt = parse_to_datetime(col)                                   # Пробуем распарсить заголовок
        if dt is not None and not pd.isna(dt):                        # Если парсинг успешный и это не NaN —
            date_list.append((col, dt))                               # сохраняем пару (оригинальное имя, дата)

    # Сортируем список по времени, чтобы продажи шли строго хронологически
    date_list.sort(key=lambda x: x[1])
    return date_list                                                  # Возвращаем хронологический список дат
