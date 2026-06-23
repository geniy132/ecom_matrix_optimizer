from utils import find_table_header_in_sheet                          # Поиск шапки матрицы в листах


def scan_and_load_matrix(xl_file, choice):
    """Автоматически находит и загружает целевой лист книги Excel."""
    target_sheet = None                                               # Переменная для имени целевого листа
    header_idx = 0                                                    # Индекс строки заголовков шапки

    if choice in ["2", "3"]:                                          # Если запущен расчет Задания 2 или 3 (Промо)
        for sheet_name in xl_file.sheet_names:                        # Сканируем листы книги Excel по очереди
            df_tmp = xl_file.parse(sheet_name, nrows=5)                # Экспресс-чтение первых 5 строк данных
            df_tmp.columns = [str(c).lower().strip() for c in df_tmp.columns] # Нормализация имен шапки листа

            # Проверяем маркеры журнала транзакций
            has_revenue = any("выруч" in c for c in df_tmp.columns)
            has_quantity = any("колич" in c for c in df_tmp.columns)
            has_sku = any("sku" in c or "артикул" in c for c in df_tmp.columns)

            if has_revenue and has_quantity and has_sku:              # Если все 3 маркера лога совпали —
                df_full_check = xl_file.parse(sheet_name)             # загружаем лист целиком для проверки
                if len(df_full_check) > 20:                           # Фильтр объема: отсекаем Лист 2 ТЗ
                    target_sheet = sheet_name                         # Задаем Лист 3 как базовый источник
                    header_idx = 0                                    # Шапка лога транзакций всегда на 0 строке
                    print(f"🔍 Журнал транзакций найден на листе: '{sheet_name}' (Строк: {len(df_full_check)})")
                    break                                             # Прерываем цикл поиска лога транзакций
    else:
        # СЦЕНАРИЙ ОПЕРАЦИИ 1: Ищем классическую матрицу продаж
        for sheet_name in xl_file.sheet_names:                        # Сканируем листы книги Excel по очереди
            df_check = xl_file.parse(sheet_name, header=None, nrows=30) # Читаем первые 30 строк для проверки
            idx = find_table_header_in_sheet(df_check)                # Валидируем наличие 4 ИТ-маркеров матрицы
            if idx is not None:                                       # Если шапка успешно обнаружена —
                target_sheet = sheet_name                             # фиксируем имя целевого листа матрицы
                header_idx = idx                                      # сохраняем точный индекс строки
                print(f"🔍 Матрица ассортимента найдена на листе: '{sheet_name}'")
                break                                                 # Прерываем глобальный перебор листов

    if not target_sheet:                                              # Если ни один лист не прошел фильтры —
        return None, None                                             # возвращаем пустые маркеры сбоя наверх

    df = xl_file.parse(target_sheet, header=header_idx)               # Загружаем выбранный лист данных в Pandas
    df.columns = [str(c).strip() for c in df.columns]                 # Срез пробелов в именах колонок шапки
    return df, target_sheet                                           # Возвращаем датафрейм и имя листа книги


def identify_business_columns(df):
    """Интеллектуально распознает имена колонок по синонимам бизнеса."""
    # Поиск столбца наименований с приоритетом на "Полное название"
    name_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["полное название", "полное наименование"])), None)
    if not name_col:
        name_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["название", "товар", "наименование"])), None)

    # Поиск остальных коммерческих векторов матрицы
    sku_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["sku", "артикул"])), None)
    price_col = next((c for c in df.columns if "цена" in str(c).lower()), None)
    margin_col = next((c for c in df.columns if "маржа" in str(c).lower()), None)

    return name_col, sku_col, price_col, margin_col                  # Возвращаем кортеж имен столбцов листа
