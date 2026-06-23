import pandas as pd                                                   # Библиотека манипулирования таблицами


class PromoEffectivenessOperation:
    """Операция 4: Динамический анализ промо-акций и рисков каннибализации."""
    def execute(self, df, xl_file, context):                          # Главный метод запуска команды
        # 1. ЖЕЛЕЗОБЕТОННЫЙ ВЕКТОРНЫЙ ПОИСК И ОЧИСТКА ЛИСТА ТРАНЗАКЦИЙ
        target_sheet = next((s for s in xl_file.sheet_names if "лист3" in s.lower() or "лог" in s.lower() or "транз" in s.lower()), None)
        if not target_sheet:                                          # Фолбэк: если по имени не нашли, ищем по структуре
            for s in xl_file.sheet_names:                             # Сканируем все листы книги Excel по очереди
                df_test = xl_file.parse(s, nrows=5)                   # Читаем первые 5 строк для проверки полей
                if any("выруч" in str(c).lower() for c in df_test.columns) and any("колич" in str(c).lower() for c in df_test.columns):
                    target_sheet = s                                  # Фиксируем валидное имя листа лога
                    break                                             # Выходим из цикла сканирования книги

        if not target_sheet: return df                                # Защита от отсутствия листа продаж

        try:                                                          # Блок защиты расчетов промо-акции
            df_raw = xl_file.parse(target_sheet, header=None)         # Загружаем лист сырым массивом без структуры
            df_raw = df_raw.dropna(how='all').reset_index(drop=True)  # Удаляем полностью пустые строки из Excel

            # Находим индекс строки, которая является реальной шапкой таблицы
            h_idx = 0                                                 # Дефолтный индекс строки заголовков
            for idx, row in df_raw.iterrows():                        # Перебираем строки сырого массива данных
                # Безопасно переводим каждую ячейку в строку, полностью защищаясь от объектов Series
                row_strings = [str(cell).lower().strip() for cell in row.tolist() if not pd.isna(cell) and not isinstance(cell, (list, tuple, dict, pd.Series))]
                # Если в строке есть пересечение базовых ИТ-маркеров e-com аналитики — это шапка!
                if any("sku" in c or "артикул" in c for c in row_strings) and any("продаж" in c or "колич" in c or "выруч" in c for c in row_strings):
                    h_idx = idx                                       # Фиксируем точный индекс строки шапки
                    break                                             # Выходим из сканирования строк превью

            # Пересобираем датафрейм, назначив найденную строку в качестве официальной шапки
            df_s3 = xl_file.parse(target_sheet, header=h_idx)
            df_s3.columns = [str(c).strip() for c in df_s3.columns]   # Очистка имен столбцов от пробелов

            # Локализация коммерческих столбцов по мягкому текстовому вхождению
            s3_date = next((c for c in df_s3.columns if "дата" in str(c).lower() or "date" in str(c).lower() or "время" in str(c).lower()), df_s3.columns)
            s3_sku = next((c for c in df_s3.columns if "sku" in str(c).lower() or "артикул" in str(c).lower()), df_s3.columns if len(df_s3.columns) > 1 else df_s3.columns)
            s3_qty = next((c for c in df_s3.columns if "колич" in str(c).lower() or "кол-во" in str(c).lower() or "шт" in str(c).lower() or "продаж" in str(c).lower()), None)
            s3_rev = next((c for c in df_s3.columns if "выруч" in str(c).lower() or "сумма" in str(c).lower() or "руб" in str(c).lower() or "оборот" in str(c).lower()), None)

            if not all([s3_date, s3_sku, s3_qty, s3_rev]):            # Если критическая колонка потерялась —
                print("❌ Ошибка: Не удалось сопоставить изменившиеся столбцы на Листе 3.") # Системное сообщение
                return df                                             # Прерывание выполнения операции

            # Адаптивный гибридный парсер дат с защитой от американского формата и скрытых объектов
            def robust_hybrid_date_parser(val):
                if pd.isna(val) or isinstance(val, (list, tuple, dict, pd.Series)): 
                    return pd.NaT                                     # Возвращаем пустую дату при конфликтах типов
                try: return pd.to_datetime(val, errors='raise')       # Попытка 1: стандартный автопарсер
                except Exception: pass                                # Если сбой — переходим к ручному разбору

                s_val = str(val).strip()                              # Принудительно приводим значение к тексту
                if s_val.replace('.', '').isdigit() and len(s_val) < 7: # Защита от внутренних чисел Excel
                    try: return pd.to_datetime(float(s_val), unit='D', origin='1899-12-30')
                    except: return pd.NaT                             # Игнорируем битые числовые ячейки

                separator = '-' if '-' in s_val else '/' if '/' in s_val else None # Ищем разделитель даты
                if not separator: return pd.NaT                       # Если разделитель отсутствует — мусор

                parts = s_val.split(separator)                        # Расщепляем строку на компоненты времени
                if len(parts) == 3:                                   # Если получили 3 элемента (день/месяц/год)
                    try:
                        if separator == '-':                          # Если дефис — формат уже ISO (ГГГГ-ММ-ДД)
                            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                        else:                                         # Если слэш — американский (ММ/ДД/ГГГГ)
                            m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
                        return pd.Timestamp(year=y, month=m, day=d)   # Собираем эталонный Timestamp Pandas
                    except Exception: return pd.NaT                   # Защита от некорректных символов
                return pd.NaT                                         # Нарушена структура даты

            df_s3[s3_date] = df_s3[s3_date].apply(robust_hybrid_date_parser) # Применяем бронированный парсер
            df_s3 = df_s3.dropna(subset=[s3_date]).copy()             # Удаляем строки, где дата не распозналась

            min_a = df_s3[s3_date].min().strftime('%Y-%m-%d')         # Вычисляем минимальную дату в логе
            max_a = df_s3[s3_date].max().strftime('%Y-%m-%d')         # Вычисляем максимальную дату в логе
            print(f"\n📊 Диапазон дат в логе: с {min_a} по {max_a}")  # Подсказка доступных интервалов

            # Лаконичный ввод периодов в консоли (с клавиатуры)
            print("\n📅 Настройка периодов (Вводите даты строго в формате ГГГГ-ММ-ДД):")
            pre_start = pd.to_datetime(input("👉 Дата НАЧАЛА базового периода ДО: ").strip())
            pre_end = pd.to_datetime(input("👉 Дата ОКОНЧАНИЯ базового периода ДО: ").strip())
            promo_start = pd.to_datetime(input("👉 Дата НАЧАЛА промо-акции: ").strip())
            promo_end = pd.to_datetime(input("👉 Дата ОКОНЧАНИЯ промо-акции: ").strip())

            print("\n👉 Вставьте список промо-SKU из Excel (ПКМ) и нажмите Enter ДВАЖДЫ:")
            target_skus = []                                          # Буфер вставленных менеджером SKU
            while True:                                               # Цикл чтения строк буфера консоли
                line = input().strip()                                # Считываем текущую строку ввода
                if not line: break                                    # Нажатие на пустой Enter завершает ввод
                import re                                             # Локальный импорт регулярных выражений
                target_skus.extend([s.strip() for s in re.split(r'[\s,\t\n]+', line) if s.strip()])

            def clean_sku_base(val):                                  # Явная функция нормализации артикулов
                if pd.isna(val) or isinstance(val, (list, tuple, dict, pd.Series)): 
                    return ""                                         # Проверка ячейки на NaN пропуск или Series
                s = str(val).strip()                                  # Срезаем концевые пробелы текста
                return s[:-2] if s.endswith('.0') else s              # Отсекаем вещественный хвост .0

            for frame in [df, df_s3]:                                 # Перебор таблиц пайплайна для чистки
                f_col = context["sku_col"] if frame is df else s3_sku # Выбираем целевую колонку SKU для маппинга
                frame[f_col] = frame[f_col].apply(clean_sku_base)     # Применяем очистку от вещественных точек .0
            cleaned_targets = list(dict.fromkeys([clean_sku_base(s) for s in target_skus if s.strip()])) # SKU

            df_pre = df_s3[(df_s3[s3_date] >= pre_start) & (df_s3[s3_date] <= pre_end)].copy() # Фильтр ДО
            df_promo = df_s3[(df_s3[s3_date] >= promo_start) & (df_s3[s3_date] <= promo_end)].copy() # В акцию

            for sub in [df_pre, df_promo]:                            # Приведение к численным типам данных
                sub[s3_qty] = pd.to_numeric(sub[s3_qty], errors="coerce").fillna(0).astype(int)
                sub[s3_rev] = pd.to_numeric(sub[s3_rev], errors="coerce").fillna(0.0)

            agg_pre = df_pre.groupby(s3_sku).agg({s3_qty: "sum", s3_rev: "sum"}).reset_index() # Суммы ДО
            agg_promo = df_promo.groupby(s3_sku).agg({s3_qty: "sum", s3_rev: "sum"}).reset_index() # В акцию
            tot_qty_pre, tot_qty_promo = agg_pre[s3_qty].sum(), agg_promo[s3_qty].sum() # Объемы всей категории

            if cleaned_targets:                                       # Срез агрегаций под промо-список
                agg_pre = agg_pre[agg_pre[s3_sku].isin(cleaned_targets)]
                agg_promo = agg_promo[agg_promo[s3_sku].isin(cleaned_targets)]

            sku_m = context["sku_col"]                                # Извлекаем имя SKU базового листа
            df_res = df.copy().drop_duplicates(subset=[sku_m])        # Схлопываем дубли матрицы
            if cleaned_targets: df_res = df_res[df_res[sku_m].isin(cleaned_targets)].copy() # Фильтр по SKU

            df_res["Продажи ДО акции, шт"] = df_res[sku_m].map(agg_pre.set_index(s3_sku)[s3_qty].to_dict()).fillna(0).astype(int)
            df_res["Выручка ДО акции, руб"] = df_res[sku_m].map(agg_pre.set_index(s3_sku)[s3_rev].to_dict()).fillna(0.0).round(2)
            df_res["Продажи В АКЦИЮ, шт"] = df_res[sku_m].map(agg_promo.set_index(s3_sku)[s3_qty].to_dict()).fillna(0).astype(int)
            df_res["Выручка В АКЦИЮ, руб"] = df_res[sku_m].map(agg_promo.set_index(s3_sku)[s3_rev].to_dict()).fillna(0.0).round(2)

            # 1. МАТЕМАТИЧЕСКИ ТОЧНЫЙ РАСЧЕТ ЦЕН ЗА ЕДИНИЦУ ТОВАРA
            df_res["Цена ДО акции, руб"] = (df_res["Выручка ДО акции, руб"] / df_res["Продажи ДО акции, шт"]).fillna(df_res["_calc_price"]).round().astype(int)
            df_res["Цена В АКЦИЮ, руб"] = (df_res["Выручка В АКЦИЮ, руб"] / df_res["Продажи В АКЦИЮ, шт"]).fillna(df_res["_calc_price"] * 0.85).round().astype(int)

            # 2. ИСПРАВЛЕННЫЙ И НАДЁЖНЫЙ РАСЧЁТ SALES LIFT (БЕЗ РИСКА ИСКАЖЕНИЯ В EXCEL)
            def calc_exact_sales_lift(row):
                pre_qty = float(row["Продажи ДО акции, шт"])          # Штучные продажи до промо-периода
                promo_qty = float(row["Продажи В АКЦИЮ, шт"])         # Штучные продажи во время промо-периода
                if pre_qty <= 0:                                      # Защита от деления на ноль новинок матрицы
                    return 100.0 if promo_qty > 0 else 0.0            # Возвращаем 100% при старте с нулевой точки
                # Считаем точный коммерческий прирост (Sales Lift)
                lift_val = ((promo_qty - pre_qty) / pre_qty) * 100.0  # Процентное изменение объёма спроса
                return round(lift_val, 1)                             # Округляем строго до 1 знака после запятой

            # Записываем чистые числовые значения Sales Lift в датафрейм
            df_res["Прирост продаж (Sales Lift), %"] = df_res.apply(calc_exact_sales_lift, axis=1)

            # 3. КОРРЕКТНЫЙ АНАЛИЗАТОР РИСКОВ МАРКЕТИНГОВОЙ КАМПАНИИ
            def detect_cannibalism(row):                              # Анализатор перераспределения трафика
                lift = row["Прирост продаж (Sales Lift), %"]          # Считываем вычисленный процентный Lift
                p_pre = row["Цена ДО акции, руб"]                     # Считываем среднюю цену до скидок
                p_promo = row["Цена В АКЦИЮ, руб"]                    # Считываем среднюю цену во время скидок

                # Изменение штучного спроса на сопутствующий ассортимент остальной матрицы
                oth_change = ((tot_qty_promo - row["Продажи В АКЦИЮ, шт"]) - (tot_qty_pre - row["Продажи ДО акции, шт"])) / max(tot_qty_pre - row["Продажи ДО акции, шт"], 1) * 100

                if p_promo >= p_pre:                                  # Если цена фактически не снизилась —
                    return "Фактического снижения цены нет. Метрики стабильны." # акция не проводилась

                # ТРИГГЕРЫ ДЛЯ УСПЕШНОГО ПРОМО (Когда Sales Lift строго положительный)
                if lift > 0:                                          # Если штучные продажи реально выросли
                    if lift > 40 and oth_change < -15:                # Проверка критериев жесткого перетекания спроса
                        return f"КАННИБАЛИЗМ КАТЕГОРИИ Риск! Рост промо-SKU на +{lift}% подавил продажи остальной матрицы на {round(oth_change, 1)}%. Произошло замещение прибыли."
                    if tot_qty_promo <= tot_qty_pre:                  # Проверка критерия замещения внутри бренда
                        return "ЗАМЕЩЕНИЕ СПРОСА. Перераспределение трафика внутри бренда. Общий объем категории не вырос, акция не принесла новых клиентов."
                    # Если продажи выросли и соседи по матрице не пострадали
                    return f"ЧИСТЫЙ SALES LIFT Эффект успешного Промо! Скидка привлекла чистый новый трафик. Продажи соседей не пострадали (+{round(oth_change, 1)}%)."

                # ТРИГГЕР ДЛЯ ПРОВАЛЬНОГО ПРОМО (Когда Sales Lift отрицательный или равен нулю — как на скриншоте)
                return f"Эластичность спроса нулевая. Промо-акция полностью провалена. Снижение цены до {p_promo} руб. не вызвало роста спроса. Зафиксированы чистые убытки маржи."

            # Запускаем построчную разметку аналитических вердиктов по всей таблице промо-отчета
            df_res["Рекомендация по оптимизации"] = df_res.apply(detect_cannibalism, axis=1)
            context["has_task4"] = True                               # Помечаем успех выполнения шага
            return df_res                                             # Возвращаем рассчитанный датафрейм отчета
        except Exception as e: print(f"❌ Критическая ошибка при анализе промо-периодов: {e}")
        return df                                                     # Возвращаем исходный фрейм матрицы
