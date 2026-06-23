import pandas as pd                                                   # Библиотека манипулирования таблицами
from utils import clean_numeric_string                                # Слой ETL-очистки ячеек от мусора


class TransactionAggregationOperation:
    """Операция 2: Интерактивный точечный анализ транзакций по списку SKU."""
    def execute(self, df, xl_file, context):                          # Главный метод запуска команды
        target_skus = context.get("target_skus", [])                  # Извлекаем вставленный список из контекста

        if not target_skus:                                           # Проверка буфера ввода на пустоту
            print("⚠️ Список SKU пуст. Операция отменена.")          # Сообщение пользователю в терминал
            return df                                                 # Возвращаем исходную матрицу

        target_sheet = None                                           # Маркер целевого листа логов транзакций
        for sheet_name in xl_file.sheet_names:                        # Перебор вкладок книги Excel
            df_tmp = xl_file.parse(sheet_name, nrows=5)               # Экспресс-чтение первых 5 строк
            df_tmp.columns = [str(c).lower().strip() for c in df_tmp.columns] # Нормализация шапки листа от пробелов

            has_revenue = any("выруч" in c for c in df_tmp.columns)   # Поиск вектора выручки в шапке
            has_quantity = any("колич" in c for c in df_tmp.columns)  # Поиск вектора количества в шапке
            has_sku = any("sku" in c or "артикул" in c for c in df_tmp.columns) # Поиск вектора SKU в шапке

            if has_revenue and has_quantity and has_sku:              # Многокритериальная проверка шапки
                df_full_check = xl_file.parse(sheet_name)             # Загрузка полного листа в память
                if len(df_full_check) > 20:                           # Фильтр объема: отсекаем Лист 2 задания
                    target_sheet = sheet_name                         # Фиксируем валидный лист лога транзакций
                    break                                             # Выходим из цикла сканирования книги

        if not target_sheet:                                          # Если лог продаж не обнаружен —
            print("⚠️ Журнал транзакций не найден в файле.")         # вывод системного предупреждения
            return df                                                 # Прерывание выполнения команды

        try:                                                          # Блок защиты расчетов агрегации
            df_sheet3 = xl_file.parse(target_sheet, header=0)         # Полная загрузка листа логов транзакций
            df_sheet3.columns = [str(c).strip() for c in df_sheet3.columns] # Защита шапки лога от пробелов

            sku_col_s3 = next((c for c in df_sheet3.columns if "sku" in c.lower()), None) # Находим SKU в логе
            if not sku_col_s3:                                        # Если имя SKU не найдено в логе —
                sku_col_s3 = next((c for c in df_sheet3.columns if "артикул" in c.lower()), None) # берем Артикул

            qty_col_s3 = next((c for c in df_sheet3.columns if "колич" in c.lower()), None) # Столбец объемов в шт
            rev_col_s3 = next((c for c in df_sheet3.columns if "выруч" in c.lower()), None) # Столбец выручки в руб

            sku_col_main = context["sku_col"]                         # Получаем имя SKU базового листа

            def clean_sku(val):                                       # Локальная функция очистки артикулов
                if pd.isna(val):                                      # Проверка ячейки на NaN
                    return ""                                         # Возвращаем пустую строку
                s = str(val).strip()                                  # Срезаем концевые пробелы текста
                return s[:-2] if s.endswith('.0') else s              # Отсекаем вещественный хвост .0

            df_sheet3[sku_col_s3] = df_sheet3[sku_col_s3].apply(clean_sku) # Очищаем столбец SKU в логе

            # Уникализация буфера обмена пользователя
            cleaned_target_skus = list(dict.fromkeys([clean_sku(s) for s in target_skus if s.strip()]))

            # Фильтруем лог транзакций Листа 3 строго по уникальным введенным ID/SKU
            df_sheet3_filtered = df_sheet3[df_sheet3[sku_col_s3].isin(cleaned_target_skus)].copy()

            if df_sheet3_filtered.empty:                              # Если совпадений в логе транзакций нет
                print("⚠️ Введенные SKU не обнаружены в логе транзакций.") # Вывод предупреждения в консоль
                df_res = df.copy().drop_duplicates(subset=[sku_col_main]) # Клонируем матрицу без дублей
                df_res = df_res[df_res[sku_col_main].isin(cleaned_target_skus)].copy() # Оставляем искомые
                df_res["Количество проданных товаров, шт"] = 0        # Зануляем столбец количества штук
                df_res["Сумма выручки, руб"] = 0.0                    # Зануляем столбец финансовой выручки
                context["has_sheet3"] = True                          # Сигнализируем о завершении шага
                return df_res                                         # Возвращаем датафрейм-заглушку

            df_sheet3_filtered[qty_col_s3] = df_sheet3_filtered[qty_col_s3].apply(clean_numeric_string).astype(int)
            df_sheet3_filtered[rev_col_s3] = df_sheet3_filtered[rev_col_s3].apply(clean_numeric_string)

            # ---------------------------------------------------------
            # ИНТЕЛЛЕКТУАЛЬНЫЙ АНАЛИЗ ЦЕН ИЗ ТРАНЗАКЦИЙ (Discovery Price)
            # ---------------------------------------------------------
            # Находим реальную розничную цену за 1 шт внутри каждой транзакции
            df_sheet3_filtered["_calculated_price_per_item"] = (
                df_sheet3_filtered[rev_col_s3] / df_sheet3_filtered[qty_col_s3]
            ).fillna(0).round().astype(int)                           # Округляем до целых рублей e-com

            # Строим карту вычисленных цен {SKU: Цена} на основе лога транзакций
            calculated_price_map = df_sheet3_filtered.groupby(sku_col_s3)["_calculated_price_per_item"].max().to_dict()

            # Агрегируем лог и суммируем продажи, схлопывая дублирующиеся строки лога
            aggregated = df_sheet3_filtered.groupby(sku_col_s3).agg({ # Группируем лог транзакций по SKU
                qty_col_s3: "sum",                                    # Суммируем штучные продажи по артикулу
                rev_col_s3: "sum"                                     # Суммируем рублевую выручку по артикулу
            }).reset_index()                                          # Сбрасываем индексы таблицы

            qty_map = aggregated.set_index(sku_col_s3)[qty_col_s3].to_dict() # Карта маппинга количества штук
            rev_map = aggregated.set_index(sku_col_s3)[rev_col_s3].to_dict() # Карта маппинга суммарной выручки

            # Исключаем появление дубликатов из Листа 1 ассортимента
            df_res = df.copy().drop_duplicates(subset=[sku_col_main]) # Убираем дубли артикулов базового листа
            df_res = df_res[df_res[sku_col_main].isin(cleaned_target_skus)].copy() # Фильтруем под искомый список

            missing_skus = [s for s in cleaned_target_skus if s not in df_res[sku_col_main].values] # Новые артикулы
            if missing_skus:                                          # If введены новые, отсутствующие SKU
                df_missing = pd.DataFrame({sku_col_main: missing_skus}) # Строим фрейм для новых позиций лога
                df_res = pd.concat([df_res, df_missing], ignore_index=True) # Добавляем новые SKU в конец таблицы

            # Заполняем коммерческие метрики отчета
            df_res["Количество проданных товаров, шт"] = df_res[sku_col_main].map(qty_map).fillna(0).astype(int) # Шт
            df_res["Сумма выручки, руб"] = df_res[sku_col_main].map(rev_map).fillna(0.0).round(2) # Накладываем выручку

            # ---------------------------------------------------------
            # АВТО-УСТРАНЕНИЕ НУЛЕЙ В ЦЕНАХ
            # ---------------------------------------------------------
            # Переопределяем розничную цену: если на Листе 1 был ноль или пропуск,
            # скрипт автоматически подтянет вычисленную цену из лога продаж Листа 3
            price_col_main = context.get("price_col")                 # Извлекаем оригинальное имя столбца цены
            if price_col_main and price_col_main in df_res.columns:   # Если столбец цен присутствует в структуре
                def fallback_price(row):                              # Метод умного заполнения пропусков цен
                    current_p = pd.to_numeric(row[price_col_main], errors='coerce') # Проверяем текущую цену
                    if pd.isna(current_p) or current_p == 0:          # Если цена равна нулю или отсутствует (NaN)
                        return calculated_price_map.get(row[sku_col_main], 0) # берем расчетную цену транзакции
                    return row[price_col_main]                        # Иначе оставляем исходную цену
                df_res[price_col_main] = df_res.apply(fallback_price, axis=1) # Применяем автозамену ко всей матрице

            context["has_sheet3"] = True                              # Помечаем успех выполнения операции
            return df_res                                             # Возвращаем чистый агрегированный датафрейм
        except Exception as e:                                        # Отлов системных исключений расчетов
            print(f"❌ Ошибка при точечной агрегации: {e}")           # Вывод текста ошибки в консоль терминала
            context["has_sheet3"] = False                             # Фиксируем сбой в контексте отчета
        return df                                                     # Возвращаем исходный фрейм матрицы
