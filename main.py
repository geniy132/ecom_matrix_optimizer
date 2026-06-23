import os                                                             # Модуль файловой системы ОС
import re                                                             # Модуль регулярных выражений текста
import pandas as pd                                                   # Аналитическая библиотека таблиц Python

from config_loader import load_business_config                        # Загрузчик настроек и параметров меню
from core import OPERATION_REGISTRY                                   # Реестр SOLID операций аналитики
from discovery import identify_business_columns, scan_and_load_matrix # Инструменты Excel-Discovery листов
from styler import apply_corporate_style                              # Модуль визуального дизайна Excel таблиц
from utils import clean_numeric_string, extract_date_columns          # Модуль очистки данных и ETL


def run_pipeline(file_path, choice):
    xl_file = pd.ExcelFile(file_path)                                 # Дескриптор книги Excel в памяти ОЗУ

    # Вызов адаптивного сканера из модуля Discovery
    df, matrix_sheet = scan_and_load_matrix(xl_file, choice)
    if df is None:                                                    # Если необходимый лист не обнаружен
        print("❌ Критическая ошибка: Целевой лист данных не найден.") # выводим системный сбой в консоль
        return                                                         # Прерывание выполнения пайплайна

    # Распознавание названий столбцов по ключевым бизнес-синонимам
    name_col, sku_col, price_col, margin_col = identify_business_columns(df)

    if choice in ["2", "3"]:
        if not sku_col:                                                # Если автосканер не локализовал SKU
            sku_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["sku", "артикул"])), None)
        if not sku_col:                                                # Если артикул совсем отсутствует на листе
            print("❌ Ошибка: На листе транзакций не найден столбец SKU или Артикул.")
            return                                                     # Аварийный тихий выход из оркестратора
    else:
        # Для стандартного сценария (1) жестко требуем всю плоскую коммерческую шапку
        if not all([name_col, sku_col, price_col, margin_col]):
            print("❌ Ошибка: Структура бизнес-метрик на найденном листе повреждена.")
            return                                                     # Аварийный тихий выход из оркестратора

    if choice in ["2", "3"]:
        df = df.dropna(subset=[sku_col]).copy()
    else:
        df = df.dropna(subset=[name_col, sku_col]).copy()             # чистим по имени и SKU

    def clean_sku_base(val):
        """Очищает SKU от пробелов и срезает вещественные хвосты типа .0"""
        if pd.isna(val):                                               # Проверка ячейки на пустые NaN значения
            return ""                                                  # Возвращаем пустую текстовую строку
        s = str(val).strip()                                           # Срезаем случайные боковые пробелы текста
        return s[:-2] if s.endswith('.0') else s                       # Удаляем .0, если это дробный формат Excel

    df[sku_col] = df[sku_col].apply(clean_sku_base)                   # Безопасно нормализуем артикулы в таблице
    df = df[df[sku_col].astype(str).str.contains(r"^\d+$", regex=True)].copy() # Исключаем строку "Итого"

    # Ищем временные ряды посуточных дат продаж (только если они есть на выбранном листе)
    date_cols = [c for c in df.columns if c not in [name_col, sku_col, price_col, margin_col] and any(char.isdigit() for char in str(c))]
    count_days = len(date_cols)                                       # Вычисляем длину временного ряда в днях

    if count_days == 0 and choice == "1":                             # Если дат нет, а выбран режим прогноза
        print("❌ Ошибка: Временные ряды продаж не обнаружены.")       # Вывод сообщения о сбое данных Листа 1
        return                                                         # Остановка аналитического конвейера

    for d_col in date_cols:                                            # Перебор хронологических столбцов дат
        df[d_col] = pd.to_numeric(df[d_col], errors="coerce").fillna(0) # Преобразуем ячейки дат в чистые числа

    # Страхуем расчёт цен и маржи, если на листе их нет (например, на чистом логе Листа 3)
    df["_calc_price"] = df[price_col].apply(clean_numeric_string) if price_col else 0.0
    df["_calc_margin_raw"] = df[margin_col].apply(clean_numeric_string) if margin_col else 0.0
    df["_calc_margin_pct"] = df["_calc_margin_raw"].apply(lambda x: x / 100 if x > 1 else x)

    if choice != "3":
        business_fields = [name_col, sku_col, price_col, margin_col]   # Собираем список коммерческих полей
        date_list = extract_date_columns(df, business_fields)          # Извлекаем список горизонтальных дат
        date_cols = [item for item, _ in date_list]                    # Сохраняем текстовые имена колонок дат
        count_days = len(date_cols)                                    # Считаем общее число дней в периоде
        _, last_date = date_list[-1] if date_list else (None, None)    # Фиксируем крайний Timestamp календаря
    else:
        date_cols = []                                                 # Пустой список горизонтальных столбцов
        count_days = 0                                                 # Обнуляем счетчик дней базового листа
        last_date = pd.Timestamp.now()                                 # Задаем текущее время как фолбэк дат

    # Составляем единый словарь контекста выполнения для SOLID-команд
    context = {                                                        # Инициализация контекста для выгрузки отчета
        "date_cols": date_cols, "count_days": count_days,              # Параметры длины временного ряда продаж
        "last_date": last_date, "sku_col": sku_col,                    # Метрики дат и имя целевого столбца SKU
        "price_col": price_col,                                        # Передаем оригинальное имя столбца цен
        "forecast_cols": [], "has_sheet3": False,                      # Буферы созданных колонок и лога
        "target_skus": []                                              # Сюда запишутся вставленные из консоли SKU
    }                                                                  # Конец структуры контекста пайплайна

    if choice == "2":                                                 # ИНТЕРАКТИВНЫЙ ВВОД СТОЛБЦА ИЗ EXCEL
        print("\n👉 Вставьте список нужных SKU/Артикулов (ПКМ для вставки столбца из Excel) и нажмите Enter ДВАЖДЫ:")
        while True:                                                   # Цикл ожидания многострочного буфера консоли
            line = input().strip()                                    # Считываем текущую строку ввода
            if not line: break                                        # Пустой Enter — прерываем цикл ввода данных
            context["target_skus"].extend([s.strip() for s in re.split(r'[\s,\t\n]+', line) if s.strip()]) # Сохраняем в контекст

    operations = OPERATION_REGISTRY.get(choice, [])                   # Извлекаем цепочку классов операций SOLID
    for op in operations:                                             # Последовательный конвейерный запуск
        df = op.execute(df, xl_file, context)                         # Передаем СТРОГО единый словарь context

    if df is None or df.empty:                                        # Проверка результирующего массива на пустоту
        print("❌ Критическая ошибка: Данные для построения отчета отсутствуют.")
        return

    # Динамически выстраиваем структуру финального документа Excel
    output_structure = {name_col: "Полное название товара", sku_col: "SKU"}

    if choice == "3" and context.get("has_task4"):
        output_structure["Цена ДО акции, руб"] = "Цена ДО акции, руб"
        output_structure["Продажи ДО акции, шт"] = "Продоно ДО акции, шт"
        output_structure["Выручка ДО акции, руб"] = "Выручка ДО акции, руб"
        output_structure["Цена В АКЦИЮ, руб"] = "Цена В АКЦИЮ, руб"
        output_structure["Продажи В АКЦИЮ, шт"] = "Продоно В АКЦИЮ, шт"
        output_structure["Выручка В АКЦИЮ, руб"] = "Выручка В АКЦИЮ, руб"
        output_structure["Прирост продаж (Sales Lift), %"] = "Прирост продаж (Sales Lift), %"
        output_structure["Рекомендация по оптимизации"] = "Аналитический вывод по акции и рискам каннибализма"
    else:
        # Стандартная структура вывода для остальных сценариев (1, 2)
        if price_col: output_structure[price_col] = "Цена розничная, руб"
        if margin_col: output_structure[margin_col] = "Маржа, %"

        # Если выполнялась операция 1 (Прогноз спроса) — добавляем весенние столбцы
        if choice == "1" and context["forecast_cols"]:
            output_structure["Всего_продано_шт"] = "Продано за период, шт"
            for f in context["forecast_cols"]: output_structure[f] = f
            output_structure["Маржинальная прибыль за период, руб"] = "Маржинальная прибыль за период, руб"

        # Если выполнялась операция 2 (Агрегация логов) — добавляем штуки и выручку
        if choice == "2" and context["has_sheet3"]:
            output_structure["Количество проданных товаров, шт"] = "Количество проданных товаров, шт"
            output_structure["Сумма выручки, руб"] = "Сумма выручки, руб"

        # Добавляем инсайты робота, только если они рассчитывались в пайплайне
        if "Рекомендация по оптимизации" in df.columns and choice == "1":
            output_structure["Рекомендация по оптимизации"] = "Рекомендация по оптимизации"

    valid_keys = [k for k in output_structure.keys() if k in df.columns]
    final_df = df[valid_keys].rename(columns={k: output_structure[k] for k in valid_keys})

    output_file = os.path.join(os.path.dirname(file_path), "Итоговый_анализ.xlsx") if os.path.dirname(file_path) else "Итоговый_анализ.xlsx"

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:    # Открываем классический Excel-writer
        final_df.to_excel(writer, index=False, sheet_name='Аналитика матрицы') # Сохраняем чистый датафрейм
        worksheet = writer.sheets['Аналитика матрицы']                 # Передаем лист openpyxl в наш styler
        apply_corporate_style(worksheet, final_df)                     # Запуск безопасной внешней стилизации

    print(f"🎉 Проект успешно выполнен. Красивый отчет сохранен в: {output_file}")


if __name__ == "__main__":
    print("=== Утилита автоматизации маржинального e-com анализа ===")
    cfg = load_business_config()                                      # Загружаем меню из внешнего файла JSON
    menu = cfg.get("MENU_OPTIONS", {"1": "Прогноз спроса", "2": "Агрегация транзакций"}) # Чтение
    print("\n📋 Доступные аналитические сценарии:")                     # Вывод заголовка консольного меню
    for key, val in menu.items(): print(f" [{key}] - {val}")          # Построчный вывод доступных шагов меню
    choice_input = input("\n👉 Выберите номер операции: ").strip()     # Считываем шаг выбранного сценария пользователя
    user_input = input("👉 Перетащите файл Excel или вставьте путь (ПКМ): ").strip() # Запрос пути к Excel-файлу
    clean_path = user_input.strip("'\"")                              # Срезаем кавычки операционной системы
    if os.path.exists(clean_path):                                     # Валидация физического наличия файла
        try:
            run_pipeline(clean_path, choice_input)                   # Запуск аналитического конвейера расчетов
        except Exception as e: print(f"❌ Критический сбой при расчете: {e}") # Отлов и изолирование непредвиденных сбоев
    else:
        print(f"❌ Ошибка: Файл по пути '{clean_path}' не найден.")  # Предупреждение о неверно вставленном пути
    input("\nНажмите Enter для выхода...")                            # Удержание консольного окна открытым
