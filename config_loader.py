import os                                                             # Модуль файловой системы ОС
import json                                                           # Модуль для чтения файлов JSON
import sys                                                            # Системный модуль интерпретатора Python


def load_business_config():
    """Загружает конфигурацию из внешнего файла или встроенных ресурсов EXE."""
    # Определяем базовый путь (учитываем временную папку распаковки PyInstaller)
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS                                      # Путь внутри изолированного EXE файла
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))         # Обычный путь в режиме скрипта .py

    config_name = "config.json"
    
    # Попытка 1: Ищем внешний кастомный конфиг рядом с запущенным EXE-файлом
    external_path = os.path.join(os.path.dirname(sys.executable if hasattr(sys, 'frozen') else __file__), config_name)
    if os.path.exists(external_path):
        try:
            with open(external_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    # Попытка 2: Если внешнего нет, берем встроенный дефолтный из памяти EXE
    internal_path = os.path.join(base_path, config_name)
    if os.path.exists(internal_path):
        try:
            with open(internal_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    # Попытка 3: Жесткий фолбэк на случай полного отсутствия файлов на диске
    return {
        "MARGIN_THRESHOLDS": {"HIGH": 15.0, "LOW": 12.0},
        "MENU_OPTIONS": {
            "1": "Прогнозирование спроса и маржинальный аудит матрицы",
            "2": "Агрегация объемов и выручки из журнала транзакций",
            "3": "Продвинутый анализ промо-акции (Динамические периоды + Каннибализация)"
        }
    }
