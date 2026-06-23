from config_loader import load_business_config                        # Импорт общего загрузчика настроек
from operations_forecast import DemandForecastOperation               # Сценарий 1: Прогнозы и аудит
from operations_transactions import TransactionAggregationOperation   # Сценарий 2: Агрегация транзакций
from operations_promo import PromoEffectivenessOperation              # Сценарий 4: Анализ промо-акций

# РЕЕСТР ОПЕРАЦИЙ SOLID: Связывает пункты меню консоли с изолированными модулями
OPERATION_REGISTRY = {
    "1": [DemandForecastOperation()],
    "2": [TransactionAggregationOperation()],
    "3": [PromoEffectivenessOperation()]                              # Запуск продвинутого анализа промо
}
