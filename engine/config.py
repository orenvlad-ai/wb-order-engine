import os

# Версия алгоритма (будет прокидываться в ответ)
ALGO_VERSION = os.getenv("ALGO_VERSION", "v0.6-MVP")

# Мягкий буфер для снижения плана (единиц)
# Используем, когда считаем рекомендуемое снижение продаж (reduce_plan_to).
SOFT_BUFFER = int(os.getenv("SOFT_BUFFER", "50"))

# Резерв под будущие корректировки, сейчас 0 (выключено)
BUFFER_WEIGHT = float(os.getenv("BUFFER_WEIGHT", "0"))

# Место для правил безопасности/ограничений (пока не используем)
SAFETY_RULES = {}
