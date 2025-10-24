import os

# Версия алгоритма (будет прокидываться в ответ)
ALGO_VERSION = os.getenv("ALGO_VERSION", "v0.6-MVP")

# Место для правил безопасности/ограничений (пока не используем)
SAFETY_RULES = {}
