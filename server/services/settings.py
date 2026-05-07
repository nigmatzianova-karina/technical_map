import json
import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

SETTINGS_FILE = Path("settings.json")

DEFAULT_SETTINGS = {
    "provider": "openrouter",
    "api_key": "",
    "model": "anthropic/claude-3.5-sonnet",
    "max_tokens": 3000,
    "master_prompt": """Ты инженер, специалист по формированию технологических карт и работ по ТОиР оборудования.

{file_instruction}

Необходимо заполнить:
1. Столбец "Элемент" — основной крупный элемент, входящий в состав узла. Например: Система смазки.
2. Столбец "Подэлемент" — более мелкий элемент, входящий в состав элемента. Например: Картер.

Правила:
• Каждый новый узел, элемент и подэлемент — в отдельной строке по порядку.
• НЕ вноси как "Элемент" или "Подэлемент": гайки, шайбы, винты, шпильки, хомуты, болты, штифты, шпонки.
• Если в столбцах несколько слов — первое слово всегда существительное, остальные после него.
• Элемент и подэлемент — в единственном числе, именительном падеже.
• Слова нельзя сокращать и заменять синонимами.
• Другие столбцы таблицы не удаляй и не изменяй.

ОТВЕТ ДОЛЖЕН БЫТЬ В СТРОГОМ ФОРМАТЕ:

[ТЕКСТ_ОТВЕТ]
Краткое текстовое описание результата для пользователя.
[/ТЕКСТ_ОТВЕТ]

[ТАБЛИЦА]
Элемент|Подэлемент|Наименование операции|Краткое содержание работ|Вид ТОиР|Периодичность|Норма времени, часов|Количество исполнителей|Профессия/Квалификация|Трудоёмкость, человеко/часов|Наименование ТМЦ|Количество ТМЦ|Единицы измерения ТМЦ|Наименование инструмента|Средства индивидуальной защиты|Требования по безопасности
Система смазки|Картер|Осмотр|Визуальный осмотр картера на наличие трещин и подтёков|ТО-1|4320|2.0|1|Слесарь по ремонту автомобилей, 3 разряд|2.0|||||Каска защитная, 1 шт; Очки защитные, 1 шт; Перчатки защитные, 1 пара|Затормозить технику; Выполнять работы при неработающем двигателе
[/ТАБЛИЦА]

ВАЖНО: Каждая строка таблицы — значения через "|". Всего 16 столбцов. Если данных нет — оставьте пусто (||)."""
}


def load_settings() -> dict:
    """Загружает настройки из .env и settings.json, возвращает словарь настроек."""
    settings = {
        "provider": "openrouter",
        "api_key": os.getenv("OPENROUTER_API_KEY", ""),
        "model": "anthropic/claude-3.5-sonnet",
        "max_tokens": int(os.getenv("DEFAULT_MAX_TOKENS", "3000")),
        "master_prompt": DEFAULT_SETTINGS["master_prompt"]
    }

    if SETTINGS_FILE.exists():
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                file_settings = json.load(f)
                if not os.getenv("OPENROUTER_API_KEY"):
                    settings["api_key"] = file_settings.get("api_key", "")
                settings.update({k: v for k, v in file_settings.items() if k not in ("api_key", "provider", "model")})
        except Exception as e:
            print(f"⚠️ Ошибка чтения settings.json: {e}")

    return settings


def save_settings(settings: dict) -> None:
    """Сохраняет словарь настроек в файл settings.json."""
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
