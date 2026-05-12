import json
from pathlib import Path
from dotenv import load_dotenv
import os

load_dotenv()

SETTINGS_FILE = Path("settings.json")

DEFAULT_SETTINGS = {
    "provider": "openrouter",
    "api_key": os.getenv("OPENROUTER_API_KEY", ""),
    "model": "openai/gpt-4o-mini",
    "max_tokens": 3000,
    "temperature": 0.3,
    "master_prompt": (
        "Ты инженер, специалист по формированию технологических карт и работ по ТОиР оборудования.\n\n"
        "{file_instruction}\n\n"
        "Необходимо заполнить:\n"
        "1. Столбец \"Элемент\" — основной крупный элемент, входящий в состав узла. Например: Система смазки.\n"
        "2. Столбец \"Подэлемент\" — более мелкий элемент, входящий в состав элемента. Например: Картер.\n\n"
        "Правила:\n"
        "• Каждый новый узел, элемент и подэлемент — в отдельной строке по порядку.\n"
        "• НЕ вноси как \"Элемент\" или \"Подэлемент\": гайки, шайбы, винты, шпильки, хомуты, болты, штифты, шпонки.\n"
        "• Если в столбцах несколько слов — первое слово всегда существительное, остальные после него.\n"
        "• Элемент и подэлемент — в единственном числе, именительном падеже.\n"
        "• Слова нельзя сокращать и заменять синонимами.\n"
        "• Другие столбцы таблицы не удаляй и не изменяй."
    )
}

def load_settings():
    settings = DEFAULT_SETTINGS.copy()
    if SETTINGS_FILE.exists():
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            file_settings = json.load(f)
            if not settings["api_key"]:
                settings["api_key"] = file_settings.get("api_key", "")
            for key in ["provider", "model", "max_tokens", "temperature", "master_prompt"]:
                if key in file_settings:
                    settings[key] = file_settings[key]
    return settings

def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
