import json
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

SETTINGS_FILE = Path("settings.json")

DEFAULT_SETTINGS = {
    "provider": "openrouter",
    "api_key": "",
    "model": "openai/gpt-4o-mini",
    "max_tokens": 3000,
    "temperature": 0.3,
    "master_prompt": "Ты инженер, специалист по формированию технологических карт и работ по ТОиР оборудования..."
}

def load_settings():
    settings = DEFAULT_SETTINGS.copy()
    if SETTINGS_FILE.exists():
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            file_settings = json.load(f)
            settings.update(file_settings)
    return settings

def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
