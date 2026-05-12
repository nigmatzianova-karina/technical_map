from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
import json
from pathlib import Path
import httpx
from server.core.config import load_settings


router = APIRouter(prefix="/api", tags=["settings"])

SETTINGS_FILE = Path(__file__).parent.parent.parent / "settings.json"


class SettingsRequest(BaseModel):
    provider: str = "openrouter"
    api_key: str = ""
    model: str = "openai/gpt-4o-mini"
    max_tokens: int = 3000
    master_prompt: str = ""
    temperature: float = 0.3


@router.get("/settings")
async def get_settings():
    """Получить текущие настройки."""
    try:
        return load_settings()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка чтения настроек: {str(e)}")


@router.post("/settings")
async def save_settings(request: SettingsRequest):
    """Сохранить настройки."""
    try:
        settings_data = {
            "provider": request.provider,
            "api_key": request.api_key,
            "model": request.model,
            "max_tokens": request.max_tokens,
            "master_prompt": request.master_prompt,
            "temperature": request.temperature
        }
        
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings_data, f, indent=2, ensure_ascii=False)
        
        return {"status": "success", "message": "Настройки сохранены"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка сохранения: {str(e)}")


@router.get("/models/{provider}")
async def get_models(provider: str):
    """Получить список моделей для провайдера."""
    models = {
        "openrouter": [
            {"value": "openai/gpt-4o-2024-08-06", "label": "GPT-4o"},
            {"value": "openai/gpt-4o-mini-2024-07-18", "label": "GPT-4o Mini"},
            {"value": "anthropic/claude-3-5-sonnet-20240620", "label": "Claude 3.5 Sonnet"},
            {"value": "anthropic/claude-3-haiku-20240307", "label": "Claude 3 Haiku"},
            {"value": "meta-llama/llama-3.3-70b-instruct", "label": "Llama 3.3 70B"},
            {"value": "google/gemini-2.0-flash-exp", "label": "Gemini 2.0 Flash"},
            {"value": "mistralai/mistral-large-2411", "label": "Mistral Large 2"},
            {"value": "z-ai/glm-4.5-air:free", "label": "GLM-4.5 Air (Free)"},
            {"value": "inclusionai/ling-2.6-1t:free", "label": "Ling-2.6-1T (Free)"},
            {"value": "openai/gpt-oss-120b:free", "label": "GPT-OSS 120B (Free)"},
            {"value": "openrouter/free", "label": "OpenRouter Free"}
        ],
        "openai": [
            {"value": "gpt-4o", "label": "GPT-4o"},
            {"value": "gpt-4o-mini", "label": "GPT-4o Mini"},
            {"value": "o1", "label": "o1"},
            {"value": "o3-mini", "label": "o3-mini"}
        ]
    }
    
    return models.get(provider, models["openrouter"])


class KeyValidationRequest(BaseModel):
    api_key: str

@router.post("/key/validate")
async def validate_key_endpoint(req: KeyValidationRequest):
    """Проверяет валидность API ключа через OpenRouter."""
    api_key = req.api_key.strip()
    if not api_key:
        return {"valid": False, "message": "Ключ пустой"}

    try:
        async with httpx.AsyncClient(timeout=8.0) as client:
            response = await client.get(
                "https://openrouter.ai/api/v1/auth/key",
                headers={"Authorization": f"Bearer {api_key}"}
            )
            
            if response.status_code == 200:
                return {"valid": True, "message": "Ключ валиден"}
            else:
                return {"valid": False, "message": f"Ошибка {response.status_code}: {response.text[:100]}"}
                
    except httpx.TimeoutException:
        return {"valid": False, "message": "Таймаут соединения с OpenRouter"}
    except Exception as e:
        return {"valid": False, "message": str(e)}
    
    