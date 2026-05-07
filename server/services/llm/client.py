"""
Утилиты для работы с OpenRouter API.
Адаптировано под async-архитектуру твоего проекта.
"""

import json
import re
import logging
from typing import Dict, Any, List, Tuple, Optional

import httpx

logger = logging.getLogger(__name__)

OPENROUTER_BASE = "https://openrouter.ai/api/v1"


async def fetch_models_list() -> List[Dict[str, Any]]:
    """
    Асинхронно получает список моделей из публичного эндпоинта OpenRouter.
    Не требует API-ключа.
    """
    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            response = await client.get(f"{OPENROUTER_BASE}/models")
            response.raise_for_status()
            data = response.json()
            return data.get("data", [])
    except Exception as e:
        logger.error(f"Ошибка получения списка моделей: {e}")
        return []


def validate_api_key_sync(api_key: str) -> Tuple[bool, str]:
    """
    Синхронная проверка валидности API-ключа.
    Возвращает (is_valid, message).
    """
    import requests
    
    if not api_key or not api_key.strip():
        return False, "API ключ пустой"
    
    try:
        response = requests.get(
            f"{OPENROUTER_BASE}/auth/key",
            headers={"Authorization": f"Bearer {api_key.strip()}"},
            timeout=15
        )
        if response.status_code == 200:
            return True, "API ключ валиден"
        else:
            return False, f"Ошибка {response.status_code}: {response.text[:200]}"
    except requests.exceptions.Timeout:
        return False, "Таймаут соединения с OpenRouter"
    except requests.exceptions.SSLError:
        return False, "Ошибка SSL. Возможно, соединение блокируется."
    except Exception as e:
        return False, f"Ошибка соединения: {str(e)}"


async def validate_api_key_async(api_key: str) -> Tuple[bool, str]:
    """
    Асинхронная проверка валидности API-ключа.
    """
    if not api_key or not api_key.strip():
        return False, "API ключ пустой"
    
    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            response = await client.get(
                f"{OPENROUTER_BASE}/auth/key",
                headers={"Authorization": f"Bearer {api_key.strip()}"}
            )
            if response.status_code == 200:
                return True, "API ключ валиден"
            else:
                return False, f"Ошибка {response.status_code}: {response.text[:200]}"
    except httpx.TimeoutException:
        return False, "Таймаут соединения с OpenRouter"
    except httpx.SSLError:
        return False, "Ошибка SSL"
    except Exception as e:
        return False, f"Ошибка соединения: {str(e)}"


def get_model_context_length(model_id: str, models_list: List[Dict[str, Any]]) -> int:
    """Извлекает context_length для заданной модели из списка моделей."""
    for model in models_list:
        if model.get("id") == model_id:
            context_length = model.get("context_length")
            if isinstance(context_length, int) and context_length > 0:
                return context_length
    logger.warning(f"Не удалось найти context_length для модели {model_id}, используем 8192")
    return 8192


def parse_llm_json_response(content: str) -> Optional[Dict[str, Any]]:
    """
    Пытается распарсить ответ LLM как JSON.
    Обрабатывает случаи с ```json ... ``` обёрткой.
    """
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        match = re.search(r'```json\s*(.*?)\s*```', content, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(1))
            except json.JSONDecodeError:
                pass
        if content.strip().startswith('{') or content.strip().startswith('['):
            try:
                return json.loads(content.strip())
            except json.JSONDecodeError:
                pass
        logger.error(f"Невалидный JSON от LLM. Начало ответа:\n{content[:300]}")
        return None
    

def call_openrouter_sync(
    prompt: str,
    model: str,
    api_key: str,
    temperature: float = 0.3,
    max_tokens: int = 4000,
    response_format: Optional[str] = None,
    max_retries: int = 3
) -> str:
    """
    Синхронный вызов OpenRouter API с повторными попытками при 429.
    
    Args:
        prompt: Текст промпта для модели
        model: ID модели (например, "openai/gpt-4o-mini")
        api_key: API-ключ OpenRouter
        temperature: Температура генерации (0.0-2.0)
        max_tokens: Максимальное количество токенов в ответе
        response_format: "json_object" для принудительного JSON или None
        max_retries: Количество повторных попыток при ошибках
        
    Returns:
        str: Сырой текст ответа от модели (content из message)
        
    Raises:
        requests.RequestException: При исчерпании попыток или критической ошибке
        ValueError: Если ответ не содержит ожидаемой структуры
    """
    import requests
    import time
    
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
        "HTTP-Referer": "http://localhost:8000",
        "X-Title": "TK AI Generator"
    }
    
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
        "max_tokens": max_tokens
    }
    
    if response_format == "json_object":
        payload["response_format"] = {"type": "json_object"}
    
    last_error = None
    
    for attempt in range(max_retries):
        try:
            response = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=(10, 120)
            )
            
            if response.status_code == 429:
                wait_time = (2 ** attempt) + 1
                logger.warning(f"429 Too Many Requests, попытка {attempt+1}/{max_retries}, ждём {wait_time} сек")
                time.sleep(wait_time)
                continue
            
            response.raise_for_status()
            result = response.json()
            
            content = result.get("choices", [{}])[0].get("message", {}).get("content", "")
            if not content:
                raise ValueError("Пустой ответ от OpenRouter")
                
            return content
            
        except requests.exceptions.Timeout:
            last_error = "Таймаут соединения"
            logger.warning(f"Таймаут запроса, попытка {attempt+1}/{max_retries}")
        except requests.exceptions.SSLError as e:
            last_error = f"SSL ошибка: {e}"
            logger.error(last_error)
            break
        except requests.exceptions.RequestException as e:
            last_error = str(e)
            logger.warning(f"Ошибка запроса: {e}, попытка {attempt+1}/{max_retries}")
        except ValueError as e:
            raise ValueError(f"Невалидный ответ от OpenRouter: {e}")
        
        if attempt < max_retries - 1:
            time.sleep(2 ** attempt)
    
    raise requests.exceptions.RequestException(
        f"Не удалось получить ответ от OpenRouter после {max_retries} попыток. Последняя ошибка: {last_error}"
    )
