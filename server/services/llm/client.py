"""
Клиент для OpenRouter API: асинхронный вызов, парсинг JSON-ответов.
"""

import json
import re
from typing import Dict, Any, Tuple, Optional
import asyncio
import httpx

OPENROUTER_BASE = "https://openrouter.ai/api/v1"


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
        return None


async def call_openrouter_async(
        prompt: str,
        model: str,
        api_key: str,
        temperature: float = 0.3,
        max_tokens: int = 4000,
        response_format: Optional[str] = None,
        max_retries: int = 3
) -> str:
    """Асинхронный вызов OpenRouter API с повторными попытками."""
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
            async with httpx.AsyncClient(timeout=120.0) as client:
                response = await client.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers=headers,
                    json=payload
                )
                if response.status_code == 429:
                    wait_time = (2 ** attempt) + 1
                    await asyncio.sleep(wait_time)
                    continue
                response.raise_for_status()
                result = response.json()
                content = result.get("choices", [{}])[0].get("message", {}).get("content", "")
                if not content:
                    raise ValueError("Пустой ответ от OpenRouter")
                return content
        except (httpx.TimeoutException, httpx.HTTPStatusError, ValueError) as e:
            last_error = str(e)
            if attempt < max_retries - 1:
                await asyncio.sleep(2 ** attempt)
    raise Exception(f"Не удалось получить ответ после {max_retries} попыток. Последняя ошибка: {last_error}")
