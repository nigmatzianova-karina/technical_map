import os
from fastapi import HTTPException
import httpx
from dotenv import load_dotenv
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

load_dotenv()

REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "120"))
MAX_RETRIES = int(os.getenv("MAX_RETRIES", "3"))


@retry(
    stop=stop_after_attempt(MAX_RETRIES),
    wait=wait_exponential(multiplier=1, min=2, max=10),
    retry=retry_if_exception_type((httpx.RequestError, httpx.TimeoutException))
)
async def call_ai(messages: list, settings: dict) -> str:
    """Вызывает LLM провайдер (OpenRouter или HuggingFace) и возвращает ответ модели."""
    provider = settings.get("provider", "openrouter")
    api_key = settings.get("api_key", "")
    model = settings.get("model", "openai/gpt-4o")
    max_tokens = settings.get("max_tokens", 3000)

    if not api_key:
        raise HTTPException(status_code=400, detail="API ключ не установлен в настройках или .env")

    timeout = httpx.Timeout(timeout=REQUEST_TIMEOUT, connect=10.0)

    if provider == "openrouter":
        return _call_openrouter(api_key, model, messages, max_tokens)
    else:
        raise HTTPException(status_code=400, detail=f"Неизвестный провайдер: {provider}")


async def _call_openrouter(api_key: str, model: str, messages: list, max_tokens: int) -> str:
    """Выполняет запрос к OpenRouter API и возвращает содержимое ответа."""
    async with httpx.AsyncClient(timeout=httpx.Timeout(timeout=REQUEST_TIMEOUT, connect=10.0)) as client:
        resp = await client.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
                "HTTP-Referer": "http://localhost:8000",
                "X-Title": "TK AI Generator"
            },
            json={
                "model": model,
                "messages": messages,
                "temperature": float(os.getenv("DEFAULT_TEMPERATURE", "0.3")),
                "max_tokens": max_tokens
            }
        )
        if resp.status_code != 200:
            raise HTTPException(status_code=resp.status_code, detail=f"OpenRouter error: {resp.text}")
        data = resp.json()
        return data["choices"][0]["message"]["content"]
