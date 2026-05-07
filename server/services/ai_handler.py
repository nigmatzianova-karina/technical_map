import json
import logging
from typing import Dict, Any
import httpx
from tenacity import retry, stop_after_attempt, wait_exponential

logger = logging.getLogger(__name__)

class AIHandler:
    """Надёжный обработчик запросов к AI с повторными попытками и валидацией."""
    
    def __init__(self, api_key: str, model: str, timeout: int = 120):
        self.api_key = api_key.strip()
        self.model = model
        self.timeout = timeout
        
        if not self.api_key:
            raise ValueError("API ключ пустой")
    
    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        reraise=True
    )
    async def call_openrouter(
        self,
        system_prompt: str,
        user_message: str,
        temperature: float = 0.3,
        max_tokens: int = 4000,
        expect_json: bool = False
    ) -> Dict[str, Any]:
        """
        Вызов OpenRouter AI с обработкой ошибок.
        
        Returns:
            {
                "success": bool,
                "content": str,
                "json_data": Optional[dict],
                "error": Optional[str]
            }
        """
        try:
            async with httpx.AsyncClient(timeout=self.timeout) as client:
                response = await client.post(
                    "https://openrouter.ai/api/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json",
                        "HTTP-Referer": "http://localhost:8000",
                        "X-Title": "TK Generator"
                    },
                    json={
                        "model": self.model,
                        "messages": [
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_message}
                        ],
                        "temperature": temperature,
                        "max_tokens": max_tokens
                    }
                )
                
                if response.status_code == 401:
                    return {
                        "success": False,
                        "content": "",
                        "json_data": None,
                        "error": "Неверный API ключ"
                    }
                
                if response.status_code == 429:
                    return {
                        "success": False,
                        "content": "",
                        "json_data": None,
                        "error": "Превышен лимит запросов. Подождите немного."
                    }
                
                if response.status_code != 200:
                    error_text = response.text[:200]
                    return {
                        "success": False,
                        "content": "",
                        "json_data": None,
                        "error": f"Ошибка API ({response.status_code}): {error_text}"
                    }
                
                data = response.json()
                content = data["choices"][0]["message"]["content"]
                
                json_data = None
                if expect_json:
                    try:
                        json_start = content.find('{')
                        json_end = content.rfind('}') + 1
                        if json_start != -1 and json_end > json_start:
                            json_str = content[json_start:json_end]
                            json_data = json.loads(json_str)
                    except json.JSONDecodeError:
                        logger.warning(f"Не удалось распарсить JSON: {content[:200]}")
                        return {
                            "success": False,
                            "content": content,
                            "json_data": None,
                            "error": "AI вернул невалидный JSON"
                        }
                
                return {
                    "success": True,
                    "content": content,
                    "json_data": json_data,
                    "error": None
                }
                
        except httpx.TimeoutException:
            return {
                "success": False,
                "content": "",
                "json_data": None,
                "error": "Таймаут ответа от AI (превышено время ожидания)"
            }
        except Exception as e:
            logger.error(f"Ошибка вызова AI: {e}")
            return {
                "success": False,
                "content": "",
                "json_data": None,
                "error": f"Ошибка соединения: {str(e)}"
            }
        