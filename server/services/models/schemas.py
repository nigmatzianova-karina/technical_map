"""
Pydantic-схемы для валидации запросов и ответов API.
"""

from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional


class ParseResult(BaseModel):
    """Результат парсинга документа."""
    pages_text: List[str] = Field(default_factory=list)
    tables_markdown: List[str] = Field(default_factory=list)
    full_markdown: str = ""
    metadata: Dict[str, Any] = Field(default_factory=dict)


class LLMGenerateRequest(BaseModel):
    """Запрос на генерацию через LLM."""
    prompt: str
    model: str
    api_key: str
    temperature: Optional[float] = Field(default=0.3, ge=0.0, le=2.0)
    max_tokens: int = Field(default=3000, ge=1, le=32000)
    response_format: Optional[str] = None


class LLMGenerateResponse(BaseModel):
    """Ответ от LLM."""
    content: str
    model_used: str
    usage: Optional[Dict[str, int]] = None


class KeyCheckRequest(BaseModel):
    api_key: str


class KeyCheckResponse(BaseModel):
    valid: bool
    message: str


class ModelInfo(BaseModel):
    id: str
    name: str
    context_length: Optional[int] = None
    pricing_prompt: Optional[str] = None
    pricing_completion: Optional[str] = None
    is_free: bool = False


class ExtractedEntity(BaseModel):
    """Текстовая сущность из документа."""
    name: str
    value: str
    unit: Optional[str] = None
    source_page: Optional[int] = None


class ExtractedTable(BaseModel):
    """Таблица из документа."""
    caption: Optional[str] = None
    headers: List[str] = Field(default_factory=list)
    rows: List[List[str]] = Field(default_factory=list)
    footnotes: List[str] = Field(default_factory=list)
    source_page: Optional[int] = None


class ExtractionResult(BaseModel):
    """Результат извлечения структурированных данных."""
    entities: List[ExtractedEntity] = Field(default_factory=list)
    tables: List[ExtractedTable] = Field(default_factory=list)
    page_summaries: Dict[int, str] = Field(default_factory=dict)
