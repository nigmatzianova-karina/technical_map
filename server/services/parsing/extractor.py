"""
Базовый экстрактор данных из документа через LLM.
Простая синхронная версия — не зависит от сложного batch_processor.
"""

import logging
import json
from typing import Dict, Any
from server.services.parsing.pdf_parser import parse_pdf_bytes
from server.services.llm.client import (
    call_openrouter_sync,
    parse_llm_json_response
)
from server.services.llm.prompts import (
    get_batch_tables_prompt,
    get_tech_card_prompt,
    get_bom_prompt,
    get_tables_summary_prompt
)

logger = logging.getLogger(__name__)


def extract_tech_card(
    pdf_bytes: bytes,
    model: str,
    api_key: str,
    has_document: bool = True,
    temperature: float = 0.3
) -> Dict[str, Any]:
    """
    Извлекает технологическую карту из PDF.
    
    Returns:
        {
            "success": bool,
            "result": dict | None,
            "error": str | None,
            "pages_processed": int
        }
    """
    try:
        parsed = parse_pdf_bytes(pdf_bytes)
        pages = parsed["pages_text"]
        
        if not pages or not any(pages):
            return {
                "success": False,
                "result": None,
                "error": "Не удалось извлечь текст из PDF",
                "pages_processed": 0
            }
        
        sample_content = "\n\n".join(pages[:2])
        prompt = get_tech_card_prompt(has_document=has_document)
        full_prompt = f"{prompt}\n\nСодержимое документа (фрагмент):\n{sample_content}"
        
        response_text = call_openrouter_sync(
            prompt=full_prompt,
            model=model,
            api_key=api_key,
            temperature=temperature,
            response_format="json_object"
        )
        
        result = parse_llm_json_response(response_text)
        if not result or "rows" not in result:
            logger.warning(f"LLM вернул невалидный формат. Начало ответа: {response_text[:200]}")
            return {
                "success": False,
                "result": None,
                "error": "LLM вернул ответ в неверном формате",
                "pages_processed": len(pages)
            }
        
        return {
            "success": True,
            "result": result,
            "error": None,
            "pages_processed": len(pages)
        }
        
    except Exception as e:
        logger.error(f"Ошибка в extract_tech_card: {e}", exc_info=True)
        return {
            "success": False,
            "result": None,
            "error": str(e),
            "pages_processed": 0
        }


def extract_bom(
    pdf_bytes: bytes,
    model: str,
    api_key: str,
    max_pages: int = 5,
    temperature: float = 0.2
) -> Dict[str, Any]:
    """
    Извлекает иерархическую спецификацию (BOM) из PDF.
    """
    try:
        parsed = parse_pdf_bytes(pdf_bytes)
        pages = parsed["pages_text"]
        
        if not pages:
            return {
                "success": False,
                "result": None,
                "error": "Пустой документ",
                "pages_processed": 0
            }
        
        sample_pages = pages[:max_pages]
        sample_content = "\n\n".join(
            f"<!-- СТРАНИЦА {i+1} -->\n{txt}" 
            for i, txt in enumerate(sample_pages)
        )
        
        prompt = get_bom_prompt(1, len(sample_pages), sample_content)
        
        response_text = call_openrouter_sync(
            prompt=prompt,
            model=model,
            api_key=api_key,
            temperature=temperature,
            response_format="json_object"
        )
        
        result = parse_llm_json_response(response_text)
        if not result or not isinstance(result, list):
            return {
                "success": False,
                "result": None,
                "error": "LLM вернул невалидный BOM-формат",
                "pages_processed": len(pages)
            }
        
        return {
            "success": True,
            "result": {"bom_tree": result},
            "error": None,
            "pages_processed": len(pages)
        }
        
    except Exception as e:
        logger.error(f"Ошибка в extract_bom: {e}", exc_info=True)
        return {
            "success": False,
            "result": None,
            "error": str(e),
            "pages_processed": 0
        }


def extract_tables_summary(
    pdf_bytes: bytes,
    model: str,
    api_key: str,
    temperature: float = 0.1
) -> Dict[str, Any]:
    """
    Извлекает числовые характеристики и формирует сводную таблицу.
    """
    try:
        parsed = parse_pdf_bytes(pdf_bytes)
        pages = parsed["pages_text"]
        
        if not pages:
            return {
                "success": False,
                "result": None,
                "error": "Пустой документ",
                "pages_processed": 0
            }
        
        sample_content = "\n\n".join(pages[:3])
        prompt = get_batch_tables_prompt(1, min(3, len(pages)), sample_content)
        
        response_text = call_openrouter_sync(
            prompt=prompt,
            model=model,
            api_key=api_key,
            temperature=temperature,
            response_format="json_object"
        )
        
        entities = parse_llm_json_response(response_text)
        if not entities or not isinstance(entities, list):
            return {
                "success": False,
                "result": None,
                "error": "Не удалось извлечь характеристики",
                "pages_processed": len(pages)
            }
        
        summary_prompt = get_tables_summary_prompt()
        summary_response = call_openrouter_sync(
            prompt=f"{summary_prompt}\n\nДанные для анализа:\n{json.dumps(entities, ensure_ascii=False)}",
            model=model,
            api_key=api_key,
            temperature=temperature,
            response_format="json_object"
        )
        
        summary = parse_llm_json_response(summary_response)
        
        return {
            "success": True,
            "result": {
                "raw_entities": entities,
                "summary_table": summary
            },
            "error": None,
            "pages_processed": len(pages)
        }
        
    except Exception as e:
        logger.error(f"Ошибка в extract_tables_summary: {e}", exc_info=True)
        return {
            "success": False,
            "result": None,
            "error": str(e),
            "pages_processed": 0
        }
    