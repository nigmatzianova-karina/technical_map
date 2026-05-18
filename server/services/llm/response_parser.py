"""
Парсинг ответа LLM: извлечение текста и таблицы из JSON или pipe-формата.
"""

import re
import json
from typing import List, Dict, Tuple, Any

def parse_tech_card_response(raw_text: str) -> Tuple[str, List[Dict[str, str]]]:
    """
    Извлекает текст и таблицу из ответа LLM.
    Поддерживает два формата:
    1. JSON с полем "ТАБЛИЦА" как массив массивов (новый рекомендуемый).
    2. JSON с полем "ТАБЛИЦА" как строка с pipe-разделителями (старый, fallback).
    Возвращает (текст_описания, список_словарей_строк).
    """
    # Попытка найти JSON
    json_match = re.search(r'\{.*?"ТЕКСТ_ОТВЕТ".*?\}', raw_text, re.DOTALL)
    if not json_match:
        # Иногда модель оборачивает в ```json...```
        code_match = re.search(r'```json\s*(.*?)\s*```', raw_text, re.DOTALL)
        if code_match:
            json_match = code_match  # чтобы ниже попробовать парсить

    if json_match:
        json_str = json_match.group(1) if json_match.re.groups else json_match.group()
        try:
            data = json.loads(json_str)
        except json.JSONDecodeError:
            # Попробуем исправить обрезанный JSON: добавим недостающие скобки
            if not json_str.strip().endswith('}'):
                json_str += '}'
            # и удалим лишний текст после последней скобки
            json_str = json_str[:json_str.rfind('}')+1]
            try:
                data = json.loads(json_str)
            except json.JSONDecodeError:
                data = None
        if data and isinstance(data, dict):
            text_desc = data.get("ТЕКСТ_ОТВЕТ", "")
            table_data = data.get("ТАБЛИЦА", [])

            if isinstance(table_data, list) and all(isinstance(row, list) for row in table_data):
                # Новый формат: массив массивов
                if len(table_data) > 1:
                    headers = table_data[0]
                    rows = []
                    for row in table_data[1:]:
                        # Дополняем строку до длины заголовков
                        while len(row) < len(headers):
                            row.append("")
                        rows.append({headers[i]: row[i] for i in range(len(headers))})
                    return text_desc, rows
                else:
                    return text_desc, []

            elif isinstance(table_data, str):
                # Старый pipe-формат
                rows = _parse_pipe_table(table_data)
                return text_desc, rows

    # Если JSON не найден, ищем pipe-таблицу в сыром тексте
    table_match = re.search(r'Элемент\|.*', raw_text, re.DOTALL)
    if table_match:
        table_str = table_match.group()
        rows = _parse_pipe_table(table_str)
        text_desc = raw_text[:table_match.start()].strip()
        return text_desc, rows

    return raw_text.strip(), []

def _parse_pipe_table(table_str: str) -> List[Dict[str, str]]:
    """Парсит строку с pipe-таблицей в список словарей."""
    lines = [line.strip() for line in table_str.strip().split('\n') if line.strip()]
    if len(lines) < 2:
        return []
    headers = [h.strip() for h in lines[0].split('|') if h.strip()]
    if not headers:
        return []
    rows = []
    for line in lines[1:]:
        cells = [c.strip() for c in line.split('|')]
        # Дополняем пустыми, если не хватает
        while len(cells) < len(headers):
            cells.append("")
        row = {headers[i]: cells[i] for i in range(len(headers))}
        rows.append(row)
    return rows