import re
import json
from typing import List, Dict, Tuple

def parse_tech_card_response(raw_text: str) -> Tuple[str, List[Dict[str, str]]]:
    """
    Извлекает текстовое описание и строки таблицы из ответа LLM.
    Возвращает (текст_описания, список_строк_таблицы).
    """
    match = re.search(r'\{[^{}]*?"ТЕКСТ_ОТВЕТ"[^{}]*?\}', raw_text, re.DOTALL)
    if match:
        try:
            data = json.loads(match.group())
            text_desc = data.get("ТЕКСТ_ОТВЕТ", "")
            table_str = data.get("ТАБЛИЦА", "")
            if table_str:
                rows = _parse_pipe_table(table_str)
                return text_desc, rows
        except json.JSONDecodeError:
            pass

    match = re.search(r'\{[^{}]*?"text"[^{}]*?\}', raw_text, re.DOTALL)
    if match:
        try:
            data = json.loads(match.group())
            text_desc = data.get("text", "")
            rows = data.get("rows", [])
            if rows:
                return text_desc, rows
        except json.JSONDecodeError:
            pass

    table_match = re.search(r'Элемент\|.*?(?:\n|$)', raw_text, re.DOTALL)
    if table_match:
        table_str = raw_text[table_match.start():]
        rows = _parse_pipe_table(table_str)
        if rows:
            text_desc = raw_text[:table_match.start()].strip()
            return text_desc, rows

    return raw_text.strip(), []

def _parse_pipe_table(table_str: str) -> List[Dict[str, str]]:
    """Парсит строку с pipe-таблицей в список словарей."""
    table_str = table_str.encode().decode('unicode_escape')
    lines = [line.strip() for line in table_str.strip().split('\n') if line.strip()]
    if len(lines) < 2:
        return []
    headers = [h.strip() for h in lines[0].split('|') if h.strip()]
    if not headers:
        return []
    rows = []
    for line in lines[1:]:
        cells = [c.strip() for c in line.split('|')]
        while len(cells) < len(headers):
            cells.append("")
        row = {headers[i]: cells[i] if i < len(cells) else "" for i in range(len(headers))}
        rows.append(row)
    return rows
