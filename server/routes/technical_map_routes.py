import json, asyncio, base64, re
from pathlib import Path
from typing import Optional, List, Dict
from fastapi import APIRouter, File, Form, UploadFile, HTTPException
from fastapi.responses import HTMLResponse
from server.utils.settings import load_settings
from server.services.filework import create_xlsx
from server.services.parsing.pdf_parser import parse_pdf_bytes
from server.services.llm.client import call_openrouter_sync

router = APIRouter(prefix="/tech_map", tags=["tech_map"])

CSV_HEADERS = [
    "Элемент","Подэлемент","Наименование операции","Краткое содержание работ",
    "Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей",
    "Профессия/Квалификация","Трудоёмкость, человеко/часов",
    "Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ",
    "Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"
]

@router.get("/", response_class=HTMLResponse)
async def tech_map_page():
    html_path = Path("client/technical_map.html")
    return html_path.read_text(encoding="utf-8") if html_path.exists() else HTMLResponse("Страница не найдена", 404)

@router.post("/api/generate")
async def generate_tech_card(
    model_name: str = Form(...),
    equipment_class: Optional[str] = Form(""),
    subclass: Optional[str] = Form(""),
    file: Optional[UploadFile] = File(None),
    model: str = Form("openai/gpt-4o-mini"),
    api_key: str = Form(...),
    temperature: float = Form(0.3),
    max_tokens: int = Form(3000),
    master_prompt: Optional[str] = Form("")
):
    if not api_key.strip():
        raise HTTPException(status_code=400, detail="API ключ обязателен")
    
    if not master_prompt:
        settings = load_settings()
        master_prompt = settings.get("master_prompt", "")
    format_instruction = (
        "Ответ должен быть строго JSON-объектом с полями:\n"
        "{\n"
        '  "ТЕКСТ_ОТВЕТ": "текстовое описание",\n'
        '  "ТАБЛИЦА": "строка с заголовками и данными через | и \\n"\n'
        "}\n"
        "Не добавляй никаких других ключей. Убедись, что JSON валидный."
    )
    full_prompt = f"{master_prompt}\n\n{format_instruction}\n\n"
    full_prompt += f"Модель: {model_name}\nКласс: {equipment_class}\nПодкласс: {subclass}\nСформируй техкарту."

    file_text = ""
    if file:
        file_bytes = await file.read()
        parsed = parse_pdf_bytes(file_bytes)
        file_text = "\n".join(parsed["pages_text"][:2])
        full_prompt += f"\n\nТехпаспорт:\n{file_text[:3000]}"

    loop = asyncio.get_event_loop()
    try:
        response_text = await loop.run_in_executor(
            None, call_openrouter_sync, full_prompt, model, api_key, temperature, max_tokens, "json_object"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    text_desc = ""
    rows = []
    try:
        data = json.loads(response_text)
        if isinstance(data, dict):
            text_desc = data.get("ТЕКСТ_ОТВЕТ", "")
            table_str = data.get("ТАБЛИЦА", "")
            if table_str:
                rows = _parse_pipe_table(table_str)
    except json.JSONDecodeError:
        json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
        if json_match:
            try:
                data = json.loads(json_match.group())
                text_desc = data.get("ТЕКСТ_ОТВЕТ", "")
                table_str = data.get("ТАБЛИЦА", "")
                if table_str:
                    lines = table_str.split('\n')
                    clean = []
                    for line in lines:
                        if line.count('|') >= 15:
                            clean.append(line)
                    if clean:
                        table_str = '\n'.join(clean)
                    rows = _parse_pipe_table(table_str)
            except:
                pass
        if not rows:
            text_desc = response_text

    xlsx_data = None
    if rows:
        full_rows = [dict(zip(CSV_HEADERS, [row.get(h, "") for h in CSV_HEADERS])) for row in rows]
        xlsx_bytes = create_xlsx(CSV_HEADERS, [list(r.values()) for r in full_rows], equipment_class, subclass, model_name)
        xlsx_data = base64.b64encode(xlsx_bytes).decode()

    return {
        "success": True,
        "data": {
            "text": text_desc,
            "rows": rows,
            "xlsx_file": xlsx_data,
            "xlsx_filename": f"ТК_{model_name}.xlsx"
        }
    }


def _parse_pipe_table(table_str: str) -> List[Dict[str, str]]:
    lines = table_str.strip().split('\n')
    clean_lines = []
    for line in lines:
        if '|' not in line:
            continue
        idx = line.find('{iNT')
        if idx != -1:
            line = line[:idx].strip()
        if '|' in line:
            clean_lines.append(line)
        else:
            pass
    
    if len(clean_lines) < 2:
        return []
    headers = [h.strip() for h in clean_lines[0].split('|') if h.strip()]
    if not headers:
        return []
    rows = []
    for line in clean_lines[1:]:
        cells = [c.strip() for c in line.split('|')]
        while len(cells) < len(headers):
            cells.append("")
        row = {headers[i]: cells[i] if i < len(cells) else "" for i in range(len(headers))}
        rows.append(row)
    return rows

@router.post("/api/chat")
async def chat_endpoint(message: str = Form(...), history: Optional[str] = Form("[]")):
    settings = load_settings()
    system_msg = "Ты – технический эксперт. Отвечай только на вопросы по промышленному оборудованию, ТОиР, технологическим картам. Игнорируй другие запросы."
    messages = [{"role": "system", "content": system_msg}]
    try:
        hist = json.loads(history)
        messages.extend(hist[-5:])
    except:
        pass
    messages.append({"role": "user", "content": message})
    try:
        loop = asyncio.get_event_loop()
        response = await loop.run_in_executor(
            None, lambda: call_openrouter_sync(
                system_msg + "\n" + message,
                settings.get("model", "openai/gpt-4o-mini"),
                settings.get("api_key", ""),
                0.3, 500
            )
        )
        return {"reply": response}
    except Exception as e:
        return {"reply": f"Ошибка: {str(e)}"}
    