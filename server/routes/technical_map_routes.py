import base64, json
from typing import Optional, List, Dict
from fastapi import APIRouter, File, Form, UploadFile, HTTPException
from server.core.config import load_settings
from server.services.filework import create_xlsx
from server.services.parsing.pdf_parser import parse_pdf_bytes
from server.services.llm.client import call_openrouter_async
from server.services.llm.response_parser import parse_tech_card_response

router = APIRouter(prefix="/tech_map", tags=["tech_map"])

CSV_HEADERS = [
    "Элемент","Подэлемент","Наименование операции","Краткое содержание работ",
    "Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей",
    "Профессия/Квалификация","Трудоёмкость, человеко/часов",
    "Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ",
    "Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"
]


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

    try:
        response_text = await call_openrouter_async(
            prompt=full_prompt,
            model=model,
            api_key=api_key,
            temperature=temperature,
            max_tokens=max_tokens,
            response_format="json_object"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    text_desc, rows = parse_tech_card_response(response_text)

    expected_keys = [
        "Элемент","Подэлемент","Наименование операции","Краткое содержание работ",
        "Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей",
        "Профессия/Квалификация","Трудоёмкость, человеко/часов",
        "Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ",
        "Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"
    ]
    fixed_rows = []
    for row in rows:
        fixed_row = {}
        for k, v in row.items():
            try:
                new_key = k.encode('latin1').decode('utf-8')
            except (UnicodeDecodeError, UnicodeEncodeError):
                new_key = k
            try:
                if isinstance(v, str):
                    new_value = v.encode('latin1').decode('utf-8')
                else:
                    new_value = v
            except (UnicodeDecodeError, UnicodeEncodeError):
                new_value = v
            fixed_row[new_key] = new_value
        clean_row = {key: fixed_row.get(key, '') for key in expected_keys}
        fixed_rows.append(clean_row)
    rows = fixed_rows

    xlsx_data = None
    if rows:
        full_rows = [dict(zip(expected_keys, [row.get(h, "") for h in expected_keys])) for row in rows]
        xlsx_bytes = create_xlsx(expected_keys, [list(r.values()) for r in full_rows],
                                 equipment_class, subclass, model_name)
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

    api_key = settings.get("api_key", "")
    if not api_key:
        return {"reply": "API ключ не настроен."}

    try:
        response = await call_openrouter_async(
            prompt=system_msg + "\n" + message,
            model=settings.get("model", "openai/gpt-4o-mini"),
            api_key=api_key,
            temperature=0.3,
            max_tokens=500
        )
        return {"reply": response}
    except Exception as e:
        return {"reply": f"Ошибка: {str(e)}"}
    