"""
Роуты для генерации технологических карт и чата.
"""

import base64, json
from typing import Optional, List, Dict
from fastapi import APIRouter, File, Form, UploadFile, HTTPException
from server.core.config import load_settings
from server.services.filework import create_xlsx
from server.services.parsing.pdf_parser import parse_pdf_bytes
from server.services.llm.client import call_openrouter_async
from server.services.llm.response_parser import parse_tech_card_response
import logging

logger = logging.getLogger(__name__)


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
    logger.info("🚀 ЗАПРОС НА ГЕНЕРАЦИЮ ПОЛУЧЕН")
    logger.info(f"📦 Параметры: model={model_name}, class={equipment_class}, file={file.filename if file else None}")
    if not api_key.strip():
        raise HTTPException(status_code=400, detail="API ключ обязателен")

    if not master_prompt:
        settings = load_settings()
        master_prompt = settings.get("master_prompt", "")

    format_instruction = (
        "Ответ должен быть строго JSON-объектом с полями:\n"
        "{\n"
        '  "ТЕКСТ_ОТВЕТ": "текстовое описание",\n'
        '  "ТАБЛИЦА": [\n'
        '    ["Заголовок1", "Заголовок2", ...],\n'
        '    ["Значение1", "Значение2", ...],\n'
        '    ...\n'
        '  ]\n'
        "}\n\n"
        "ТАБЛИЦА — это массив строк. Первая строка — заголовки (ровно те, что в макете!). "
        "Каждая следующая — значения через запятую внутри строки. "
        "Не используй символ '|' в значениях. "
        "Заполняй ВСЕ ячейки осмысленно, не оставляй пустых. "
        "Для пустых данных пиши прочерк '-'. "
        "Убедись, что JSON валидный и не содержит лишних символов.\n\n"
        
        # 🔥 ВАЖНО: пример только про СТРУКТУРУ, не про содержание!
        "ПРИМЕР ФОРМАТА (НЕ КОПИРУЙ СОДЕРЖАНИЕ!):\n"
        '{"ТЕКСТ_ОТВЕТ": "Краткое описание", "ТАБЛИЦА": [\n'
        '["Элемент","Подэлемент","Наименование операции","Краткое содержание работ","Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей","Профессия/Квалификация","Трудоёмкость, человеко/часов","Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ","Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"],\n'
        '["[Элемент из документа]","[Подэлемент из документа]","[Операция]","[Описание]","[ТО-1/ТО-2/СР/КР]","[число + единица из документа]","[число]","[число]","[профессия]","[расчёт]","[ТМЦ из документа]","[число]","[ед.изм.]","[инструмент]","[СИЗ]","[требование безопасности]"]\n'
        ']}\n\n'
        
        "❗ КРИТИЧЕСКИ: Все значения в ТАБЛИЦЕ бери ТОЛЬКО из текста в <TECH_PASSPORT>. "
        "Пример выше — только для понимания структуры JSON. Не копируй 'Система смазки', '2160', '10W-40' и другие значения из примера!\n"
        "Если в документе периодичность указана в месяцах — пиши '12 месяцев', а не '2160 часов'.\n"
        "Если в документе нет данных для ячейки — ставь '-'.\n\n"
        "Создай не более 15 строк данных."
        "❗ ФИНАЛЬНАЯ ПРОВЕРКА ПЕРЕД ОТВЕТОМ:\n"
        "1. Все значения в таблице взяты из <TECH_PASSPORT>?\n"
        "2. Периодичность указана в тех же единицах, что в документе (месяцы/часы)?\n"
        "3. Названия операций совпадают с формулировками документа?\n"
        "Если хоть один ответ 'нет' — перепиши строки.\n"
    )

    # 🔥 Читаем файл ПЕРВЫМ делом
    file_text = ""
    if file:
        file_bytes = await file.read()
        filename = file.filename or ""
        
        if filename.lower().endswith(('.md', '.txt')):
            file_text = file_bytes.decode('utf-8', errors='replace')
        else:
            parsed = parse_pdf_bytes(file_bytes)
            file_text = "\n\n---\n\n".join(p.strip() for p in parsed["pages_text"] if p.strip())
        
        # Безопасный лимит: 80 000 символов ≈ 20 000 токенов
        MAX_CHARS = 80_000
        if len(file_text) > MAX_CHARS:
            half = MAX_CHARS // 2
            file_text = (
                file_text[:half] + 
                "\n\n...[файл обрезан для соблюдения лимита контекста, основные данные сохранены]...\n\n" + 
                file_text[-half:]
            )
        
        logger.info(f"📄 Файл {filename}: {len(file_text)} символов отправлено в промт (~{len(file_text)//4} токенов)")
        
        # Проверка ключевых слов для отладки
        keywords = ["ремонт", "то-", "обслуживание", "смазка", "периодичность"]
        found = [kw for kw in keywords if kw in file_text.lower()]
        if not found:
            logger.warning("⚠️ В извлечённом тексте НЕТ ключевых разделов ТОиР!")
        else:
            logger.info(f"✅ Найдены разделы: {', '.join(found)}")

    # 🔥 ЕДИНСТВЕННАЯ сборка промта (без дублирования!)
    full_prompt = (
        f"{master_prompt}\n\n"
        f"{format_instruction}\n\n"
        f"Модель: {model_name}\nКласс: {equipment_class}\nПодкласс: {subclass}\n\n"
    )
    
    if file_text.strip():
        full_prompt += (
            "<TECH_PASSPORT>\n"
            "НИЖЕ ПРЕДСТАВЛЕН ТЕКСТ ТЕХНИЧЕСКОГО ПАСПОРТА. ВСЕ ДАННЫЕ ДЛЯ КАРТЫ БРАТЬ СТРОГО ОТСЮДА.\n"
            "Особое внимание удели разделам: «Техническое обслуживание», «Ремонт», «Смазка», «Периодичность».\n"
            f"{file_text}\n"
            "</TECH_PASSPORT>\n\n"
        )
    
    full_prompt += "Сформируй технологическую карту в указанном JSON-формате."
    full_prompt += (
    "\n\n[ПРОВЕРКА] Перед генерацией таблицы ответь кратко:\n"
    "1. Какие операции ТО указаны для 'Станина' в документе?\n"
    "2. Есть ли в файле нормы времени для этих операций?\n"
    "3. Какие ТМЦ упоминаются в разделе 'Смазка'?\n"
    "Затем формируй таблицу."
)

    try:
        response_text = await call_openrouter_async(
            prompt=full_prompt,
            model=model,
            api_key=api_key,
            temperature=temperature,
            max_tokens=max_tokens,
            response_format="json_object"
        )
            # 🔥 Отладка: покажите, что вернула модель
        logger.info(f"🤖 RAW RESPONSE (first 300 chars): {response_text[:300]}")
        
        text_desc, rows = parse_tech_card_response(response_text)
        
        # 🔥 Отладка: сколько строк распарсилось
        logger.info(f"📊 Распарсено строк таблицы: {len(rows)}")
        if rows:
            logger.info(f"📋 Первая строка: {rows[0]}")
    except Exception as e:
        logger.error(f"Ошибка вызова LLM: {e}")
        raise HTTPException(status_code=500, detail=str(e))

    text_desc, rows = parse_tech_card_response(response_text)

    # Нормализация строк таблицы
    expected_keys = [
        "Элемент","Подэлемент","Наименование операции","Краткое содержание работ",
        "Вид ТОиР","Периодичность","Норма времени, часов","Количество исполнителей",
        "Профессия/Квалификация","Трудоёмкость, человеко/часов",
        "Наименование ТМЦ","Количество ТМЦ","Единицы измерения ТМЦ",
        "Наименование инструмента","Средства индивидуальной защиты","Требования по безопасности"
    ]
    normalized_rows = []
    for row in rows:
        clean_row = {}
        for k in expected_keys:
            matching_key = next((key for key in row if key.strip().lower() == k.lower()), None)
            val = row.get(matching_key) if matching_key else row.get(k, '')
            if not val or str(val).strip() == '':
                val = '-'
            clean_row[k] = str(val).strip()
        normalized_rows.append(clean_row)
    rows = normalized_rows

    # Убираем проблемную конверсию latin1→utf8 (может ломать кириллицу)
    # Если модель возвращает корректный UTF-8, этот блок не нужен
    # Если есть проблемы с кодировкой — лучше фиксить на уровне парсинга ответа

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
    