import base64
from fastapi import APIRouter, File, UploadFile, HTTPException
from fastapi.responses import HTMLResponse
from pathlib import Path
from server.services.parsing.pdf_parser import parse_pdf_bytes
from server.services.parsing.doc_parser import parse_docx_bytes
from server.services.filework import create_xlsx

router = APIRouter(prefix="/pdf_parser", tags=["pdf_parser"])

ALLOWED_EXTENSIONS = {".pdf", ".docx"}

@router.get("/", response_class=HTMLResponse)
async def pdf_parser_page():
    html_path = Path("client/pdf_parser.html")
    if html_path.exists():
        return html_path.read_text(encoding="utf-8")
    return HTMLResponse("Страница не найдена", 404)

@router.post("/api/parse")
async def parse_document(file: UploadFile = File(...)):
    suffix = Path(file.filename).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Поддерживаются только PDF и DOCX")

    contents = await file.read()

    try:
        if suffix == ".pdf":
            result = parse_pdf_bytes(contents)
        elif suffix == ".docx":
            result = parse_docx_bytes(contents)
        else:
            raise HTTPException(400, "Неподдерживаемый формат")

        xlsx_b64 = None
        xlsx_filename = f"{Path(file.filename).stem}.xlsx"
        tables = result.get("tables", [])
        if tables and len(tables) > 0:
            headers = tables[0][0]
            all_rows = []
            for table in tables:
                all_rows.extend(table[1:])
            xlsx_bytes = create_xlsx(headers, all_rows, "", "", "")
            xlsx_b64 = base64.b64encode(xlsx_bytes).decode("utf-8")

        return {
            "success": True,
            "filename": file.filename,
            "file_type": suffix[1:].upper(),
            "data": result,
            "xlsx_file": xlsx_b64,
            "xlsx_filename": xlsx_filename
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка парсинга: {str(e)}")
    