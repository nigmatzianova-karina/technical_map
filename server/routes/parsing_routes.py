from fastapi import APIRouter, UploadFile, File, HTTPException

router = APIRouter(prefix="/api/parsing", tags=["parsing"])

@router.post("/parse-pdf")
async def parse_pdf_endpoint(file: UploadFile = File(...)):
    """
    Парсит PDF файл и возвращает извлечённый текст и таблицы.
    """
    try:
        contents = await file.read()
        
        from server.services.filework import parse_pdf_advanced_async
        
        result = await parse_pdf_advanced_async(contents)
        
        return {
            "success": True,
            "filename": file.filename,
            "pages_processed": result["metadata"]["page_count"],
            "text": result["full_markdown"][:50000],
            "tables": result["tables_markdown"],
            "metadata": result["metadata"]
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка парсинга: {str(e)}")
    