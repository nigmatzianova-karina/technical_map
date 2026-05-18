"""
Роуты для загрузки Excel, нормализации данных и работы с иерархией.
"""

from fastapi import APIRouter, File, UploadFile
from fastapi.responses import HTMLResponse
from pathlib import Path
from pydantic import BaseModel
from typing import Any, Dict, List
import pandas as pd
from server.services.normalize_utils import smart_normalize

router = APIRouter(prefix="/hierarchy")

class NormalizeRequest(BaseModel):
    data: list


@router.post("/uploadfile/")
async def create_upload_file(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        df = pd.read_excel(pd.io.common.BytesIO(contents))
        df = df.fillna("")
        return {"filename": file.filename, "status": "success", "data": df.to_dict(orient="records")}
    except Exception as e:
        return {"filename": file.filename, "status": "error", "message": str(e)}

@router.post("/normalize")
async def normalize_data(request: NormalizeRequest):
    try:
        df = pd.DataFrame(request.data)
        if df.empty:
            return {"status": "error", "message": "Нет данных"}
        df.columns = [str(c).strip() for c in df.columns]
        target_col = None
        for col in df.columns:
            col_lower = col.lower()
            if ('нормализ' in col_lower and 'до' in col_lower) or ('модель' in col_lower and 'до' in col_lower):
                target_col = col
                break
        if not target_col:
            for col in df.columns:
                if 'модель' in col.lower():
                    target_col = col
                    break
        if not target_col:
            return {"status": "error", "message": f"Колонка не найдена. Доступные: {list(df.columns)}"}
        df[target_col] = df[target_col].apply(smart_normalize)
        if target_col.lower() != 'модель':
            df.rename(columns={target_col: 'Модель'}, inplace=True)
        df = df.replace({pd.NA: None, float('nan'): None})
        return {"status": "success", "data": df.to_dict(orient="records"), "message": "Нормализация по ГОСТ выполнена."}
    except Exception as e:
        return {"status": "error", "message": str(e)}