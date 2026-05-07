from fastapi import APIRouter, File, UploadFile
from fastapi.responses import HTMLResponse
from pathlib import Path
from pydantic import BaseModel
from typing import Any, Dict, List
import pandas as pd
from server.services.normilaze_utils import smart_normalize

router = APIRouter(prefix="/hierarchy")

class NormalizeRequest(BaseModel):
    data: list

class ClassifyRequest(BaseModel):
    data: List[Dict[str, Any]]
    classifier: List[Dict[str, Any]]

@router.get("/", response_class=HTMLResponse)
async def hierarchy_page():
    html_path = Path("client/hierarchy.html")
    if html_path.exists():
        with open(html_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Поместите hierarchy.html в папку client/</h1>"

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
        return {"status": "success", "data": df.to_dict(orient="records"), "message": "Нормализация выполнена."}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@router.post("/classify")
async def classify_data(request: ClassifyRequest):
    try:
        df = pd.DataFrame(request.data)
        classifier_df = pd.DataFrame(request.classifier)
        class_map = {}
        for _, row in classifier_df.iterrows():
            if 'Модель' in row and not pd.isna(row['Модель']):
                clean_model = str(row['Модель']).strip().lower()
                class_map[clean_model] = {'Класс': row.get('Класс', ''), 'Подкласс': row.get('Подкласс', '')}
        target_column = None
        if 'Модель после нормализации' in df.columns:
            target_column = 'Модель после нормализации'
        elif 'Модель' in df.columns:
            target_column = 'Модель'
        else:
            return {"status": "error", "message": "Колонка не найдена"}
        if 'Класс' not in df.columns: df['Класс'] = None
        if 'Подкласс' not in df.columns: df['Подкласс'] = None
        def apply_classification(row):
            model_value = row[target_column]
            if pd.isna(model_value): return None, None
            clean_key = str(model_value).strip().lower()
            match = class_map.get(clean_key)
            return (match['Класс'], match['Подкласс']) if match else (None, None)
        classified = df.apply(lambda row: apply_classification(row), axis=1, result_type='expand')
        df['Класс'] = classified[0]
        df['Подкласс'] = classified[1]
        return {"status": "success", "data": df.to_dict(orient="records"), "message": f"Классифицировано строк: {df['Класс'].notna().sum()}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}
    