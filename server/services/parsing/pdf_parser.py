"""
Парсинг PDF: извлечение текста и таблиц с объединением по страницам.
Использует библиотеки fitz (PyMuPDF) и pdfplumber.
"""

import fitz
import pdfplumber
from typing import List, Dict, Any, Tuple
import io


def extract_text_by_pages(pdf_bytes: bytes) -> List[str]:
    """
    Извлекает текст из PDF тремя способами по очереди:
    1. PyMuPDF (fitz) — для цифровых PDF.
    2. pdfplumber — если fitz не дал результата.
    3. OCR (pytesseract) — для отсканированных документов.
    Возвращает список текстов страниц.
    """
    pages = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            text = page.get_text("text")
            pages.append(text.strip() if text else "")

    if not any(p for p in pages):
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                pages = []
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    pages.append(text.strip())
        except Exception as e:
            pass

    if not any(p for p in pages):
        try:
            from pdf2image import convert_from_bytes
            import pytesseract

            images = convert_from_bytes(pdf_bytes, dpi=300)
            pages = []
            for i, img in enumerate(images):
                text = pytesseract.image_to_string(img, lang='rus+eng')
                pages.append(text.strip())
        except ImportError as e:
            pass
        except Exception as e:
            pass

    return pages


def _normalize_cells(row: List[str]) -> List[str]:
    return [cell.strip().lower() for cell in row]


def _merge_on_page(
        tables_with_coords: List[Tuple[List[List[str]], Tuple[float, float, float, float]]]
) -> List[List[List[str]]]:
    """
    Слияние таблиц на одной странице, если они расположены близко и имеют совместимые колонки.
    Разрыв до 70pt, сравнение с учётом нормализации заголовков.
    """
    if not tables_with_coords:
        return []
    tables_with_coords.sort(key=lambda t: (t[1][1], t[1][0]))
    merged = []
    current_table = tables_with_coords[0][0]
    current_bbox = tables_with_coords[0][1]

    for table, bbox in tables_with_coords[1:]:
        while table and all(cell.strip() == "" for cell in table[0]):
            table.pop(0)
        if not table:
            continue

        cur_cols = len(current_table[0]) if current_table else 0
        new_cols = len(table[0]) if table else 0

        if cur_cols <= 1 or new_cols <= 1:
            merged.append(current_table)
            current_table = table
            current_bbox = bbox
            continue

        if abs(new_cols - cur_cols) > 1:
            merged.append(current_table)
            current_table = table
            current_bbox = bbox
            continue

        gap = bbox[1] - current_bbox[3]
        if gap <= 70:
            max_cols = max(cur_cols, new_cols)
            for row in current_table:
                while len(row) < max_cols:
                    row.append("")
            for row in table:
                while len(row) < max_cols:
                    row.append("")

            if _normalize_cells(current_table[0]) == _normalize_cells(table[0]):
                current_table.extend(table[1:])
            else:
                current_table.extend(table)
            current_bbox = (
                min(current_bbox[0], bbox[0]),
                min(current_bbox[1], bbox[1]),
                max(current_bbox[2], bbox[2]),
                max(current_bbox[3], bbox[3]),
            )
        else:
            merged.append(current_table)
            current_table = table
            current_bbox = bbox

    if current_table:
        merged.append(current_table)
    return merged


def extract_tables_with_pdfplumber(pdf_bytes: bytes) -> List[List[List[str]]]:
    """Извлечение таблиц из PDF с межстраничным объединением."""
    raw_tables: List[Tuple[List[List[str]], Tuple[float, float, float, float], int]] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.find_tables()
            for table in tables:
                data = table.extract()
                cleaned = [[cell or "" for cell in row] for row in data]
                raw_tables.append((cleaned, table.bbox, page_num))

    page_tables: Dict[int, List[Tuple[List[List[str]], Tuple[float, float, float, float]]]] = {}
    for tbl, bbox, page in raw_tables:
        page_tables.setdefault(page, []).append((tbl, bbox))
    merged_across_pages = []
    for page in sorted(page_tables):
        merged = _merge_on_page(page_tables[page])
        merged_across_pages.extend(merged)

    if not merged_across_pages:
        return []

    final = [merged_across_pages[0]]
    for next_table in merged_across_pages[1:]:
        prev_table = final[-1]
        if len(prev_table[0]) != len(next_table[0]):
            final.append(next_table)
            continue
        if _normalize_cells(prev_table[0]) == _normalize_cells(next_table[0]):
            final.append(next_table)
            continue
        if _normalize_cells(prev_table[-1]) == _normalize_cells(next_table[0]):
            dummy = next_table[0]
            if len(next_table) > 1 and _normalize_cells(next_table[1]) != _normalize_cells(dummy):
                prev_table.extend(next_table[1:])
            else:
                prev_table.extend(next_table)
        else:
            prev_table.extend(next_table)

    cleaned_final = []
    for table in final:
        if not table:
            continue
        table = [row for row in table if any(cell.strip() for cell in row)]
        if not table:
            continue
        col_count = len(table[0])
        valid_cols = []
        for c in range(col_count):
            if any(row[c].strip() if c < len(row) else False for row in table):
                valid_cols.append(c)
        if valid_cols:
            table = [[row[c] if c < len(row) else "" for c in valid_cols] for row in table]
        cleaned_final.append(table)
    return cleaned_final


def parse_pdf_bytes(pdf_bytes: bytes) -> Dict[str, Any]:
    pages_text = extract_text_by_pages(pdf_bytes)
    tables = extract_tables_with_pdfplumber(pdf_bytes)

    return {
        "pages_text": pages_text,
        "tables": tables,
        "page_count": len(pages_text),
        "has_tables": len(tables) > 0
    }
