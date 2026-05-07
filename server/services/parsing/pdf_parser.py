import fitz
import pdfplumber
import logging
from typing import List, Dict, Any, Tuple
import io

logger = logging.getLogger(__name__)

def extract_text_by_pages(pdf_bytes: bytes) -> List[str]:
    pages = []
    with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
        for page in doc:
            text = page.get_text("text")
            pages.append(text.strip())
    return pages

def _normalize_cells(row: List[str]) -> List[str]:
    return [cell.strip().lower() for cell in row]

def _merge_on_page(
    tables_with_coords: List[Tuple[List[List[str]], Tuple[float, float, float, float]]]
) -> List[List[List[str]]]:
    """Слияние близких таблиц в пределах одной страницы (разрыв <= 50pt)."""
    if not tables_with_coords:
        return []
    tables_with_coords.sort(key=lambda t: (t[1][1], t[1][0]))
    merged = []
    current_table = tables_with_coords[0][0]
    current_bbox = tables_with_coords[0][1]

    for table, bbox in tables_with_coords[1:]:
        if not current_table or not table:
            current_table = table
            current_bbox = bbox
            continue
        while table and all(cell.strip() == "" for cell in table[0]):
            table.pop(0)
        if not table:
            continue
        cur_cols = len(current_table[0]) if current_table else 0
        new_cols = len(table[0]) if table else 0
        if new_cols != cur_cols:
            merged.append(current_table)
            current_table = table
            current_bbox = bbox
            continue
        gap = bbox[1] - current_bbox[3]
        if gap <= 50:
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
    return final


def extract_tables_with_docling(pdf_bytes: bytes) -> List[str]:
    tables = extract_tables_with_pdfplumber(pdf_bytes)
    markdown_tables = []
    for table in tables:
        if not table:
            continue
        md = "| " + " | ".join(table[0]) + " |\n"
        md += "| " + " | ".join(["---"] * len(table[0])) + " |\n"
        for row in table[1:]:
            md += "| " + " | ".join(row) + " |\n"
        markdown_tables.append(md)
    return markdown_tables

def parse_pdf_bytes(pdf_bytes: bytes) -> Dict[str, Any]:
    pages_text = extract_text_by_pages(pdf_bytes)
    tables = extract_tables_with_pdfplumber(pdf_bytes)
    return {
        "pages_text": pages_text,
        "tables": tables,
        "page_count": len(pages_text),
        "has_tables": len(tables) > 0
    }
