import re
from typing import Dict
import pandas as pd

UNITS_MAP: Dict[str, str] = {
    'метр': 'м', 'метры': 'м', 'метров': 'м',
    'миллиметр': 'мм', 'миллиметры': 'мм', 'миллиметров': 'мм',
    'сантиметр': 'см', 'сантиметры': 'см', 'сантиметров': 'см',
    'километр': 'км', 'километры': 'км', 'километров': 'км',
    'килограмм': 'кг', 'килограммы': 'кг', 'килограммов': 'кг',
    'грамм': 'г', 'граммы': 'г', 'граммов': 'г',
    'тонна': 'т', 'тонны': 'т', 'тонн': 'т',
    'ампер': 'А', 'ампера': 'А', 'амперов': 'А',
    'вольт': 'В', 'вольта': 'В',
    'киловатт': 'кВт', 'киловатты': 'кВт', 'киловаттов': 'кВт',
    'штука': 'шт', 'штуки': 'шт', 'штук': 'шт',
    'комплект': 'компл', 'комплекты': 'компл', 'комплектов': 'компл',
    'паскаль': 'Па', 'паскаля': 'Па', 'паскалей': 'Па',
    'мегапаскаль': 'МПа', 'мегапаскаля': 'МПа', 'мегапаскалей': 'МПа',
    'бар': 'бар', 'бары': 'бар',
    'час': 'ч', 'часа': 'ч', 'часов': 'ч',
    'минута': 'мин', 'минуты': 'мин', 'минут': 'мин',
    'секунда': 'с', 'секунды': 'с', 'секунд': 'с',
}

ATTR_ABBREV: Dict[str, str] = {
    'длина': 'L', 'length': 'L',
    'ширина': 'B', 'width': 'B',
    'высота': 'H', 'глубина': 'H', 'height': 'H', 'depth': 'H',
    'толщина': 's', 'thickness': 's',
    'диаметр': 'D', 'diameter': 'D',
    'радиус': 'R', 'radius': 'R',
    'межосевое расстояние': 'A', 'center distance': 'A',
    'шаг резьбы': 't', 'thread pitch': 't',
    'площадь': 'S', 'area': 'S',
}

LATIN_TO_CYRILLIC: Dict[str, str] = {
    'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н',
    'K': 'К', 'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т',
    'X': 'Х', 'I': 'І'
}


def normalize_units_and_attrs(text) -> str:
    """Нормализует единицы измерения и характеристики текста согласно правилам ГОСТ."""
    result = str(text)

    sorted_units = sorted(UNITS_MAP.keys(), key=len, reverse=True)
    for unit_full in sorted_units:
        pattern = re.compile(r'\b' + re.escape(unit_full) + r'\b', re.IGNORECASE)
        result = pattern.sub(UNITS_MAP[unit_full], result)

    sorted_attrs = sorted(ATTR_ABBREV.keys(), key=len, reverse=True)
    for attr_full in sorted_attrs:
        pattern = re.compile(r'\b' + re.escape(attr_full) + r'\b', re.IGNORECASE)
        result = pattern.sub(ATTR_ABBREV[attr_full], result)

    result = re.sub(r'\bд\.\s?н\b', 'DN', result, flags=re.IGNORECASE)
    result = re.sub(r'\bп\.\s?н\b', 'PN', result, flags=re.IGNORECASE)
    result = re.sub(r'\bdn\s*(\d+)', r'DN\1', result, flags=re.IGNORECASE)
    result = re.sub(r'\bpn\s*(\d+)', r'PN\1', result, flags=re.IGNORECASE)
    result = re.sub(r'(\d+\.?\d*)\s*[\.\.]+\s*(\d+\.?\d*)', r'\1..\2', result)
    result = re.sub(r'(\d+)\s*/\s*(\d+)', r'\1/\2', result)

    return result


def normalize_model_name(text) -> str:
    """Нормализует наименование модели согласно 13 правилам ГОСТ."""
    if pd.isna(text) or not isinstance(text, str):
        return text

    s = text.strip().upper().replace('Ё', 'Е')

    normalized_chars = []
    for char in s:
        normalized_chars.append(LATIN_TO_CYRILLIC.get(char, char))
    s = ''.join(normalized_chars)
    
    s = re.sub(r'(\d)\s*,\s*(\d)', r'\1.\2', s)
    s = re.sub(r'[xхХ]', '-', s)

    match_no = re.search(r'(?:№|NO\.?|N)\s*(\d+)', s)
    suffix_no = f"({match_no.group(1)})" if match_no else ""
    if match_no:
        s = s[:match_no.start()] + s[match_no.end():]
    
    s = re.sub(r'(\d)\s+([\s_/\-]*)\s*(\d)', r'\1-\3', s)
    s = re.sub(r'(\d)[_/](\d)', r'\1-\2', s)
    s = re.sub(r'\s+', '', s)
    s = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', s)
    s = re.sub(r'([A-Za-zА-Яа-я])[-_/.]+(\d)', r'\1\2', s)
    s = re.sub(r'([A-Za-zА-Яа-я])[-_/.]+(?=[A-Za-zА-Яа-я])', r'\1-', s)
    s = re.sub(r'\.+', '.', s)
    s = re.sub(r'-+', '-', s)
    s = s.strip('-')
    
    return s + suffix_no if suffix_no else s


def smart_normalize(text) -> str:
    """Комбинированная нормализация текста: единицы, атрибуты и наименование модели."""
    if pd.isna(text):
        return text
    result = normalize_units_and_attrs(text)
    return normalize_model_name(result)
