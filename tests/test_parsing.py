import pytest
from pathlib import Path

PDF_FILE = Path(__file__).parent.parent / "Паспорт Чиллер Delta моноблок с конденсатром воздушного охлаждения-ДХ.pdf"

@pytest.mark.skipif(not PDF_FILE.exists(), reason="Тестовый PDF не найден в корне проекта")
def test_parse_pdf_returns_correct_structure(client):
    """Проверяем, что при загрузке реального PDF возвращается успешный ответ с текстом и таблицами."""
    with open(PDF_FILE, "rb") as f:
        response = client.post("/pdf_parser/api/parse", files={"file": ("chiller.pdf", f, "application/pdf")})
    assert response.status_code == 200, f"Ожидался 200, получен {response.status_code}: {response.text}"
    data = response.json()
    assert data["success"] is True
    assert "data" in data
    result = data["data"]
    assert "pages_text" in result
    assert "tables" in result
    assert isinstance(result["pages_text"], list)
    assert len(result["pages_text"]) > 0, "Должен быть извлечён хотя бы один фрагмент текста"
    assert isinstance(result["tables"], list)

def test_parse_pdf_tables_have_header(client):
    """Проверяем, что если есть таблицы, то у каждой есть заголовок (первая строка не пуста)."""
    if not PDF_FILE.exists():
        pytest.skip("PDF файл отсутствует")
    with open(PDF_FILE, "rb") as f:
        response = client.post("/pdf_parser/api/parse", files={"file": ("chiller.pdf", f, "application/pdf")})
    data = response.json()
    tables = data["data"]["tables"]
    for i, table in enumerate(tables):
        assert len(table) > 0, f"Таблица {i+1} пуста"
        assert len(table[0]) > 0, f"У таблицы {i+1} нет строк"

def test_parse_pdf_valid_xlsx_generation(client):
    """Проверяем, что при наличии таблиц генерируется корректный xlsx_file (base64)."""
    if not PDF_FILE.exists():
        pytest.skip("PDF файл отсутствует")
    with open(PDF_FILE, "rb") as f:
        response = client.post("/pdf_parser/api/parse", files={"file": ("chiller.pdf", f, "application/pdf")})
    data = response.json()
    tables = data["data"]["tables"]
    if tables:
        assert "xlsx_file" in data, "При наличии таблиц должен быть xlsx_file"
        assert data["xlsx_file"] is not None
        import base64
        try:
            decoded = base64.b64decode(data["xlsx_file"])
            assert len(decoded) > 0
        except Exception:
            pytest.fail("xlsx_file не является валидной base64-строкой")

def test_parse_docx_not_supported_without_file(client):
    """Проверяем, что загрузка не-PDF/DOCX возвращает ошибку 400."""
    response = client.post("/pdf_parser/api/parse", files={"file": ("test.txt", b"hello", "text/plain")})
    assert response.status_code == 400, f"Ожидался 400, получен {response.status_code}"
    detail = response.json().get("detail", "")
    # Сравнение без учёта регистра
    assert "поддерживаются только pdf и docx" in detail.lower()