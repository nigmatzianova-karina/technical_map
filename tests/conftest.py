import sys
from pathlib import Path

# Добавляем корень проекта в sys.path, чтобы работали импорты server.*
root_dir = Path(__file__).parent.parent
sys.path.insert(0, str(root_dir))

import pytest
from fastapi.testclient import TestClient

# Импортируем приложение FastAPI из app.py
# Предполагаем, что app.py лежит в корне и называется app:app
from app import app as test_app

@pytest.fixture
def client():
    return TestClient(test_app)