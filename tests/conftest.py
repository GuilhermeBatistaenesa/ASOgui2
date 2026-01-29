
import sys
import os

# Adiciona o diretório raiz ao PYTHONPATH para importar os módulos
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Mock items to avoid import errors if not installed/configured in test env
from unittest.mock import MagicMock
modules_to_mock = [
    'win32com', 'win32com.client', 
    'dotenv', 
    'rpa_yube', 
    'pytesseract',
    'pdf2image'
]

for mod in modules_to_mock:
    if mod not in sys.modules:
        sys.modules[mod] = MagicMock()
