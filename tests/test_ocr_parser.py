
import pytest
from main import extrair_dados_completos, eh_aso

# Mock class for Image (since we won't process real images in unit tests, 
# but we need to pass something to the function if we were not mocking the inner calls.
# However, since we mock pytesseract, the image object can be anything.)
class MockImage:
    pass

@pytest.fixture
def mock_pytesseract(mocker):
    return mocker.patch('main.pytesseract')

def test_eh_aso_positivo():
    texto = "Este é um Atestado de Saúde Ocupacional oficial."
    assert eh_aso(texto) is True

    texto_caps = "NUMERO 123 - ASO - ADMISSIONAL"
    assert eh_aso(texto_caps) is True

def test_eh_aso_negativo():
    texto = "Este é apenas um recibo de pagamento."
    assert eh_aso(texto) is False

def test_extrair_dados_completos_sucesso(mock_pytesseract):
    # Simula o retorno do OCR (image_to_string)
    ocr_output = """
    CLINICA MEDICA
    ATESTADO DE SAUDE OCUPACIONAL - ASO
    
    Nome Completo: JOAO DA SILVA
    CPF: 123.456.789-00
    Data do Exame: 15/05/2024
    
    Função: ELETRICISTA DE MANUTENCAO
    Setor: OPERACIONAL
    
    Riscos: Ruido, Altura.
    """
    mock_pytesseract.image_to_string.return_value = ocr_output

    nome, cpf, data, funcao, _ = extrair_dados_completos(MockImage())

    assert nome == "JOAO DA SILVA"
    assert cpf == "123.456.789-00"
    assert data == "15/05/2024"
    assert funcao == "ELETRICISTA DE MANUTENCAO"

def test_extrair_dados_completos_rascunho(mock_pytesseract):
    ocr_output = """
    RASCUNHO RASCUNHO RASCUNHO
    RASCUNHO
    Nome: TESTE
    """
    mock_pytesseract.image_to_string.return_value = ocr_output

    nome, cpf, data, funcao, _ = extrair_dados_completos(MockImage())

    assert nome == "RASCUNHO"
    assert cpf == "Ignorar"

def test_extrair_dados_completos_heuristica_nome_cpf(mock_pytesseract):
    # Teste para quando não tem "Nome:" explícito, mas está perto do CPF
    ocr_output = """
    RELATORIO MEDICO
    
    MARIA OLIVEIRA SOUZA
    CPF: 987.654.321-99
    
    Cargo: APRENDIZ
    """
    mock_pytesseract.image_to_string.return_value = ocr_output

    nome, cpf, data, funcao, _ = extrair_dados_completos(MockImage())

    assert nome == "MARIA OLIVEIRA SOUZA"
    assert cpf == "987.654.321-99"
    assert funcao == "APRENDIZ"

def test_extrair_dados_completos_cpf_sem_pontuacao(mock_pytesseract):
    ocr_output = """
    Nome: JOSE PEREIRA
    CPF 11122233344
    Função: PEDREIRO
    """
    mock_pytesseract.image_to_string.return_value = ocr_output

    nome, cpf, data, funcao, _ = extrair_dados_completos(MockImage())

    assert nome == "JOSE PEREIRA"
    assert cpf == "111.222.333-44" # Verifica formatação
    assert funcao == "PEDREIRO"
