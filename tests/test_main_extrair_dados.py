from __future__ import annotations


def test_extrair_dados_completos_basic(load_main):
    main = load_main()
    texto = """
    ATESTADO DE SAUDE OCUPACIONAL
    Nome: JOAO DA SILVA
    CPF: 123.456.789-01
    Data ASO: 01/02/2025
    Funcao: SOLDADOR
    """
    nome, cpf, data_aso, funcao, _ = main.extrair_dados_completos(None, texto_ocr=texto)
    assert nome == "JOAO DA SILVA"
    assert cpf == "123.456.789-01"
    assert data_aso == "01/02/2025"
    assert "SOLDADOR" in funcao


def test_extrair_dados_completos_rascunho(load_main):
    main = load_main()
    texto = "RASCUNHO RASCUNHO RASCUNHO RASCUNHO"
    nome, cpf, data_aso, funcao, _ = main.extrair_dados_completos(None, texto_ocr=texto)
    assert nome == "RASCUNHO"
    assert cpf == "Ignorar"
    assert data_aso == ""
    assert funcao == ""


def test_extrair_dados_completos_nome_fallback(load_main):
    main = load_main()
    texto = """
    JOAO DA SILVA CPF 12345678901
    DATA: 03/04/2025
    """
    nome, cpf, data_aso, funcao, _ = main.extrair_dados_completos(None, texto_ocr=texto)
    assert nome == "JOAO DA SILVA"
    assert cpf == "123.456.789-01"
    assert data_aso == "03/04/2025"
