import main


class DummyImage:
    pass


def test_ocr_with_fallback_skips_when_score_high(mocker):
    mocker.patch("main._preprocess_img", return_value=DummyImage())
    image_to_string = mocker.patch("main.pytesseract.image_to_string")
    image_to_string.return_value = "ASO CPF 123.456.789-00 FUNCIONARIO"

    result = main.ocr_with_fallback(DummyImage())

    assert result == "ASO CPF 123.456.789-00 FUNCIONARIO"
    assert image_to_string.call_count == 1


def test_ocr_with_fallback_picks_best_score(mocker):
    mocker.patch("main._preprocess_img", return_value=DummyImage())
    image_to_string = mocker.patch("main.pytesseract.image_to_string")
    image_to_string.side_effect = [
        "texto ruim",
        "CPF 11122233344",
        "ASO SAUDE OCUPACIONAL",
        "CPF 123.456.789-00 ASO",
        "nada",
    ]

    result = main.ocr_with_fallback(DummyImage())

    assert "123.456.789-00" in result
