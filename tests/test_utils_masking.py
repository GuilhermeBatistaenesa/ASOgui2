from utils_masking import mask_cpf, mask_cpf_in_text, mask_pii_in_obj


def test_mask_cpf_basic():
    masked = mask_cpf("12345678901", keep_last=3, mask_char="*")
    assert masked == "***.***.**9-01"


def test_mask_cpf_invalid_returns_original():
    assert mask_cpf("123") == "123"


def test_mask_cpf_in_text_masks_both_formats():
    text = "CPF 12345678901 e 123.456.789-01 no mesmo texto."
    masked = mask_cpf_in_text(text, keep_last=2, mask_char="X")
    assert "12345678901" not in masked
    assert "123.456.789-01" not in masked
    assert "XXX.XXX.XXX-01" in masked


def test_mask_pii_in_obj_nested():
    obj = {
        "cpf": "12345678901",
        "items": [{"cpf": "123.456.789-01"}],
        "msg": "CPF=12345678901",
    }
    masked = mask_pii_in_obj(obj, keep_last=3, mask_char="*")
    flat = str(masked)
    assert "12345678901" not in flat
    assert "123.456.789-01" not in flat
