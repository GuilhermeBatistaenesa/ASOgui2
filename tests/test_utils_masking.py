from utils_masking import mask_cpf, mask_cpf_in_text, mask_pii_in_obj


def test_mask_cpf_formats_digits():
    assert mask_cpf("12345678900") == "***.***.**9-00"


def test_mask_cpf_in_text_masks_both_formats():
    text = "CPF 12345678900 and 111.222.333-44"
    masked = mask_cpf_in_text(text, keep_last=2)
    assert "***.***.***-00" in masked
    assert "***.***.***-44" in masked


def test_mask_pii_in_obj_recursive():
    obj = {
        "cpf": "12345678900",
        "items": ["CPF 11122233344", {"raw": "222.333.444-55"}],
    }
    masked = mask_pii_in_obj(obj, keep_last=2)

    assert masked["cpf"].endswith("-00")
    assert "***.***.***-44" in masked["items"][0]
    assert "***.***.***-55" in masked["items"][1]["raw"]
