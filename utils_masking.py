import re


CPF_DIGITS_RE = re.compile(r"\d{11}")
CPF_FORMAT_RE = re.compile(r"\b(\d{3}\.\d{3}\.\d{3}-\d{2})\b")


def _normalize_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def mask_cpf(value: str, keep_last: int = 3, mask_char: str = "*") -> str:
    digits = _normalize_digits(value)
    if len(digits) != 11:
        return value
    if keep_last < 0:
        keep_last = 0
    if keep_last > 11:
        keep_last = 11
    masked_digits = (mask_char * (11 - keep_last)) + digits[-keep_last:]
    return f"{masked_digits[0:3]}.{masked_digits[3:6]}.{masked_digits[6:9]}-{masked_digits[9:11]}"


def mask_cpf_in_text(text: str, keep_last: int = 3, mask_char: str = "*") -> str:
    if not text:
        return text

    def _mask_match(m):
        return mask_cpf(m.group(0), keep_last=keep_last, mask_char=mask_char)

    text = CPF_FORMAT_RE.sub(_mask_match, text)

    def _mask_digits(m):
        return mask_cpf(m.group(0), keep_last=keep_last, mask_char=mask_char)

    return CPF_DIGITS_RE.sub(_mask_digits, text)


def mask_pii_in_obj(obj, keep_last: int = 3, mask_char: str = "*"):
    if isinstance(obj, dict):
        return {k: mask_pii_in_obj(v, keep_last=keep_last, mask_char=mask_char) for k, v in obj.items()}
    if isinstance(obj, list):
        return [mask_pii_in_obj(v, keep_last=keep_last, mask_char=mask_char) for v in obj]
    if isinstance(obj, str):
        return mask_cpf_in_text(obj, keep_last=keep_last, mask_char=mask_char)
    return obj
