import json
from pathlib import Path

from reporting import ReportGenerator
from utils_masking import mask_cpf


def test_save_report_masks_cpf(tmp_path):
    stats = {
        "execution_id": "exec-1",
        "cpf": "12345678900",
        "erros": [{"arquivo": "TESTE 11122233344", "erro": "CPF 11122233344"}],
    }
    rg = ReportGenerator(str(tmp_path))
    paths = rg.save_report(stats)

    report_path = Path(paths["json"])
    data = json.loads(report_path.read_text(encoding="utf-8"))
    assert data["cpf"] == "***.***.**9-00"
    masked_111 = mask_cpf("11122233344")
    assert masked_111 in data["erros"][0]["arquivo"]
    assert masked_111 in data["erros"][0]["erro"]


def test_generate_markdown_summary_masks_cpf(tmp_path):
    stats = {
        "execution_id": "exec-2",
        "erros": [{"arquivo": "ARQ 12345678900", "erro": "CPF 12345678900"}],
        "ocr_failures": [{"arquivo": "ARQ 22233344455", "cpf": "22233344455", "nome": "JOAO 22233344455"}],
        "skipped_items": ["skip 11122233344"],
    }
    rg = ReportGenerator(str(tmp_path))
    md_path = rg.generate_markdown_summary(stats, "20260130_120000")

    content = Path(md_path).read_text(encoding="utf-8")
    assert "***.***.**9-00" in content
    assert mask_cpf("22233344455") in content
    assert mask_cpf("11122233344") in content
