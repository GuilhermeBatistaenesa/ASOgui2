import json
from pathlib import Path
from unittest.mock import MagicMock

import main
from reporting import ReportGenerator
from notification import enviar_resumo_email
from utils_masking import mask_cpf


def test_idempotency_skips_existing_file(tmp_path, mocker):
    class DummyImage:
        def save(self, *args, **kwargs):
            raise AssertionError("save should not be called for duplicate")

    mocker.patch('main.convert_from_path', return_value=[DummyImage()])
    mocker.patch('main.pytesseract').image_to_string.return_value = (
        "ASO\\nNome: TESTE\\nCPF: 12345678900\\n"
    )

    existing = tmp_path / "TESTE - 123.456.789-00.pdf"
    existing.write_bytes(b"dummy")

    stats = {
        "total_detected": 0,
        "total_processed": 0,
        "success": 0,
        "error": 0,
        "skipped_duplicate": 0,
        "skipped_draft": 0,
        "skipped_non_aso": 0,
        "skipped_items": [],
        "erros": [],
    }
    manifest_items = []
    files = []

    main.salvar_paginas_individualmente(
        "dummy.pdf",
        str(tmp_path),
        "123",
        lista_novos_arquivos=files,
        stats=stats,
        manifest_items=manifest_items,
    )

    assert files == []
    assert stats["skipped_duplicate"] == 1
    assert any(i.get("outcome") == "SKIPPED_DUPLICATE" for i in manifest_items)


def test_report_contains_execution_id_and_totals(tmp_path):
    stats = {
        "execution_id": "exec-123",
        "tempo_total": "00:00:10",
        "total_detected": 3,
        "total_processed": 2,
        "success": 1,
        "error": 1,
        "skipped_duplicate": 1,
        "skipped_draft": 0,
        "skipped_non_aso": 1,
        "sucessos": [],
        "erros": [],
        "skipped_items": [],
    }

    rg = ReportGenerator(str(tmp_path))
    paths = rg.save_report(stats)
    report_path = Path(paths["json"])

    data = json.loads(report_path.read_text(encoding="utf-8"))
    assert data["execution_id"] == "exec-123"
    assert data["total_detected"] == 3
    assert data["total_processed"] == 2


def test_email_subject_contains_execution_id_and_status(tmp_path, mocker):
    mail = MagicMock()

    class DummyAttachments:
        def __init__(self):
            self.added = []

        def Add(self, Source):
            self.added.append(Source)

    mail.Attachments = DummyAttachments()

    outlook = MagicMock()
    outlook.CreateItem.return_value = mail
    mocker.patch("notification.win32.Dispatch", return_value=outlook)

    report_json = tmp_path / "relatorio.json"
    report_md = tmp_path / "resumo.md"
    manifest = tmp_path / "manifest.json"
    report_json.write_text("{}", encoding="utf-8")
    report_md.write_text("ok", encoding="utf-8")
    manifest.write_text("{}", encoding="utf-8")

    status, err = enviar_resumo_email(
        "user@company.com",
        {"tempo_total": "00:00:01", "total_detected": 1, "total_processed": 1, "success": 1, "error": 0, "erros": []},
        "exec-999",
        "SUCCESS",
        report_paths={"json": str(report_json), "md": str(report_md)},
        manifest_path=str(manifest),
        logger=None,
    )

    assert status == "SENT"
    assert err is None
    assert "exec=exec-999" in mail.Subject
    assert "SUCCESS" in mail.Subject


def test_mask_cpf():
    assert mask_cpf("12345678900") == "***.***.*90-00"
