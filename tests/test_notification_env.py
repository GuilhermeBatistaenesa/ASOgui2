import os

from notification import _get_env_recipients, _get_env_sender, enviar_resumo_email
from outcomes import SKIPPED_NO_RECIPIENT


def test_get_env_recipients_normalizes_separators(monkeypatch):
    monkeypatch.setenv("ASO_NOTIFY_TO", "a@b.com, c@d.com ; e@f.com")
    assert _get_env_recipients() == "a@b.com;c@d.com;e@f.com"


def test_get_env_sender_prefers_from(monkeypatch):
    monkeypatch.setenv("ASO_EMAIL_FROM", "from@company.com")
    monkeypatch.setenv("ASO_EMAIL_ACCOUNT", "account@company.com")
    assert _get_env_sender() == "from@company.com"


def test_enviar_resumo_email_skips_when_no_recipient(monkeypatch):
    monkeypatch.delenv("ASO_NOTIFY_TO", raising=False)
    monkeypatch.delenv("ASO_EMAIL_TO", raising=False)

    status, err = enviar_resumo_email(
        "",
        {"tempo_total": "00:00:01", "total_detected": 0, "total_processed": 0, "success": 0, "error": 0, "erros": []},
        "exec-0",
        "SUCCESS",
        report_paths=None,
        manifest_path=None,
        logger=None,
    )

    assert status == SKIPPED_NO_RECIPIENT
    assert "nao configurado" in err.lower()
