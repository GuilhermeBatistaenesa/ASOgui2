from __future__ import annotations

from pathlib import Path

import notification
from outcomes import SKIPPED_NO_RECIPIENT


class _DummyAttachments:
    def __init__(self):
        self.items = []

    def Add(self, Source):
        self.items.append(Source)


class _DummyMail:
    def __init__(self):
        self.To = ""
        self.Subject = ""
        self.HTMLBody = ""
        self.SentOnBehalfOfName = ""
        self.Attachments = _DummyAttachments()
        self.sent = False

    def Send(self):
        self.sent = True


class _DummyOutlook:
    def __init__(self):
        self.last_mail = None

    def CreateItem(self, _):
        self.last_mail = _DummyMail()
        return self.last_mail


def test_enviar_resumo_email_skips_without_recipient(monkeypatch):
    monkeypatch.delenv("ASO_NOTIFY_TO", raising=False)
    monkeypatch.delenv("ASO_EMAIL_TO", raising=False)

    status, msg = notification.enviar_resumo_email(
        destinatario="",
        relatorio={"erros": []},
        execution_id="exec-1",
        run_status="OK",
        report_paths=None,
        manifest_path=None,
        logger=None,
    )
    assert status == SKIPPED_NO_RECIPIENT
    assert "nao configurado" in msg.lower()


def test_enviar_resumo_email_sends_with_attachments(monkeypatch, tmp_path):
    dummy_outlook = _DummyOutlook()

    def _dispatch(_):
        return dummy_outlook

    monkeypatch.setattr(notification.win32, "Dispatch", _dispatch)
    monkeypatch.setenv("ASO_NOTIFY_TO", "dest@empresa.com.br")

    report_json = Path(tmp_path) / "relatorio.json"
    report_md = Path(tmp_path) / "resumo.md"
    manifest = Path(tmp_path) / "manifest.json"
    report_json.write_text("{}", encoding="utf-8")
    report_md.write_text("# ok", encoding="utf-8")
    manifest.write_text("{}", encoding="utf-8")

    status, err = notification.enviar_resumo_email(
        destinatario="",
        relatorio={"erros": [], "tempo_total": "00:00:01", "total_detected": 1, "total_processed": 1, "success": 1, "error": 0},
        execution_id="exec-2",
        run_status="OK",
        report_paths={"json": str(report_json), "md": str(report_md)},
        manifest_path=str(manifest),
        logger=None,
    )

    assert status == "SENT"
    assert err is None
    assert dummy_outlook.last_mail is not None
    assert dummy_outlook.last_mail.sent is True
    assert len(dummy_outlook.last_mail.Attachments.items) == 3
