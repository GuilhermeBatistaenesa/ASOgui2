from __future__ import annotations

from datetime import datetime, timedelta
from pathlib import Path


class _FakeAttachment:
    def __init__(self, filename: str, content: bytes):
        self.FileName = filename
        self._content = content

    def SaveAsFile(self, path):
        Path(path).write_bytes(self._content)


class _FakeAttachments:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Item(self, i):
        return self._items[i - 1]


class _FakeMessage:
    def __init__(self, subject, received, attachments=None, body="", html_body=""):
        self.Class = 43
        self.Subject = subject
        self.ReceivedTime = received
        self.Body = body
        self.HTMLBody = html_body
        self.Attachments = _FakeAttachments(attachments or [])


class _FakeItems:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Sort(self, *_args, **_kwargs):
        return None

    def Item(self, i):
        return self._items[i - 1]


class _FakeFolder:
    def __init__(self, items):
        self.Items = _FakeItems(items)
        self.FolderPath = "Inbox"


class _FakeFolders:
    def __init__(self, mapping):
        self._mapping = mapping

    def __call__(self, name):
        if name in self._mapping:
            return self._mapping[name]
        raise KeyError(name)


class _FakeAccount:
    def __init__(self, name, inbox):
        self.DisplayName = name
        self.Name = name
        self.Folders = _FakeFolders({"Caixa de Entrada": inbox, "Inbox": inbox})


class _FakeRecipient:
    def __init__(self):
        self.Resolved = False

    def Resolve(self):
        self.Resolved = True


class _FakeNamespace:
    def __init__(self, account, inbox):
        self.Accounts = [account]
        self.Folders = _FakeFolders({account.DisplayName: account, "Aso": account})
        self._inbox = inbox

    def CreateRecipient(self, _smtp):
        return _FakeRecipient()

    def GetSharedDefaultFolder(self, _recip, _):
        return self._inbox


def test_captar_emails_flow(load_main, tmp_path, monkeypatch):
    main = load_main(env={"ASO_EMAIL_ACCOUNT": "aso@enesa.com.br", "ASO_MAILBOX_NAME": "Aso"})

    now = datetime.now()
    subject_ok = "ASO ADMISSIONAL - 123 - 01/02/2025"
    msg_with_attachments = _FakeMessage(
        subject_ok,
        now,
        attachments=[
            _FakeAttachment("a.pdf", b"same"),
            _FakeAttachment("b.pdf", b"same"),  # duplicate by hash
        ],
    )
    msg_with_gdrive = _FakeMessage(
        subject_ok,
        now,
        attachments=[],
        body="https://drive.google.com/file/d/ABCdef12345/view",
    )
    msg_old = _FakeMessage(
        subject_ok,
        now - timedelta(days=10),
        attachments=[_FakeAttachment("c.pdf", b"old")],
    )

    inbox = _FakeFolder([msg_with_attachments, msg_with_gdrive, msg_old])
    account = _FakeAccount("aso@enesa.com.br", inbox)
    namespace = _FakeNamespace(account, inbox)

    monkeypatch.setattr(main, "get_outlook_namespace_robusto", lambda *_: (None, namespace))

    calls = []

    def _fake_salvar(pdf_path, pasta_destino, numero_obra, lista_novos_arquivos=None, stats=None, manifest_items=None):
        calls.append(pdf_path)
        if stats is not None:
            stats["total_detected"] += 1
        if lista_novos_arquivos is not None:
            lista_novos_arquivos.append(str(Path(pasta_destino) / "JOAO - 123.456.789-01.pdf"))
        if manifest_items is not None:
            manifest_items.append(
                {"file_display": "JOAO - ***.***.***-01.pdf", "cpf_masked": "***.***.***-01", "outcome": "SUCCESS", "message": "ok"}
            )

    monkeypatch.setattr(main, "salvar_paginas_individualmente", _fake_salvar)
    monkeypatch.setattr(main, "run_from_main", lambda *_args, **_kwargs: {"sucessos": ["JOAO - 123.456.789-01.pdf"], "erros": []})
    monkeypatch.setattr(main, "enviar_resumo_email", lambda *_args, **_kwargs: ("SENT", None))

    def _fake_download(_gid, dest_dir):
        p = Path(dest_dir) / "ASOS ENESA 1.pdf"
        p.write_bytes(b"pdf")
        return str(p)

    monkeypatch.setattr(main, "download_gdrive_file", _fake_download)

    manifest = {
        "execution_id": "exec-1",
        "started_at": now.isoformat(),
        "finished_at": None,
        "duration_sec": None,
        "run_status": None,
        "paths": {},
        "email_status": None,
        "email_error": None,
        "totals": {},
        "items": [],
    }

    main.captar_emails(limit=10, execution_id="exec-1", started_at=now, manifest=manifest)

    assert len(calls) == 2  # dedup attachments, plus gdrive
    assert manifest["email_status"] == "SENT"
    assert len(manifest["items"]) >= 2
