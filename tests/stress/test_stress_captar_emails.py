from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import pytest


@pytest.mark.stress
def test_stress_captar_emails(load_main, tmp_path, monkeypatch):
    if os.getenv("RUN_STRESS") != "1":
        pytest.skip("stress test skipped (set RUN_STRESS=1)")

    main = load_main(env={"ASO_EMAIL_ACCOUNT": "aso@enesa.com.br", "ASO_MAILBOX_NAME": "Aso"})

    class _Attachment:
        def __init__(self, filename, content):
            self.FileName = filename
            self._content = content

        def SaveAsFile(self, path):
            Path(path).write_bytes(self._content)

    class _Attachments:
        def __init__(self, items):
            self._items = list(items)

        @property
        def Count(self):
            return len(self._items)

        def Item(self, i):
            return self._items[i - 1]

    class _Msg:
        def __init__(self, subject, received, attachments):
            self.Class = 43
            self.Subject = subject
            self.ReceivedTime = received
            self.Attachments = _Attachments(attachments)
            self.Body = ""
            self.HTMLBody = ""

    class _Items:
        def __init__(self, items):
            self._items = list(items)

        @property
        def Count(self):
            return len(self._items)

        def Sort(self, *_args, **_kwargs):
            return None

        def Item(self, i):
            return self._items[i - 1]

    class _Folder:
        def __init__(self, items):
            self.Items = _Items(items)
            self.FolderPath = "Inbox"

    class _Folders:
        def __init__(self, mapping):
            self._mapping = mapping

        def __call__(self, name):
            return self._mapping[name]

    class _Account:
        def __init__(self, name, inbox):
            self.DisplayName = name
            self.Name = name
            self.Folders = _Folders({"Caixa de Entrada": inbox, "Inbox": inbox})

    class _Recipient:
        def __init__(self):
            self.Resolved = False

        def Resolve(self):
            self.Resolved = True

    class _Namespace:
        def __init__(self, account, inbox):
            self.Accounts = [account]
            self.Folders = _Folders({account.DisplayName: account, "Aso": account})
            self._inbox = inbox

        def CreateRecipient(self, _smtp):
            return _Recipient()

        def GetSharedDefaultFolder(self, _recip, _):
            return self._inbox

    now = datetime.now()
    subject_ok = "ASO ADMISSIONAL - 123 - 01/02/2025"
    msgs = []
    for i in range(300):
        att = _Attachment(f"f{i}.pdf", f"data-{i}".encode())
        msgs.append(_Msg(subject_ok, now, [att]))

    inbox = _Folder(msgs)
    account = _Account("aso@enesa.com.br", inbox)
    namespace = _Namespace(account, inbox)

    monkeypatch.setattr(main, "get_outlook_namespace_robusto", lambda *_: (None, namespace))
    monkeypatch.setattr(main, "salvar_paginas_individualmente", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(main, "run_from_main", lambda *_args, **_kwargs: {"sucessos": [], "erros": []})
    monkeypatch.setattr(main, "enviar_resumo_email", lambda *_args, **_kwargs: ("SKIPPED", None))

    manifest = {
        "execution_id": "stress",
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

    main.captar_emails(limit=400, execution_id="stress", started_at=now, manifest=manifest)
