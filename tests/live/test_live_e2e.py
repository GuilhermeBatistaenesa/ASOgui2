from __future__ import annotations

import os
from datetime import datetime

import pytest


if os.getenv("RUN_LIVE_TESTS") != "1":
    pytest.skip("live tests skipped (set RUN_LIVE_TESTS=1)", allow_module_level=True)


@pytest.mark.live
def test_live_outlook_pipeline(load_main, tmp_path, monkeypatch):
    base = os.getenv("ASO_LIVE_BASE") or str(tmp_path)
    env = {
        "PROCESSO_ASO_BASE": base,
        "ASO_DAYS_BACK": os.getenv("ASO_LIVE_DAYS_BACK", "0"),
    }

    main = load_main(env=env, stub_win32=False)

    if os.getenv("RUN_LIVE_RPA") != "1":
        monkeypatch.setattr(main, "run_from_main", lambda *_args, **_kwargs: {"sucessos": [], "erros": []})
    if os.getenv("RUN_LIVE_EMAIL") != "1":
        monkeypatch.setattr(main, "enviar_resumo_email", lambda *_args, **_kwargs: ("SKIPPED", None))

    limit = int(os.getenv("ASO_LIVE_LIMIT", "20"))
    started = datetime.now()

    manifest = {
        "execution_id": "live",
        "started_at": started.isoformat(),
        "finished_at": None,
        "duration_sec": None,
        "run_status": None,
        "paths": {},
        "email_status": None,
        "email_error": None,
        "totals": {},
        "items": [],
    }

    main.captar_emails(limit=limit, execution_id="live", started_at=started, manifest=manifest)
    assert manifest["run_status"] in {"CONSISTENT", "INCONSISTENT"}


@pytest.mark.live
def test_live_full_rpa(load_main, tmp_path):
    if os.getenv("RUN_LIVE_RPA") != "1":
        pytest.skip("full RPA live test skipped (set RUN_LIVE_RPA=1)")

    base = os.getenv("ASO_LIVE_BASE") or str(tmp_path)
    env = {
        "PROCESSO_ASO_BASE": base,
        "ASO_DAYS_BACK": os.getenv("ASO_LIVE_DAYS_BACK", "0"),
    }

    main = load_main(env=env, stub_win32=False)
    limit = int(os.getenv("ASO_LIVE_LIMIT", "5"))
    started = datetime.now()

    manifest = {
        "execution_id": "live-rpa",
        "started_at": started.isoformat(),
        "finished_at": None,
        "duration_sec": None,
        "run_status": None,
        "paths": {},
        "email_status": None,
        "email_error": None,
        "totals": {},
        "items": [],
    }

    main.captar_emails(limit=limit, execution_id="live-rpa", started_at=started, manifest=manifest)
    assert manifest["run_status"] in {"CONSISTENT", "INCONSISTENT"}
