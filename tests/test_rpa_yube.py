from __future__ import annotations

import importlib
import os
from pathlib import Path


def test_extrair_cpf_do_nome_and_registrar_log(tmp_path, monkeypatch):
    monkeypatch.setenv("YUBE_USER", "user")
    monkeypatch.setenv("YUBE_PASS", "pass")
    monkeypatch.setenv("PROCESSO_ASO_BASE", str(tmp_path))

    if "rpa_yube" in globals():
        del globals()["rpa_yube"]

    import rpa_yube  # noqa: E402

    importlib.reload(rpa_yube)

    assert rpa_yube.extrair_cpf_do_nome("JOAO - 12345678901.pdf") == "12345678901"
    assert rpa_yube.extrair_cpf_do_nome("JOAO - 123.456.789-01.pdf") == "12345678901"
    assert rpa_yube.extrair_cpf_do_nome("SEM_CPF.pdf") is None

    log_path = Path(tmp_path) / "rpa_log.csv"
    rpa_yube.LOG_CSV = str(log_path)
    rpa_yube.registrar_log("12345678901", "C:\\temp\\12345678901.pdf", "OK", "teste")

    content = log_path.read_text(encoding="utf-8")
    assert "12345678901" not in content
