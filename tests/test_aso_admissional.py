from __future__ import annotations

import importlib
from pathlib import Path


def _load_module(tmp_path, monkeypatch):
    monkeypatch.setenv("ASO_DEST_BASE", str(tmp_path))
    import aso_admissional_email  # noqa: E402
    return importlib.reload(aso_admissional_email)


def test_sanitize_filename(tmp_path, monkeypatch):
    mod = _load_module(tmp_path, monkeypatch)
    name = 'A<>:"/\\|?*  B'
    out = mod.sanitize_filename(name)
    assert "<" not in out and ">" not in out
    assert "  " not in out
    assert out.startswith("A") and out.endswith("B")


def test_hash_file(tmp_path, monkeypatch):
    mod = _load_module(tmp_path, monkeypatch)
    p = Path(tmp_path) / "f.bin"
    p.write_bytes(b"abc")
    h = mod.hash_file(str(p))
    assert h is not None and len(h) == 32
