from __future__ import annotations

import sys
import types
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


@pytest.fixture
def load_main(tmp_path, monkeypatch):
    """
    Importa o modulo main com ambiente controlado (sem Outlook real)
    e pastas apontando para tmp_path.
    """

    def _load(env: dict | None = None, stub_win32: bool = True):
        if stub_win32:
            win32_client = types.SimpleNamespace(GetActiveObject=lambda _: None, DispatchEx=lambda _: None)
            win32com = types.ModuleType("win32com")
            win32com.client = win32_client
            monkeypatch.setitem(sys.modules, "win32com", win32com)
            monkeypatch.setitem(sys.modules, "win32com.client", win32_client)

            pywintypes = types.SimpleNamespace(com_error=Exception, IID=lambda x: x)
            monkeypatch.setitem(sys.modules, "pywintypes", pywintypes)

            def _co_register_message_filter(_):
                return None

            pythoncom = types.SimpleNamespace(
                IID_IOleMessageFilter=object(),
                CoRegisterMessageFilter=_co_register_message_filter,
                CoInitialize=lambda: None,
                CoUninitialize=lambda: None,
                SERVERCALL_RETRYLATER=2,
            )
            monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)

        monkeypatch.setenv("PROCESSO_ASO_BASE", str(tmp_path))
        if env:
            for k, v in env.items():
                if v is None:
                    monkeypatch.delenv(k, raising=False)
                else:
                    monkeypatch.setenv(k, str(v))

        sys.modules.pop("main", None)
        import importlib

        import main  # noqa: E402

        return importlib.reload(main)

    return _load
