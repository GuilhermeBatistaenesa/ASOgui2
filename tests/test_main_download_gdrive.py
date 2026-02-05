from __future__ import annotations


class _FakeResponse:
    def __init__(self, data: bytes, headers: dict):
        self._data = data
        self._pos = 0
        self.headers = headers

    def read(self, n: int = -1):
        if self._pos >= len(self._data):
            return b""
        if n == -1:
            n = len(self._data) - self._pos
        chunk = self._data[self._pos : self._pos + n]
        self._pos += n
        return chunk

    def close(self):
        return None


class _FakeOpener:
    def __init__(self, responses):
        self._responses = list(responses)
        self.calls = 0

    def open(self, req, timeout=0):
        self.calls += 1
        return self._responses.pop(0)


def test_download_gdrive_file_direct(load_main, tmp_path, monkeypatch):
    main = load_main(env={"ASO_GDRIVE_NAME_FILTER": "asos enesa"})

    data = b"%PDF-1.4 fake pdf bytes"
    headers = {
        "Content-Disposition": 'attachment; filename="ASOS ENESA 1.pdf"',
        "Content-Type": "application/pdf",
    }
    opener = _FakeOpener([_FakeResponse(data, headers)])

    monkeypatch.setattr(main.urllib.request, "build_opener", lambda *_: opener)

    out = main.download_gdrive_file("fileid", str(tmp_path))
    assert out is not None
    assert out.lower().endswith(".pdf")


def test_download_gdrive_file_confirm_flow(load_main, tmp_path, monkeypatch):
    main = load_main(env={"ASO_GDRIVE_NAME_FILTER": "asos enesa"})

    html = b"<html><title>ASOS ENESA 2.pdf - Google Drive</title>confirm=abc123</html>"
    first_headers = {"Content-Type": "text/html"}
    second_headers = {
        "Content-Disposition": "attachment; filename=ASOS ENESA 2.pdf",
        "Content-Type": "application/pdf",
    }
    opener = _FakeOpener([
        _FakeResponse(html, first_headers),
        _FakeResponse(b"%PDF-1.4 data", second_headers),
    ])

    monkeypatch.setattr(main.urllib.request, "build_opener", lambda *_: opener)

    out = main.download_gdrive_file("fileid2", str(tmp_path))
    assert out is not None
    assert out.lower().endswith(".pdf")
