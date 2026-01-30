from pathlib import Path

from runner import _is_onedir_release, atomic_replace, verify_sha256


def test_verify_sha256_matches(tmp_path):
    file_path = tmp_path / "data.bin"
    file_path.write_bytes(b"abc")
    assert verify_sha256(str(file_path), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")


def test_verify_sha256_skips_when_expected_empty(tmp_path):
    file_path = tmp_path / "data.bin"
    file_path.write_bytes(b"abc")
    assert verify_sha256(str(file_path), "") is True
    assert verify_sha256(str(file_path), None) is True


def test_atomic_replace_swaps_files(tmp_path):
    src = tmp_path / "src.txt"
    dst = tmp_path / "dst.txt"
    src.write_text("new", encoding="utf-8")
    dst.write_text("old", encoding="utf-8")

    atomic_replace(str(src), str(dst))
    assert dst.read_text(encoding="utf-8") == "new"


def test_is_onedir_release_detects_layout(tmp_path):
    release_dir = tmp_path / "ASOgui"
    release_dir.mkdir()
    (release_dir / "ASOgui.exe").write_bytes(b"")
    (release_dir / "_internal").mkdir()

    assert _is_onedir_release(str(release_dir)) is True


def test_is_onedir_release_false_when_missing_exe(tmp_path):
    release_dir = tmp_path / "ASOgui"
    release_dir.mkdir()
    assert _is_onedir_release(str(release_dir)) is False
