from idempotency import should_skip_duplicate


def test_should_skip_duplicate(tmp_path):
    path = tmp_path / "file.txt"
    assert should_skip_duplicate(str(path)) is False
    path.write_text("x", encoding="utf-8")
    assert should_skip_duplicate(str(path)) is True
