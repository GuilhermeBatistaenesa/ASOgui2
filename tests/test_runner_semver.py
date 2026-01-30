import os

from runner import compare_semver, parse_semver, read_current_version, write_current_version


def test_parse_semver_accepts_v_prefix():
    assert parse_semver("v1.2.3") == (1, 2, 3)


def test_parse_semver_invalid_returns_zero_tuple():
    assert parse_semver("not-a-version") == (0, 0, 0)
    assert parse_semver("") == (0, 0, 0)
    assert parse_semver(None) == (0, 0, 0)


def test_compare_semver_orders_versions():
    assert compare_semver("1.0.0", "1.0.1") == -1
    assert compare_semver("2.0.0", "1.9.9") == 1
    assert compare_semver("1.2.3", "1.2.3") == 0


def test_read_write_current_version(tmp_path):
    app_dir = tmp_path / "app"
    app_dir.mkdir()

    assert read_current_version(str(app_dir)) == "0.0.0"

    write_current_version(str(app_dir), "9.9.9")
    assert read_current_version(str(app_dir)) == "9.9.9"

    # Uppercase file should be read if lowercase is missing
    os.remove(app_dir / "version.txt")
    (app_dir / "VERSION.txt").write_text("7.7.7", encoding="utf-8")
    assert read_current_version(str(app_dir)) == "7.7.7"
