from __future__ import annotations

import json
from pathlib import Path

import runner


def test_parse_semver_and_compare():
    assert runner.parse_semver("1.2.3") == (1, 2, 3)
    assert runner.parse_semver("v2.0.1") == (2, 0, 1)
    assert runner.compare_semver("1.2.3", "1.2.4") == -1
    assert runner.compare_semver("1.2.3", "1.2.3") == 0
    assert runner.compare_semver("2.0.0", "1.9.9") == 1


def test_load_config_defaults(tmp_path):
    cfg_path = Path(tmp_path) / "config.json"
    data = {
        "network_release_dir": str(tmp_path / "releases"),
        "install_dir": str(tmp_path / "install"),
        "prefer_network": False,
    }
    cfg_path.write_text(json.dumps(data), encoding="utf-8")

    cfg = runner.load_config(str(cfg_path))
    assert cfg["network_release_dir"] == data["network_release_dir"]
    assert cfg["install_dir"] == data["install_dir"]
    assert cfg["prefer_network"] is False
    assert cfg["network_latest_json"].endswith("latest.json")


def test_version_read_write(tmp_path):
    app_dir = Path(tmp_path) / "app"
    app_dir.mkdir()
    runner.write_current_version(str(app_dir), "9.9.9")
    assert runner.read_current_version(str(app_dir)) == "9.9.9"


def test_acquire_and_release_lock(tmp_path):
    lock_path = Path(tmp_path) / "lock.txt"
    ok, msg = runner.acquire_lock(str(lock_path), max_age_minutes=30)
    assert ok is True
    runner.release_lock(str(lock_path))
    ok2, msg2 = runner.acquire_lock(str(lock_path), max_age_minutes=30)
    assert ok2 is True
