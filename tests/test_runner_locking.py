import os
import time

from runner import acquire_lock, release_lock


def test_acquire_lock_and_release(tmp_path):
    lock_path = tmp_path / "lockfile.lock"

    ok, msg = acquire_lock(str(lock_path))
    assert ok is True
    assert "locked" in msg

    ok2, msg2 = acquire_lock(str(lock_path))
    assert ok2 is False
    assert "already running" in msg2

    release_lock(str(lock_path))
    ok3, msg3 = acquire_lock(str(lock_path))
    assert ok3 is True
    assert "locked" in msg3


def test_acquire_lock_stale_file(tmp_path):
    lock_path = tmp_path / "lockfile.lock"
    lock_path.write_text("old", encoding="utf-8")
    stale_time = time.time() - (60 * 60)  # 1 hour ago
    os.utime(lock_path, (stale_time, stale_time))

    ok, msg = acquire_lock(str(lock_path), max_age_minutes=1)
    assert ok is True
    assert "locked" in msg
