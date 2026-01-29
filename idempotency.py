import os


def should_skip_duplicate(dest_path: str) -> bool:
    return os.path.exists(dest_path)
