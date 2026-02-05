import json
from pathlib import Path

from custom_logger import RpaLogger


def test_rpa_logger_masks_cpf(tmp_path):
    log_dir = Path(tmp_path) / "logs"
    logger = RpaLogger(str(log_dir), execution_id="exec-1")

    logger.info("CPF 12345678901", extra={"cpf": "12345678901"}, step="test")

    log_file = Path(logger.log_file)
    assert log_file.exists()

    line = log_file.read_text(encoding="utf-8").strip().splitlines()[-1]
    assert line.startswith("{") and line.endswith("}")
    obj = json.loads(line)
    assert obj["execution_id"] == "exec-1"
    assert "12345678901" not in line
