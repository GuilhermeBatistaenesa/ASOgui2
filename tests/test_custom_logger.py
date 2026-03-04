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
    assert obj["robot_name"] == "Automacao"
    assert obj["robot_version"] == "N/D"
    assert obj["run_status"] is None
    assert obj["event_type"] == "runtime_event"
    assert obj["severity"] == "INFO"
    assert "12345678901" not in line
    assert "__exec-1" in log_file.name


def test_rpa_logger_console_format(capsys, tmp_path):
    log_dir = Path(tmp_path) / "logs"
    logger = RpaLogger(str(log_dir), execution_id="exec-2", robot_name="ProcessoASO", robot_version="1.2.3")

    logger.info("Validando configuracoes", extra={"count": 3}, step="setup")

    captured = capsys.readouterr()
    assert "[INFO ] [SETUP] Validando configuracoes | count=3" in captured.out
