import os
import json
import logging
from datetime import datetime, timezone
from utils_masking import mask_cpf as _mask_cpf
from utils_masking import mask_cpf_in_text, mask_pii_in_obj


def mask_cpf(value: str, keep_last: int = 3, mask_char: str = "*") -> str:
    return _mask_cpf(value, keep_last=keep_last, mask_char=mask_char)


def emit_terminal(level, message, step=None, extra=None):
    level_labels = {
        "INFO": "INFO ",
        "WARNING": "WARN ",
        "ERROR": "ERRO ",
        "FATAL": "FATAL",
        "SUMMARY": "RESUM",
        "DEBUG": "DEBUG",
    }
    label = level_labels.get(level, str(level)[:5].upper().ljust(5))
    safe_message = mask_cpf_in_text(message or "")
    parts = [f"[{label}] "]
    if step:
        parts.append(f"[{str(step).upper()}] ")
    parts.append(safe_message)
    if extra:
        safe_extra = mask_pii_in_obj(extra)
        if isinstance(safe_extra, dict):
            context = ", ".join(f"{key}={safe_extra[key]}" for key in sorted(safe_extra))
        else:
            context = str(safe_extra)
        if context:
            parts.append(f" | {context}")
    print("".join(parts))


class RpaLogger:
    DIVIDER = "=" * 60
    SUBDIVIDER = "-" * 60
    LEVEL_LABELS = {
        "INFO": "INFO ",
        "WARNING": "WARN ",
        "ERROR": "ERRO ",
        "FATAL": "FATAL",
        "SUMMARY": "RESUM",
        "DEBUG": "DEBUG",
    }

    def __init__(self, log_dir, execution_id=None, robot_name="Automacao", robot_version="N/D", environment="Producao"):
        self.log_dir = log_dir
        self.execution_id = execution_id
        self.robot_name = robot_name
        self.robot_version = robot_version
        self.environment = environment
        self.run_status = None
        self.started_at = datetime.now()
        os.makedirs(log_dir, exist_ok=True)
        self.log_file = ""
        self._refresh_log_file()

        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def set_execution_id(self, execution_id: str | None):
        self.execution_id = execution_id
        self._refresh_log_file()

    def set_run_status(self, run_status: str | None):
        self.run_status = run_status

    def set_robot_context(self, robot_name=None, robot_version=None, environment=None):
        if robot_name:
            self.robot_name = robot_name
        if robot_version:
            self.robot_version = robot_version
        if environment:
            self.environment = environment

    def start_run(self, started_at=None, execution_id=None):
        if execution_id:
            self.execution_id = execution_id
        self.started_at = started_at or datetime.now()
        self.run_status = "EM_EXECUCAO"
        self._refresh_log_file()
        print(self.DIVIDER)
        print("ENESA | AUTOMACAO CORPORATIVA")
        print(f"Robo: {self.robot_name}")
        print(f"Versao: {self.robot_version}")
        print(f"Execucao: {self.execution_id or 'N/D'}")
        print(f"Ambiente: {self.environment}")
        print(f"Inicio: {self._format_dt(self.started_at)}")
        print(self.DIVIDER)

    def stage(self, current, total, title):
        print("")
        print(f"[ETAPA {current}/{total}] {title}")
        print(self.SUBDIVIDER)

    def finish_run(self, status, started_at=None, finished_at=None, totals=None, report_path=None):
        finished_at = finished_at or datetime.now()
        started_at = started_at or finished_at
        self.run_status = status
        duration = finished_at - started_at if isinstance(started_at, datetime) and isinstance(finished_at, datetime) else None
        totals = totals or {}
        print("")
        print(self.DIVIDER)
        print("RESUMO FINAL DA EXECUCAO")
        print(f"Robo: {self.robot_name}")
        print(f"Execucao: {self.execution_id or 'N/D'}")
        print(f"Status: {status}")
        if "received" in totals:
            print(f"Itens recebidos: {totals['received']}")
        if "processed" in totals:
            print(f"Itens processados: {totals['processed']}")
        if "success" in totals:
            print(f"Itens com sucesso: {totals['success']}")
        if "error" in totals:
            print(f"Itens com erro: {totals['error']}")
        if report_path:
            print(f"Relatorio: {report_path}")
        print(f"Fim: {self._format_dt(finished_at)}")
        if duration is not None:
            print(f"Duracao: {self._format_duration(duration)}")
        print(self.DIVIDER)

    def _refresh_log_file(self):
        timestamp = self._artifact_timestamp(self.started_at)
        exec_id = self.execution_id or "pending"
        filename = f"execution_{timestamp}__{exec_id}.jsonl"
        self.log_file = os.path.join(self.log_dir, filename)

    def _write_json(self, level, message, details=None, step=None, execution_id=None):
        safe_message = mask_cpf_in_text(message or "")
        safe_details = mask_pii_in_obj(details or {})
        exec_id = execution_id or self.execution_id
        severity = self._severity_from_level(level)
        entry = {
            "timestamp_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
            "robot_name": self.robot_name,
            "robot_version": self.robot_version,
            "execution_id": exec_id,
            "run_status": self.run_status,
            "step": step,
            "event_type": "runtime_event",
            "severity": severity,
            "source_file": safe_details.get("source_file", "") if isinstance(safe_details, dict) else "",
            "correlation_keys": self._build_correlation_keys(safe_details),
            "level": level,
            "message": safe_message,
            "details": safe_details,
        }
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as e:
            print(f"FATAL LOGGING ERROR: {e}")

    def _emit_console(self, level, message, extra=None, step=None):
        emit_terminal(level, message, step=step, extra=extra)

    @staticmethod
    def _artifact_timestamp(value):
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d_%H-%M-%S")
        return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    @staticmethod
    def _format_dt(value):
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return str(value)

    @staticmethod
    def _format_duration(delta):
        total_seconds = int(delta.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    @staticmethod
    def _format_context(extra):
        if not extra:
            return ""
        safe_extra = mask_pii_in_obj(extra)
        if isinstance(safe_extra, dict):
            return ", ".join(f"{key}={safe_extra[key]}" for key in sorted(safe_extra))
        return str(safe_extra)

    @staticmethod
    def _severity_from_level(level):
        mapping = {
            "DEBUG": "LOW",
            "INFO": "INFO",
            "WARNING": "MEDIUM",
            "ERROR": "HIGH",
            "FATAL": "CRITICAL",
        }
        return mapping.get(level, "INFO")

    @staticmethod
    def _build_correlation_keys(details):
        if not isinstance(details, dict):
            return []
        keys = []
        for key in ("cpf", "arquivo", "registro_id", "outcome"):
            value = details.get(key)
            if value not in (None, ""):
                keys.append(f"{key}:{value}")
        return keys

    def info(self, message, extra=None, step=None, execution_id=None):
        self._write_json("INFO", message, extra, step=step, execution_id=execution_id)
        self._emit_console("INFO", message, extra=extra, step=step)

    def warning(self, message, extra=None, step=None, execution_id=None):
        self._write_json("WARNING", message, extra, step=step, execution_id=execution_id)
        self._emit_console("WARNING", message, extra=extra, step=step)

    def error(self, message, extra=None, step=None, execution_id=None):
        self._write_json("ERROR", message, extra, step=step, execution_id=execution_id)
        self._emit_console("ERROR", message, extra=extra, step=step)

    def debug(self, message, extra=None, step=None, execution_id=None):
        self._write_json("DEBUG", message, extra, step=step, execution_id=execution_id)
        # Uncomment for verbose console output
        # ctx_str = f" | {extra}" if extra else ""
        # print(f"ðŸ› DEBUG: {message}{ctx_str}")
