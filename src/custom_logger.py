import os
import json
import logging
from datetime import datetime
from utils_masking import mask_cpf as _mask_cpf
from utils_masking import mask_cpf_in_text, mask_pii_in_obj


def mask_cpf(value: str, keep_last: int = 3, mask_char: str = "*") -> str:
    return _mask_cpf(value, keep_last=keep_last, mask_char=mask_char)


class RpaLogger:
    def __init__(self, log_dir, execution_id=None):
        self.log_dir = log_dir
        self.execution_id = execution_id
        os.makedirs(log_dir, exist_ok=True)

        timestamp = datetime.now().strftime("%Y-%m-%d")
        self.log_file = os.path.join(log_dir, f"execution_log_{timestamp}.jsonl")

        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def set_execution_id(self, execution_id: str | None):
        self.execution_id = execution_id

    def _write_json(self, level, message, details=None, step=None, execution_id=None):
        safe_message = mask_cpf_in_text(message or "")
        safe_details = mask_pii_in_obj(details or {})
        exec_id = execution_id or self.execution_id
        entry = {
            "timestamp": datetime.now().isoformat(),
            "execution_id": exec_id,
            "step": step,
            "level": level,
            "message": safe_message,
            "details": safe_details,
        }
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as e:
            print(f"FATAL LOGGING ERROR: {e}")

    def info(self, message, extra=None, step=None, execution_id=None):
        self._write_json("INFO", message, extra, step=step, execution_id=execution_id)
        ctx_str = f" | {mask_pii_in_obj(extra)}" if extra else ""
        print(f"INFO: {mask_cpf_in_text(message)}{ctx_str}")

    def warning(self, message, extra=None, step=None, execution_id=None):
        self._write_json("WARNING", message, extra, step=step, execution_id=execution_id)
        ctx_str = f" | {mask_pii_in_obj(extra)}" if extra else ""
        print(f"WARN: {mask_cpf_in_text(message)}{ctx_str}")

    def error(self, message, extra=None, step=None, execution_id=None):
        self._write_json("ERROR", message, extra, step=step, execution_id=execution_id)
        ctx_str = f" | {mask_pii_in_obj(extra)}" if extra else ""
        print(f"ERROR: {mask_cpf_in_text(message)}{ctx_str}")

    def debug(self, message, extra=None, step=None, execution_id=None):
        self._write_json("DEBUG", message, extra, step=step, execution_id=execution_id)
        # Uncomment for verbose console output
        # ctx_str = f" | {extra}" if extra else ""
        # print(f"ðŸ› DEBUG: {message}{ctx_str}")
