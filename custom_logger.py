import os
import json
import logging
from datetime import datetime

class RpaLogger:
    def __init__(self, log_dir):
        self.log_dir = log_dir
        os.makedirs(log_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y-%m-%d")
        self.log_file = os.path.join(log_dir, f"execution_log_{timestamp}.jsonl")
        
        # Configura log de console também
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def _write_json(self, level, message, extra=None):
        entry = {
            "timestamp": datetime.now().isoformat(),
            "level": level,
            "message": message,
            "extra": extra or {}
        }
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as e:
            print(f"FATAL LOGGING ERROR: {e}")

    def info(self, message, extra=None):
        self._write_json("INFO", message, extra)
        print(f"INFO: {message}")

    def warning(self, message, extra=None):
        self._write_json("WARNING", message, extra)
        print(f"WARN: {message}")

    def error(self, message, extra=None):
        self._write_json("ERROR", message, extra)
        print(f"❌ ERROR: {message}")

    def debug(self, message, extra=None):
        # Debug pode ser ruidoso, talvez não queira printar tudo
        self._write_json("DEBUG", message, extra)
        # print(f"DEBUG: {message}")
