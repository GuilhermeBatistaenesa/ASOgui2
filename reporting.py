import json
import os
from datetime import datetime

class ReportGenerator:
    def __init__(self, report_dir):
        self.report_dir = report_dir
        os.makedirs(report_dir, exist_ok=True)

    def save_report(self, stats):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"relatorio_{timestamp}.json"
        filepath = os.path.join(self.report_dir, filename)
        
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(stats, f, indent=4, ensure_ascii=False)
            print(f"📄 Relatório salvo em: {filepath}")
            return filepath
        except Exception as e:
            print(f"Erro ao salvar relatório: {e}")
            return None
