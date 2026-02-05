import json
import os
from datetime import datetime
from utils_masking import mask_pii_in_obj, mask_cpf_in_text


class ReportGenerator:
    def __init__(self, report_dir):
        self.report_dir = report_dir
        os.makedirs(report_dir, exist_ok=True)

    def save_report(self, stats):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"relatorio_{timestamp}.json"
        filepath = os.path.join(self.report_dir, filename)

        safe_stats = mask_pii_in_obj(stats)

        try:
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(safe_stats, f, indent=4, ensure_ascii=False)
            print(f"ðŸ“„ RelatÃ³rio JSON salvo em: {filepath}")

            md_path = self.generate_markdown_summary(safe_stats, timestamp)

            return {"json": filepath, "md": md_path}
        except Exception as e:
            print(f"Erro ao salvar relatÃ³rio: {e}")
            return {"json": None, "md": None}

    def generate_markdown_summary(self, stats, timestamp):
        filename = f"resumo_execucao_{timestamp}.md"
        filepath = os.path.join(self.report_dir, filename)

        execution_id = stats.get("execution_id", "")
        md_content = f"""# Resumo de ExecuÃ§Ã£o - RPA ASO
**Execution ID**: {execution_id}
**Data**: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
**Tempo Total**: {stats.get('tempo_total', 'N/A')}

## Totais
- **Total Detectado**: {stats.get('total_detected', 0)}
- **Total Processado**: {stats.get('total_processed', 0)}
- **Sucessos**: {stats.get('success', 0)}
- **Erros**: {stats.get('error', 0)}
- **Skipped Duplicate**: {stats.get('skipped_duplicate', 0)}
- **Skipped Draft**: {stats.get('skipped_draft', 0)}
- **Skipped Non-ASO**: {stats.get('skipped_non_aso', 0)}

## Detalhes de Erros
"""
        if stats.get('erros'):
            for erro in stats['erros']:
                arquivo = mask_cpf_in_text(erro.get('arquivo', 'Desconhecido'))
                msg = mask_cpf_in_text(erro.get('erro', 'Sem mensagem'))
                md_content += f"- **{arquivo}**: {msg}\n"
        else:
            md_content += "Nenhum erro registrado.\n"

        md_content += "\n## Falhas OCR (nao salvos)\n"
        ocr_items = stats.get("ocr_failures", [])
        if ocr_items:
            for item in ocr_items:
                arquivo = mask_cpf_in_text(item.get("arquivo", "Desconhecido"))
                cpf = mask_cpf_in_text(item.get("cpf", ""))
                nome = mask_cpf_in_text(item.get("nome", "Desconhecido"))
                md_content += f"- **{arquivo}**: nome={nome}, cpf={cpf}\n"
        else:
            md_content += "Nenhuma falha de OCR registrada.\n"

        md_content += "\n## Itens Skipped\n"
        skipped_items = stats.get("skipped_items", [])
        if skipped_items:
            for item in skipped_items:
                item_line = mask_cpf_in_text(item)
                md_content += f"- {item_line}\n"
        else:
            md_content += "Nenhum item skipped.\n"

        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(md_content)
            print(f"ðŸ“„ Resumo Markdown salvo em: {filepath}")
            return filepath
        except Exception as e:
            print(f"Erro ao salvar resumo markdown: {e}")
            return None
