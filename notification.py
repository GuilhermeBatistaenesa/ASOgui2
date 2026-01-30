import win32com.client as win32
from datetime import datetime
import os
from outcomes import SKIPPED_NO_RECIPIENT
from utils_masking import mask_cpf_in_text


def _get_env_recipients():
    raw = os.getenv("ASO_NOTIFY_TO") or os.getenv("ASO_EMAIL_TO") or ""
    parts = [p.strip() for p in raw.replace(",", ";").split(";") if p.strip()]
    return ";".join(parts)


def _get_env_sender():
    return os.getenv("ASO_EMAIL_FROM") or os.getenv("ASO_EMAIL_ACCOUNT") or ""


def enviar_resumo_email(destinatario, relatorio, execution_id, run_status, report_paths=None, manifest_path=None, logger=None):
    """
    Envia email de resumo com anexos.
    Retorna (status, error_message)
    """
    env_to = _get_env_recipients()
    if env_to:
        destinatario = env_to

    if not destinatario:
        msg = "Destinatario de email nao configurado. Pulo envio."
        if logger:
            logger.warning(msg, step="email")
        else:
            print(f"[WARN] {msg}")
        return (SKIPPED_NO_RECIPIENT, msg)

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        sender = _get_env_sender()
        if sender:
            try:
                mail.SentOnBehalfOfName = sender
            except Exception:
                pass
        date_str = datetime.now().strftime('%Y-%m-%d')
        mail.Subject = f"[ASO] {run_status} | {date_str} | exec={execution_id}"

        erros_html = ""
        if relatorio.get('erros'):
            for erro in relatorio['erros']:
                arquivo = mask_cpf_in_text(str(erro.get('arquivo')))
                msg = mask_cpf_in_text(str(erro.get('erro')))
                erros_html += f"<tr><td>{arquivo}</td><td style='color:red'>{msg}</td></tr>"
        else:
            erros_html = "<tr><td colspan='2'>Nenhum erro registrado.</td></tr>"

        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; }}
                h2 {{ color: #2c3e50; }}
                .summary {{ background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th, td {{ padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }}
                th {{ background-color: #f2f2f2; }}
            </style>
        </head>
        <body>
            <h2>Resumo da Execucao RPA ASO</h2>
            <p><strong>Execution ID:</strong> {execution_id}</p>
            <p><strong>Status:</strong> {run_status}</p>

            <div class="summary">
                <p><strong>Data:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                <p><strong>Tempo Total:</strong> {relatorio.get('tempo_total', 'N/A')}</p>
                <p><strong>Total Detectado:</strong> {relatorio.get('total_detected', 0)}</p>
                <p><strong>Total Processado:</strong> {relatorio.get('total_processed', 0)}</p>
                <p><strong>Sucessos:</strong> <span style="color:green; font-weight:bold">{relatorio.get('success', 0)}</span></p>
                <p><strong>Erros:</strong> <span style="color:red; font-weight:bold">{relatorio.get('error', 0)}</span></p>
            </div>

            <h3>Erros ({relatorio.get('error', 0)})</h3>
            <table>
                <tr><th>Arquivo</th><th>Erro</th></tr>
                {erros_html}
            </table>

            <h3>Evidencias</h3>
            <ul>
                <li>Report JSON: {report_paths.get('json') if report_paths else 'N/A'}</li>
                <li>Resumo MD: {report_paths.get('md') if report_paths else 'N/A'}</li>
                <li>Manifest: {manifest_path or 'N/A'}</li>
            </ul>
        </body>
        </html>
        """

        mail.HTMLBody = html_body

        for path in (report_paths or {}).values():
            if path and os.path.exists(path):
                mail.Attachments.Add(Source=path)

        if manifest_path and os.path.exists(manifest_path):
            mail.Attachments.Add(Source=manifest_path)

        mail.Send()
        if logger:
            logger.info(f"Email enviado para {destinatario}", step="email")
        else:
            print(f"Email enviado para {destinatario}")
        return ("SENT", None)
    except Exception as e:
        if logger:
            logger.error(f"Erro ao enviar email: {e}", step="email")
        else:
            print(f"[ERROR] Erro ao enviar email: {e}")
        return ("FAILED", str(e))
