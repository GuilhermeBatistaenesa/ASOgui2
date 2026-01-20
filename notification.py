import win32com.client as win32
from datetime import datetime
import os

def enviar_resumo_email(destinatario, relatorio):
    """
    Envia um email de resumo com o relatório HTML.
    destinatario: str (email)
    relatorio: dict (contendo chaves: total, sucessos, erros, tempo_total)
    """
    if not destinatario:
        print("⚠ Destinatário de email não configurado. Pulo envio.")
        return

    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = f"Resumo Processamento ASO RPA - {datetime.now().strftime('%d/%m/%Y')}"
        
        # Formatar listas para HTML
        erros_html = ""
        if relatorio.get('erros'):
            for erro in relatorio['erros']:
                erros_html += f"<tr><td>{erro.get('arquivo')}</td><td style='color:red'>{erro.get('erro')}</td></tr>"
        else:
            erros_html = "<tr><td colspan='2'>Nenhum erro registrado.</td></tr>"

        sucessos_html = ""
        if relatorio.get('sucessos'):
            for suc in relatorio['sucessos'][:50]: # Limitar visualização
                sucessos_html += f"<li>{suc}</li>"
            if len(relatorio['sucessos']) > 50:
                sucessos_html += f"<li>... e mais {len(relatorio['sucessos']) - 50} arquivos.</li>"
        else:
            sucessos_html = "<li>Nenhum arquivo processado com sucesso.</li>"

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
            <h2>Resumo da Execução RPA ASO</h2>
            
            <div class="summary">
                <p><strong>Data:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                <p><strong>Tempo Total:</strong> {relatorio.get('tempo_total', 'N/A')}</p>
                <p><strong>Total Detectado:</strong> {relatorio.get('total', 0)}</p>
                <p><strong>Sucessos:</strong> <span style="color:green; font-weight:bold">{len(relatorio.get('sucessos', []))}</span></p>
                <p><strong>Erros:</strong> <span style="color:red; font-weight:bold">{len(relatorio.get('erros', []))}</span></p>
            </div>

            <h3>Erros ({len(relatorio.get('erros', []))})</h3>
            <table>
                <tr><th>Arquivo</th><th>Erro</th></tr>
                {erros_html}
            </table>
            
            <h3>Sucessos Recentes</h3>
            <ul>
                {sucessos_html}
            </ul>
        </body>
        </html>
        """
        
        mail.HTMLBody = html_body
        mail.Send()
        print(f"📧 Email enviado para {destinatario}")
        return True
    except Exception as e:
        print(f"❌ Erro ao enviar email: {e}")
        return False
