import os
import re
import hashlib
import traceback
from datetime import datetime, timedelta

import win32com.client as win32
from dotenv import load_dotenv

load_dotenv()

# ----------------------------------------------------
# CONFIGURAÇÕES (pode sobrescrever via variáveis de ambiente)
# ----------------------------------------------------
TARGET_ACCOUNT = os.getenv("ASO_EMAIL_ACCOUNT", "aso@enesa.com.br")
SUBJECT_PREFIX = os.getenv("ASO_SUBJECT_PREFIX", "ASO ADMISSIONAL")
MAILBOX_NAME = os.getenv("ASO_MAILBOX_NAME")  # ex.: "Aso" (nome exibido na lista de pastas)
STORE_NAME = os.getenv("ASO_STORE_NAME")  # ex.: "Aso" (nome da Store no Outlook)
DEST_BASE = os.getenv("ASO_DEST_BASE", r"P:\ASO_ADMISSIONAL")

# extensões de anexos permitidas (separe por vírgula em ASO_ATTACH_EXTS)
ATTACH_EXTS = tuple(
    ext.strip().lower()
    for ext in os.getenv("ASO_ATTACH_EXTS", ".pdf").split(",")
    if ext.strip()
)

DAYS_BACK = int(os.getenv("ASO_DAYS_BACK", "3"))
MAX_EMAILS = int(os.getenv("ASO_MAX_EMAILS", "400"))

# Profundidade máxima de varredura de pastas (fallback recursivo)
MAPI_SCAN_DEPTH = int(os.getenv("ASO_MAPI_SCAN_DEPTH", "6"))

LOG_DIR = os.path.join(DEST_BASE, "logs")
os.makedirs(LOG_DIR, exist_ok=True)


def registrar_log(msg: str) -> None:
    """Grava log diário em DEST_BASE/logs e também imprime na saída."""
    data = datetime.now().strftime("%Y-%m-%d")
    caminho_log = os.path.join(LOG_DIR, f"log_{data}.txt")
    try:
        with open(caminho_log, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{datetime.now()}] {msg}\n")
    except Exception:
        pass
    print(msg)


def enviar_resumo_email(destinatario, relatorio):
    """
    Envia um email de resumo com o relatório HTML.
    destinatario: str (email)
    relatorio: dict (contendo chaves: total, sucessos, erros, tempo_total)
    """
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = f"Resumo Processamento RPA Yube - {datetime.now().strftime('%d/%m/%Y')}"
        
        # Constrói corpo HTML
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
                .success {{ color: green; font-weight: bold; }}
                .error {{ color: red; font-weight: bold; }}
            </style>
        </head>
        <body>
            <h2>Resumo da Execução RPA</h2>
            
            <div class="summary">
                <p><strong>Data:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                <p><strong>Tempo Total:</strong> {relatorio.get('tempo_total', 'N/A')}</p>
                <p><strong>Total Processado:</strong> {relatorio.get('total', 0)}</p>
                <p><strong>Sucessos:</strong> <span class="success">{len(relatorio.get('sucessos', []))}</span></p>
                <p><strong>Erros:</strong> <span class="error">{len(relatorio.get('erros', []))}</span></p>
            </div>

            <h3>Detalhes dos Erros ({len(relatorio.get('erros', []))})</h3>
            <table>
                <tr>
                    <th>Nome/Arquivo</th>
                    <th>Motivo</th>
                </tr>
        """
        
        erros = relatorio.get('erros', [])
        if not erros:
            html_body += "<tr><td colspan='2'>Nenhum erro registrado.</td></tr>"
        else:
            for erro in erros:
                html_body += f"""
                    <tr>
                        <td>{erro.get('arquivo', 'Desconhecido')}</td>
                        <td class="error">{erro.get('erro', 'N/A')}</td>
                    </tr>
                """
            
        html_body += """
            </table>
            
            <h3>Arquivos Processados com Sucesso</h3>
            <ul>
        """
        
        sucessos = relatorio.get('sucessos', [])
        if not sucessos:
             html_body += "<li>Nenhum sucesso registrado.</li>"
        else:
            for sucesso in sucessos:
                html_body += f"<li>{sucesso}</li>"
            
        html_body += """
            </ul>
            <p>Este é um email automático gerado pelo robô Esthergu.</p>
        </body>
        </html>
        """
        
        mail.HTMLBody = html_body
        mail.Send()
        print(f"📧 Email de resumo enviado para: {destinatario}")
        registrar_log(f"Email de resumo enviado para: {destinatario}")

    except Exception as e:
        registrar_log(f"Falha ao enviar email de resumo: {e}")


def sanitize_filename(name: str) -> str:
    """Remove caracteres inválidos para NTFS e normaliza espaços."""
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name or "anexo"


def hash_file(path: str) -> str | None:
    """Retorna hash MD5 do arquivo ou None em caso de erro."""
    md5 = hashlib.md5()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                md5.update(chunk)
        return md5.hexdigest()
    except Exception:
        return None


def conectar_outlook():
    """Abre namespace MAPI do Outlook."""
    return win32.Dispatch("Outlook.Application").GetNamespace("MAPI")


def obter_conta(namespace):
    """Tenta localizar a conta pela propriedade DisplayName ou SmtpAddress."""
    target = TARGET_ACCOUNT.lower()
    for acc in namespace.Accounts:
        try:
            if getattr(acc, "DisplayName", "").lower() == target:
                return acc
        except Exception:
            pass
        try:
            if hasattr(acc, "SmtpAddress") and getattr(acc, "SmtpAddress", "").lower() == target:
                return acc
        except Exception:
            pass
    return None


def listar_mailboxes(namespace):
    """Retorna nomes de pastas raiz disponíveis (para ajudar debug)."""
    nomes = []
    try:
        roots = namespace.Folders
        for i in range(1, roots.Count + 1):
            try:
                nomes.append(roots.Item(i).Name)
            except Exception:
                pass
    except Exception:
        pass
    return nomes


def listar_stores(namespace):
    """Lista Stores disponíveis (útil para shared mailbox mapeada como store)."""
    nomes = []
    try:
        stores = namespace.Stores
        for i in range(1, stores.Count + 1):
            try:
                store = stores.Item(i)
                nomes.append(getattr(store, "DisplayName", ""))
            except Exception:
                pass
    except Exception:
        pass
    return nomes


def obter_inbox_compartilhada(namespace, address: str):
    """Tenta abrir Inbox de mailbox compartilhada via CreateRecipient."""
    try:
        recip = namespace.CreateRecipient(address)
        recip.Resolve()
        if not recip.Resolved:
            return None
        # 6 = olFolderInbox
        return namespace.GetSharedDefaultFolder(recip, 6)
    except Exception:
        return None


# -----------------------------
# FALLBACK MATADOR: varredura MAPI
# -----------------------------
def _iter_folders(folder, max_depth=4, depth=0):
    """Itera recursivamente por pastas MAPI com limite de profundidade (evita travar)."""
    if depth > max_depth:
        return
    try:
        yield folder
        subs = folder.Folders
        for i in range(1, subs.Count + 1):
            try:
                sub = subs.Item(i)
            except Exception:
                continue
            yield from _iter_folders(sub, max_depth=max_depth, depth=depth + 1)
    except Exception:
        return


def procurar_pasta_por_nome(namespace, wanted_names, max_depth=4):
    """
    Procura uma pasta por nome (case-insensitive) em:
    - namespace.Folders (roots)
    - cada Store.GetRootFolder()
    Retorna o primeiro folder que casar.
    """
    wanted = [w.lower() for w in wanted_names if w]

    # 1) roots (namespace.Folders)
    try:
        roots = namespace.Folders
        for i in range(1, roots.Count + 1):
            try:
                root = roots.Item(i)
            except Exception:
                continue
            for f in _iter_folders(root, max_depth=max_depth):
                try:
                    name = (getattr(f, "Name", "") or "").lower()
                    if any(w == name or w in name for w in wanted):
                        return f
                except Exception:
                    pass
    except Exception:
        pass

    # 2) stores root folders
    try:
        stores = namespace.Stores
        for i in range(1, stores.Count + 1):
            try:
                store = stores.Item(i)
                root = store.GetRootFolder()
            except Exception:
                continue
            for f in _iter_folders(root, max_depth=max_depth):
                try:
                    name = (getattr(f, "Name", "") or "").lower()
                    if any(w == name or w in name for w in wanted):
                        return f
                except Exception:
                    pass
    except Exception:
        pass

    return None


def obter_inbox_de_uma_raiz(mailbox_root):
    """Dado um root folder, tenta achar a Inbox em PT/EN."""
    for nm in ("Caixa de Entrada", "Inbox", "Entrada"):
        try:
            return mailbox_root.Folders(nm)
        except Exception:
            continue
    return None


def salvar_anexos(msg, destino_base, hashes_vistos: set[str]) -> int:
    """Salva anexos permitidos do email e evita duplicados por hash."""
    data_pasta = msg.ReceivedTime.strftime("%Y-%m-%d")
    pasta_destino = os.path.join(destino_base, data_pasta)
    os.makedirs(pasta_destino, exist_ok=True)

    saved = 0
    total = getattr(msg.Attachments, "Count", 0)

    for idx in range(1, total + 1):
        try:
            att = msg.Attachments.Item(idx)
            nome = getattr(att, "FileName", f"anexo_{idx}")
            if ATTACH_EXTS and not nome.lower().endswith(ATTACH_EXTS):
                continue

            nome_seguro = sanitize_filename(nome)
            base, ext = os.path.splitext(nome_seguro)
            destino = os.path.join(pasta_destino, nome_seguro)

            # evita sobrescrever se já existir arquivo com mesmo nome
            seq = 1
            while os.path.exists(destino):
                destino = os.path.join(pasta_destino, f"{base}_{seq}{ext}")
                seq += 1

            att.SaveAsFile(destino)

            # dedup por hash para evitar anexos repetidos em threads
            h = hash_file(destino)
            if h and h in hashes_vistos:
                os.remove(destino)
                continue
            if h:
                hashes_vistos.add(h)

            registrar_log(
                f"Anexo salvo: {destino} | Assunto: {getattr(msg, 'Subject', '')} | "
                f"De: {getattr(msg, 'SenderEmailAddress', '')}"
            )
            saved += 1

        except Exception as e:
            registrar_log(f"Erro ao salvar anexo do email '{getattr(msg, 'Subject', '')}': {e}")

    return saved


def buscar_emails(limit: int = MAX_EMAILS) -> None:
    """
    Lê a caixa de entrada da conta alvo, filtra assunto prefixo 'ASO ADMISSIONAL'
    (case-insensitive) recebido nos últimos DAYS_BACK dias e salva anexos.
    """
    namespace = conectar_outlook()
    conta = obter_conta(namespace)

    mailbox_root = None

    # 1) tentar conta real
    if conta:
        try:
            mailbox_root = namespace.Folders(conta.DisplayName)
            registrar_log(f"Conectado à conta: {conta.DisplayName}")
        except Exception:
            mailbox_root = None

    # 2) fallback: mailbox por nome exibido (se existir como root)
    if not mailbox_root and MAILBOX_NAME:
        try:
            mailbox_root = namespace.Folders(MAILBOX_NAME)
            registrar_log(f"Usando mailbox pelo nome: {MAILBOX_NAME}")
        except Exception:
            mailbox_root = None

    # 3) fallback: inbox compartilhada via CreateRecipient
    if not mailbox_root:
        shared_inbox = obter_inbox_compartilhada(namespace, TARGET_ACCOUNT)
        if shared_inbox:
            mailbox_root = shared_inbox.Parent
            registrar_log(f"Usando inbox compartilhada de: {TARGET_ACCOUNT}")

    # 4) fallback: Store mapeada (Data Files)
    inbox = None
    if not mailbox_root:
        try:
            stores = namespace.Stores
            target_names = [n.lower() for n in (STORE_NAME, MAILBOX_NAME, TARGET_ACCOUNT) if n]
            for i in range(1, stores.Count + 1):
                store = stores.Item(i)
                display = getattr(store, "DisplayName", "") or ""
                dlow = display.lower()
                if target_names and not any(t in dlow for t in target_names):
                    continue
                try:
                    inbox = store.GetDefaultFolder(6)  # olFolderInbox
                    mailbox_root = inbox.Parent
                    registrar_log(f"Usando store: {display}")
                    break
                except Exception:
                    continue
        except Exception:
            pass

    # 5) FALLBACK EXTRA: varrer tudo no MAPI procurando pasta "Aso/ASO"
    if not mailbox_root:
        wanted_names = [STORE_NAME, MAILBOX_NAME, "Aso", "ASO", TARGET_ACCOUNT]
        found = procurar_pasta_por_nome(namespace, wanted_names, max_depth=MAPI_SCAN_DEPTH)
        if found:
            mailbox_root = found
            registrar_log(f"Encontrada pasta por varredura MAPI: {getattr(found, 'Name', '')}")

    # se ainda não achou nada: debug completo
    if not mailbox_root:
        registrar_log("Mailbox não encontrada. Verifique ASO_EMAIL_ACCOUNT/ASO_MAILBOX_NAME/ASO_STORE_NAME.")
        registrar_log(f"Mailboxes disponíveis (namespace.Folders): {listar_mailboxes(namespace)}")
        registrar_log(f"Stores disponíveis (namespace.Stores): {listar_stores(namespace)}")
        try:
            accs = [
                f"{getattr(a, 'DisplayName', '')} | {getattr(a, 'SmtpAddress', '')}"
                for a in namespace.Accounts
            ]
            registrar_log(f"Accounts visíveis: {accs}")
        except Exception:
            pass
        return

    # agora pega Inbox
    try:
        if not inbox:
            inbox = obter_inbox_de_uma_raiz(mailbox_root) or mailbox_root

        itens = inbox.Items
        itens.Sort("ReceivedTime", True)
    except Exception as e:
        registrar_log(f"Erro ao acessar caixa de entrada: {e}")
        return

    limite_data = datetime.now() - timedelta(days=DAYS_BACK)
    hashes_vistos: set[str] = set()
    processados = 0
    salvos = 0

    for i in range(1, min(limit, itens.Count) + 1):
        try:
            msg = itens.Item(i)
            if getattr(msg, "Class", None) != 43:  # não é email
                continue

            if msg.ReceivedTime < limite_data:
                break  # ordenado desc

            assunto = (getattr(msg, "Subject", "") or "").strip()
            if not assunto.upper().startswith(SUBJECT_PREFIX.upper()):
                continue

            salvos += salvar_anexos(msg, DEST_BASE, hashes_vistos)
            processados += 1

        except Exception as e:
            registrar_log(f"Erro ao processar email índice {i}: {e}")
            try:
                registrar_log(traceback.format_exc())
            except Exception:
                pass
            continue

    registrar_log(f"Emails compatíveis processados: {processados} | Anexos salvos: {salvos}")


if __name__ == "__main__":
    registrar_log("===== INÍCIO: COLETA ASO ADMISSIONAL =====")
    try:
        buscar_emails()
    except Exception as e:
        registrar_log(f"Erro fatal: {e}")
        try:
            registrar_log(traceback.format_exc())
        except Exception:
            pass
    registrar_log("===== FIM =====")
