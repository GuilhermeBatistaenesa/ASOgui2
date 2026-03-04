import os
import re
import time
import shutil
import logging
import csv
import unicodedata
from pathlib import Path
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from custom_logger import emit_terminal
from utils_masking import mask_cpf, mask_cpf_in_text


import sys
import subprocess

def _load_env():
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), ".env"))
    candidates.append(os.path.join(os.getcwd(), ".env"))
    candidates.append(os.path.join(os.path.dirname(__file__), ".env"))
    for p in candidates:
        if p and os.path.exists(p):
            load_dotenv(p, override=True)
            return
    load_dotenv(override=True)

_load_env()

def _ensure_playwright_chromium_installed():
    try:
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
        return True
    except Exception as e:
        logging.error(f"Falha ao instalar Chromium (Playwright): {e}")
        return False


# ---------- CONFIGURAÇÃO ----------
YUBE_URL = os.getenv("YUBE_URL", "https://yube.com.br/")
YUBE_USER = os.getenv("YUBE_USER")
YUBE_PASS = os.getenv("YUBE_PASS")
NAV_TIMEOUT = int(os.getenv("YUBE_NAV_TIMEOUT", "10000"))
KEEP_BROWSER_OPEN = os.getenv("KEEP_BROWSER_OPEN", "1") == "1"
UPLOAD_WAIT_MS = int(os.getenv("UPLOAD_WAIT_MS", "4000"))
POST_SAVE_WAIT_MS = int(os.getenv("POST_SAVE_WAIT_MS", "3000"))
PRE_SAVE_WAIT_MS = int(os.getenv("PRE_SAVE_WAIT_MS", "5000"))
SEARCH_WAIT_MS = int(os.getenv("YUBE_SEARCH_WAIT_MS", "5000"))
RETRY_NOT_FOUND = int(os.getenv("ASO_RETRY_NOT_FOUND", "1"))
RETRY_NOT_FOUND_DELAY_SEC = int(os.getenv("ASO_RETRY_NOT_FOUND_DELAY_SEC", "3"))

# ---------- CONFIGURAÇÃO DE PASTAS CENTRALIZADAS ----------
PROCESSO_ASO_BASE = os.getenv("PROCESSO_ASO_BASE", r"P:\ProcessoASO")
PASTA_PROCESSADOS = os.path.join(PROCESSO_ASO_BASE, "processados")
PASTA_EM_PROCESSAMENTO = os.path.join(PROCESSO_ASO_BASE, "em processamento")
PASTA_ERROS = os.path.join(PROCESSO_ASO_BASE, "erros")
PASTA_LOGS_RPA = os.path.join(PROCESSO_ASO_BASE, "logs")

# Garantir pastas antes de configurar logging
Path(PASTA_PROCESSADOS).mkdir(parents=True, exist_ok=True)
Path(PASTA_EM_PROCESSAMENTO).mkdir(parents=True, exist_ok=True)
Path(PASTA_ERROS).mkdir(parents=True, exist_ok=True)
Path(PASTA_LOGS_RPA).mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    handlers=[
        logging.FileHandler(os.path.join(PASTA_LOGS_RPA, "rpa_yube_debug.log"), encoding='utf-8')
    ],
    level=logging.INFO
)

if not YUBE_USER or not YUBE_PASS:
    logging.error("Credenciais YUBE nÃ£o configuradas. Defina YUBE_USER e YUBE_PASS no ambiente.")
    raise RuntimeError("Credenciais YUBE nÃ£o configuradas. Defina YUBE_USER e YUBE_PASS no ambiente.")

LOG_CSV = os.path.join(PASTA_LOGS_RPA, "rpa_log.csv")

def _cpf_masked(cpf: str | None) -> str:
    if not cpf:
        return "CPF_DESCONHECIDO"
    return mask_cpf(cpf, keep_last=3, mask_char="X")

CPF_REGEX = re.compile(r"(\d{11})")  # procura 11 dígitos seguidos no nome do arquivo

# Seletores principais usados no fluxo (baseados nos tooltips capturados)
SEL_LOGIN_LINK = lambda page: page.get_by_role("link", name="LOGIN")
SEL_ACESSAR_MODULO = lambda page: page.get_by_text("Acessar módulo")
SEL_EMAIL = lambda page: page.locator("#username")
SEL_SENHA = lambda page: page.locator("input[name='password']")
SEL_ENTRAR = lambda page: page.locator("input[id='kc-login']")
SEL_CHECKBOX_TODAS = lambda page: page.get_by_role("checkbox", name=re.compile("Selecionar Todas", re.I))
SEL_BUSCA = lambda page: page.get_by_placeholder(re.compile("Procure por nome, email ou telefone", re.I))
SEL_VER_PROCESSO = lambda page: page.get_by_text("Ver processo")
SEL_EXAME_ADM = lambda page: page.get_by_role("button", name=re.compile(r"(Exame.*Admissional|Saude.*Ocupacional|ASO)", re.I))
SEL_CRIAR_DOC = lambda page: page.get_by_role("button", name="Criar documento")
SEL_UPLOAD_BTN = lambda page: page.get_by_role("button", name=re.compile("Tirar foto do resultado", re.I))
SEL_SALVAR = lambda page: page.get_by_role("button", name="Salvar")
SEL_VOLTAR = lambda page: page.locator("text=Voltar")


def extrair_cpf_do_nome(filename: str) -> str | None:
    # remove tudo que não é dígito e pega os 11 finais, se houver
    digitos = re.sub(r"\D", "", filename or "")
    if len(digitos) >= 11:
        return digitos[-11:]
    m = CPF_REGEX.search(filename)
    if m:
        return m.group(1)
    return None


def _cpf_formatado(cpf: str | None) -> str | None:
    if not cpf:
        return None
    digits = re.sub(r"\D", "", cpf)
    if len(digits) != 11:
        return None
    return f"{digits[0:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:11]}"


def registrar_log(cpf, file_path, status, message=""):
    cpf_safe = _cpf_masked(cpf)
    file_path_safe = mask_cpf_in_text(file_path)
    header = ["timestamp", "cpf", "file", "status", "message"]
    row = [datetime.now().isoformat(), cpf_safe, file_path_safe, status, message]
    file_exists = os.path.isfile(LOG_CSV)
    with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(header)
        writer.writerow(row)


def registrar_screenshot(page, nome_base: str):
    """Salva screenshot em logs/ com nome_base_timestamp."""
    try:
        ts = int(time.time())
        path = os.path.join(PASTA_LOGS_RPA, f"{nome_base}_{ts}.png")
        page.screenshot(path=path, full_page=True)
        logging.info(f"Screenshot: {path}")
        return path
    except Exception as e:
        logging.warning(f"Falha ao salvar screenshot {nome_base}: {e}")
        return None


def login(page):
    # tenta múltiplas URLs/retentativas para evitar travas de rede
    urls = [
        "https://app.yube.com.br/",  # Direto para o APP (mais rápido)
        "https://app.yube.com.br/login",
        YUBE_URL,
        "https://yube.com.br/login",
    ]
    opened = False
    for attempt in range(2):
        for url in urls:
            try:
                page.goto(url, timeout=NAV_TIMEOUT, wait_until="domcontentloaded")
                opened = True
                break
            except Exception as e:
                logging.warning(f"Tentativa {attempt+1} ao acessar {url} falhou: {e}")
                page.wait_for_timeout(2000)
        if opened:
            break
    if not opened:
        emit_terminal("WARNING", "Nao foi possivel abrir a pagina automaticamente. Abra o login da Yube manualmente e pressione Enter.", step="login")
        input()

    target_page = page

    target_page = page
    
    # Verifica se JÁ estamos na tela de login (inputs visíveis)
    try:
        if SEL_EMAIL(page).count() > 0:
             emit_terminal("INFO", "Tela de login ja esta visivel. Prosseguindo.", step="login")
    except:
        pass

    # Se NÃO tem input de senha, força a navegação para login direto (evita nova aba)
    if SEL_EMAIL(page).count() == 0:
        try:
            emit_terminal("INFO", "Redirecionando para a pagina de login direta.", step="login")
            page.goto("https://app.yube.com.br/login", timeout=15000)
            page.wait_for_timeout(2000)
            target_page = page
        except Exception as e:
            logging.warning(f"Erro ao navegar para login: {e}")

    # clicar em acessar módulo (se aparecer)
    try:
        # As vezes o login redireciona para seleção de conta
        if SEL_ACESSAR_MODULO(target_page).count() > 0:
            SEL_ACESSAR_MODULO(target_page).click(timeout=5000)
    except Exception as e:
        pass

    # preencher credenciais
    emit_terminal("INFO", "Preenchendo credenciais.", step="login")
    try:
        # Se ainda assim não aparecer, tenta esperar
        target_page.wait_for_selector("#username", state="visible", timeout=10000)
        SEL_EMAIL(target_page).fill(YUBE_USER)
        target_page.wait_for_selector("input[name='password']", state="visible", timeout=10000)
        SEL_SENHA(target_page).fill(YUBE_PASS)
        SEL_ENTRAR(target_page).click()
        emit_terminal("OK", "Login submetido.", step="login")
    except Exception as e:
        # Se falhar o login, mas já estiver logado (redirecionou para home), tudo bem
        if SEL_BUSCA(target_page).count() > 0:
             emit_terminal("INFO", "Busca visivel. Sessao ja parece autenticada.", step="login")
        else:
             logging.warning(f"Falha ao preencher login: {e}")
             emit_terminal("ERROR", "Erro ao preencher login.", step="login")

    # possível segundo passo (botão idSIB)
    try:
        target_page.wait_for_selector("input[id='idSIB']", timeout=5000)
        target_page.locator("input[id='idSIB']").click()
        target_page.wait_for_selector("input[id='kc-login']", timeout=5000)
        target_page.locator("input[id='kc-login']").click()
    except Exception:
        pass

    # aguarda área autenticada (campo de busca)
    try:
        SEL_BUSCA(target_page).wait_for(timeout=25000)
    except Exception:
        logging.warning("Campo de busca não detectado; confirme manualmente se o login concluiu.")

    return target_page


def filtrar_todas_obras(page):
    """Tenta marcar Selecionar Todas com múltiplos fallbacks e force click."""
    page.wait_for_timeout(1200)
    selectors = [
        'div:has-text("Selecionar Todas") input[type="checkbox"]',
        'label:has-text("Selecionar Todas") input[type="checkbox"]',
        'input[type="checkbox"][aria-label="Selecionar Todas"]',
    ]
    for sel in selectors:
        try:
            box = page.locator(sel).first
            if box.count() > 0:
                if not box.is_checked():
                    box.check(force=True)
                    page.wait_for_timeout(600)
                return "DOC_ALREADY_EXISTS"
        except Exception:
            continue
    # fallback: clicar no texto/div
    try:
        div_sel = page.locator('div:has-text("Selecionar Todas")').first
        if div_sel.count() > 0:
            div_sel.click(force=True)
            page.wait_for_timeout(600)
            return
        page.get_by_text("Selecionar Todas", exact=True).click(force=True)
        page.wait_for_timeout(600)
    except Exception as e:
        logging.debug(f"Filtro 'Selecionar Todas' não aplicado: {e}")


def pesquisar_funcionario(page, cpf: str, nome_hint: str | None = None) -> bool:
    cpf_safe = _cpf_masked(cpf)
    cpf_fmt = _cpf_formatado(cpf)
    busca = SEL_BUSCA(page)
    
    # Tentativa com recuperação automática
    try:
        busca.click(timeout=8000)
    except Exception as e:
        logging.warning(f"Campo busca não encontrado ou clicável: {e}. Tentando recarregar página inicial.")
        try:
             logging.info("Recarregando página inicial...")
             page.goto(YUBE_URL, timeout=20000, wait_until="domcontentloaded")
             page.wait_for_timeout(3000)
             
             # Verifica se caiu no login
             if SEL_LOGIN_LINK(page).count() > 0 or SEL_EMAIL(page).count() > 0:
                 logging.warning("Sessão expirada detectada. Refazendo login...")
                 page = login(page)
             
             # Reaplica filtros pois recarregou a home
             filtrar_todas_obras(page)
             
             busca = SEL_BUSCA(page)
             busca.click(timeout=10000)
        except Exception as e2:
             logging.error(f"Falha fatal ao recuperar página/login: {e2}")
             return False

    try:
        busca.press("Control+A")
        busca.press("Backspace")
    except Exception:
        busca.fill("")
        
    page.wait_for_timeout(300)
    # VOLTA A BUSCAR POR NOME PRIMEIRO (User confirmou que CPF não busca)
    termo_busca = nome_hint if nome_hint else cpf
    
    try:
        busca.fill(termo_busca)
        busca.press("Enter")
    except Exception as e:
         logging.warning(f"Erro ao preencher busca: {e}")
         return False
         
    page.wait_for_timeout(SEARCH_WAIT_MS)
    registrar_screenshot(page, f"busca_input_{cpf_safe}")

    # Tenta clicar em link do candidato (CPF ou nome)
    candidatos = []
    if nome_hint:
        candidatos.append(page.get_by_role("link", name=re.compile(re.escape(nome_hint), re.I)))
        candidatos.append(page.locator("div.card-list a").filter(has_text=re.compile(re.escape(nome_hint), re.I)))
        candidatos.append(page.get_by_text(nome_hint, exact=False))
    candidatos.extend([
        page.get_by_role("link", name=re.compile(re.escape(cpf))),
        page.locator("div.card-list a").filter(has_text=re.compile(re.escape(cpf))),
        page.get_by_text(cpf, exact=False),
    ])
    if cpf_fmt:
        candidatos.extend([
            page.get_by_role("link", name=re.compile(re.escape(cpf_fmt))),
            page.locator("div.card-list a").filter(has_text=re.compile(re.escape(cpf_fmt))),
            page.get_by_text(cpf_fmt, exact=False),
        ])
    if nome_hint:
        try:
            candidatos.append(page.locator("div.card-list").filter(has_text=re.compile(re.escape(nome_hint), re.I)))
        except Exception:
            pass
    for cand in candidatos:
        try:
            if cand.count() > 0:
                texto_card = ""
                try:
                    texto_card = cand.first.inner_text()
                except Exception:
                    pass
                registrar_screenshot(page, f"busca_{cpf_safe}")
                logging.info(f"Selecionando resultado da busca para CPF {cpf_safe}: {texto_card or cpf_safe}")
                try:
                    cand.first.scroll_into_view_if_needed(timeout=3000)
                except Exception:
                    pass
                cand.first.click(force=True)
                page.wait_for_timeout(2000)
                # Verifica se já entrou na tela com "Ver processo"
                try:
                    SEL_VER_PROCESSO(page).first.wait_for(timeout=5000)
                    return True
                except Exception:
                    # não achou ver processo, tenta próximo candidato
                    continue
        except Exception:
            continue

    # --- REMOVIDO FALLBACK PERIGOSO ---
    # Se chegamos aqui, não encontramos o link ESPECÍFICO do funcionário.
    # Antigamente o script clicava em qualquer card, o que causava ERROS GRAVES.
    # Agora retornamos False para que o arquivo vá para a pasta de ERROS.
    
    registrar_screenshot(page, f"busca_falha_{cpf_safe}")
    logging.warning(f"Busca falhou para '{termo_busca}'. Nenhum resultado exato encontrado.")
    return False


def pesquisar_funcionario_robusto(page, cpf: str, nome_hint: str | None = None) -> bool:
    cpf_safe = _cpf_masked(cpf)
    cpf_fmt = _cpf_formatado(cpf)
    busca = SEL_BUSCA(page)

    def _norm(text: str | None) -> str:
        raw = unicodedata.normalize("NFKD", text or "")
        raw = "".join(ch for ch in raw if not unicodedata.combining(ch))
        return re.sub(r"\s+", " ", raw).strip().lower()

    def _wait_open() -> bool:
        checks = [
            lambda: SEL_VER_PROCESSO(page).first.wait_for(timeout=3000),
            lambda: SEL_EXAME_ADM(page).first.wait_for(timeout=3000),
            lambda: SEL_CRIAR_DOC(page).first.wait_for(timeout=3000),
            lambda: SEL_VOLTAR(page).first.wait_for(timeout=3000),
        ]
        for check in checks:
            try:
                check()
                return True
            except Exception:
                continue
        try:
            # Se o campo de busca sumiu, normalmente saiu da lista e entrou no candidato.
            if SEL_BUSCA(page).count() == 0 or not SEL_BUSCA(page).first.is_visible():
                return True
        except Exception:
            return True
        return False

    try:
        busca.click(timeout=8000)
    except Exception as e:
        logging.warning(f"Campo busca nao encontrado ou clicavel: {e}. Tentando recarregar pagina inicial.")
        try:
            page.goto(YUBE_URL, timeout=20000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            if SEL_LOGIN_LINK(page).count() > 0 or SEL_EMAIL(page).count() > 0:
                page = login(page)
            filtrar_todas_obras(page)
            busca = SEL_BUSCA(page)
            busca.click(timeout=10000)
        except Exception as e2:
            logging.error(f"Falha fatal ao recuperar pagina/login: {e2}")
            return False

    try:
        busca.press("Control+A")
        busca.press("Backspace")
    except Exception:
        busca.fill("")

    page.wait_for_timeout(300)
    termo_busca = nome_hint if nome_hint else cpf

    try:
        busca.fill(termo_busca)
        busca.press("Enter")
    except Exception as e:
        logging.warning(f"Erro ao preencher busca: {e}")
        return False

    page.wait_for_timeout(SEARCH_WAIT_MS)
    registrar_screenshot(page, f"busca_input_{cpf_safe}")

    result_scopes = [
        page.locator(".card-list"),
        page.locator(".ant-list"),
    ]

    candidatos = []
    if nome_hint:
        candidatos.append(page.get_by_role("link", name=re.compile(re.escape(nome_hint), re.I)))
        candidatos.append(page.locator("div.card-list a").filter(has_text=re.compile(re.escape(nome_hint), re.I)))
        for scope in result_scopes:
            try:
                candidatos.append(scope.get_by_text(nome_hint, exact=False))
            except Exception:
                continue
    candidatos.extend([
        page.get_by_role("link", name=re.compile(re.escape(cpf))),
        page.locator("div.card-list a").filter(has_text=re.compile(re.escape(cpf))),
    ])
    if cpf_fmt:
        candidatos.extend([
            page.get_by_role("link", name=re.compile(re.escape(cpf_fmt))),
            page.locator("div.card-list a").filter(has_text=re.compile(re.escape(cpf_fmt))),
        ])

    for cand in candidatos:
        try:
            if cand.count() <= 0:
                continue
            texto_card = ""
            try:
                texto_card = cand.first.inner_text()
            except Exception:
                pass
            registrar_screenshot(page, f"busca_{cpf_safe}")
            logging.info(f"Selecionando resultado da busca para CPF {cpf_safe}: {texto_card or cpf_safe}")
            try:
                cand.first.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            cand.first.click(force=True)
            page.wait_for_timeout(2000)
            if _wait_open():
                return True
        except Exception:
            continue

    if nome_hint:
        target_norm = _norm(nome_hint)
        card_scopes = [
            page.locator(".ant-list-item"),
            page.locator("[class*='card']"),
            page.locator("[class*='list-item']"),
        ]
        for scope in card_scopes:
            try:
                total = min(scope.count(), 30)
            except Exception:
                total = 0
            for i in range(total):
                try:
                    card = scope.nth(i)
                    card_text = card.inner_text(timeout=2000)
                    if target_norm not in _norm(card_text):
                        continue
                    registrar_screenshot(page, f"busca_card_{cpf_safe}")
                    logging.info(f"Selecionando card da busca para CPF {cpf_safe}: {card_text[:120]}")
                    try:
                        card.scroll_into_view_if_needed(timeout=3000)
                    except Exception:
                        pass
                    try:
                        card.get_by_text(nome_hint, exact=False).first.click(timeout=3000, force=True)
                    except Exception:
                        card.click(timeout=3000, force=True)
                    page.wait_for_timeout(2500)
                    if _wait_open():
                        return True
                except Exception:
                    continue
            if total > 0:
                break

    # Fallback conservador por CPF: clica apenas em cards cujo texto contenha
    # explicitamente o CPF (bruto ou formatado). Isso evita clique aleatorio.
    cpf_targets = [t for t in (cpf, cpf_fmt) if t]
    if cpf_targets:
        card_scopes = [
            page.locator(".ant-list-item"),
            page.locator("[class*='card']"),
            page.locator("[class*='list-item']"),
        ]
        for scope in card_scopes:
            try:
                total = min(scope.count(), 30)
            except Exception:
                total = 0
            for i in range(total):
                try:
                    card = scope.nth(i)
                    card_text = card.inner_text(timeout=2000)
                    if not any(target in card_text for target in cpf_targets):
                        continue
                    registrar_screenshot(page, f"busca_card_cpf_{cpf_safe}")
                    logging.info(f"Selecionando card da busca por CPF {cpf_safe}: {card_text[:120]}")
                    try:
                        card.scroll_into_view_if_needed(timeout=3000)
                    except Exception:
                        pass
                    try:
                        card.get_by_text(re.compile("|".join(re.escape(t) for t in cpf_targets))).first.click(timeout=3000, force=True)
                    except Exception:
                        card.click(timeout=3000, force=True)
                    page.wait_for_timeout(2500)
                    if _wait_open():
                        return True
                except Exception:
                    continue
            if total > 0:
                break

    registrar_screenshot(page, f"busca_falha_{cpf_safe}")
    logging.warning(f"Busca falhou para '{termo_busca}'. Nenhum resultado acionavel encontrado.")
    return False


def entrar_ver_processo(page):
    try:
        vp = SEL_VER_PROCESSO(page)
        vp.first.wait_for(timeout=12000)
        try:
            vp.first.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass
        vp.first.click(timeout=12000, force=True)
        page.wait_for_timeout(1200)
        return "ALREADY_APPROVED"
    except Exception as e:
        registrar_screenshot(page, "ver_processo_falha")
        raise RuntimeError(f"Não encontrou/clicou em Ver processo: {e}")


def tentar_processo(page, link_locator, file_path_abs: str, cpf: str, idx: int) -> str | None:
    cpf_safe = _cpf_masked(cpf)
    """Clica em um link 'Ver processo', tenta abrir Exame Admissional e anexar. Retorna sucesso/fracasso."""
    try:
        link_locator.wait_for(timeout=12000)
        try:
            link_locator.scroll_into_view_if_needed(timeout=3000)
        except Exception:
            pass
        link_locator.click(timeout=12000, force=True)
        page.wait_for_timeout(2000)

        # ------------------------------------------------------------------
        # VALIDAÇÃO DE SEGURANÇA: REMOVIDA A PEDIDO (CPF não aparece na tela)
        # ------------------------------------------------------------------
        # logging.info(f"Validacao de CPF desativada. Confiando na busca do nome.")
        # ------------------------------------------------------------------
        abrir_exame_admissional(page)
        return anexar_exame(page, file_path_abs, cpf)
    except Exception as e:
        logging.warning(f"Falha ao processar link Ver processo #{idx} para CPF {cpf_safe}: {e}")
        registrar_screenshot(page, f"ver_processo_falha_{cpf_safe}_{idx}")
        try:
            page.go_back(timeout=8000)
            page.wait_for_timeout(800)
        except Exception:
            pass
        return None


def abrir_exame_admissional(page):
    # Verifica se já está aprovado/validado com texto parcial
    if page.locator("text=Esse documento já foi aprovado").count() > 0:
        logging.info("Detectado 'Esse documento já foi aprovado'.")
        return "IN_REVIEW"
        
    if page.locator("text=Em validação").count() > 0:
         logging.info("Detectado 'Em validação'.")
         return

    try:
        # Tenta seletor genérico por texto (case insensitive se possível no playwright, mas aqui usamos regex ou contains)
        # O locator("text=Exame Admissional") pode falhar se estiver dentro de um span não clicável ou se tiver espaços extras
        
        # Estratégia 1: Botão padrão
        botao = SEL_EXAME_ADM(page)
        
        # Estratégia 2: Texto solto se botão falhar
        if botao.count() == 0:
            botao = page.locator("text='Exame Admissional'")
            
        # Estratégia 3: Contains text (mais permissivo)
        if botao.count() == 0:
             botao = page.locator("xpath=//*[contains(text(), 'Exame Admissional')]")

        
        # Só clica se NÃO estiver expandido
        if page.locator("text=Tirar foto do resultado").is_visible() or \
           page.locator("text=Médico Examinador").is_visible() or \
           page.locator("text=Em validação").is_visible():
             logging.info("Seção Exame Admissional já parece aberta.")
        else:
             if botao.count() > 0:
                 # Clica no último elemento encontrado se houver duplicatas (as vezes o header tem um, e o corpo outro)
                 # Mas cuidado, vamos tentar o primeiro visível
                 found_clickable = False
                 for i in range(botao.count()):
                     if botao.nth(i).is_visible():
                         try:
                             botao.nth(i).click(timeout=8000, force=True)
                             found_clickable = True
                             page.wait_for_timeout(1000)
                             break
                         except:
                             continue
                 if not found_clickable:
                     logging.warning("Elementos 'Exame Admissional' encontrados mas não clicáveis.")
             else:
                 # Nova lógica: Esperar aparecer antes de desistir
                 try:
                     logging.info("Aguardando aba 'Exame Admissional' aparecer...")
                     # Tenta esperar um pouco mais pelo seletor de botão específico
                     SEL_EXAME_ADM(page).first.wait_for(state="visible", timeout=5000)
                     SEL_EXAME_ADM(page).first.click(force=True)
                     page.wait_for_timeout(1000)
                 except:
                     logging.warning("Botão Exame Admissional não encontrado após espera.")

        # Validação final: Está na seção certa?
        indicators = [
            "text=Médico Examinador",
            "text=CRM",
            "text=Tirar foto do resultado",
            "button:has-text('Criar documento')",
            "text=Esse documento já foi aprovado",
            "text=Em validação",
            "button:has-text('Salvar')",
            "button:has-text('Cancelar')",
            "button:has-text('Aprovar documento')",
            "button:has-text('Reprovar documento')"
        ]
        
        for ind in indicators:
            if page.locator(ind).is_visible():
                emit_terminal("OK", "Aba 'Exame Admissional' detectada.", step="upload")
                registrar_screenshot(page, "exame_adm_ok")
                return

        # Tentativa final: Seletor genérico de upload visível
        if page.locator(".ant-upload").is_visible():
             emit_terminal("OK", "Area de upload detectada.", step="upload")
             return

        logging.error("CRITICO: Nao encontrou aba 'Exame Admissional' e contexto nao match.")
        raise RuntimeError("Aba 'Exame Admissional' não encontrada ou não abriu.")

    except Exception as e:
        registrar_screenshot(page, "exame_adm_falha")
        logging.warning(f"Erro ao abrir Exame Admissional: {e}")
        raise e


def anexar_exame(page, file_path: str, cpf: str):
    cpf_safe = _cpf_masked(cpf)
    if not os.path.isfile(file_path):
        raise RuntimeError(f"Arquivo não encontrado para upload: {file_path}")

    # --- VERIFICAÇÃO DE ESTADO JÁ EXISTENTE ---
    if page.locator("text=Esse documento já foi aprovado").count() > 0:
        msg = f"INFO: Pular: Documento JA APROVADO para {cpf_safe}"
        emit_terminal("INFO", msg, step="upload")
        logging.info(msg)
        return "EDIT_EXISTS"
    
    if page.locator("text=Em validação").count() > 0:
        msg = f"INFO: Pular: Documento EM VALIDACAO para {cpf_safe}"
        emit_terminal("INFO", msg, step="upload")
        logging.info(msg)
        return

    if page.locator("button:has-text('Editar documento')").count() > 0:
        msg = f"INFO: Pular: Botao Editar encontrado (ja existe) para {cpf_safe}"
        emit_terminal("INFO", msg, step="upload")
        logging.info(msg)
        return

    # Tenta clicar em Criar Documento
    btn_criar = SEL_CRIAR_DOC(page)
    if btn_criar.count() > 0:
        btn_criar.click()
    else:
        # Se não tem criar e não tem editar/aprovado, talvez já esteja na tela de upload?
        pass

    page.wait_for_timeout(600)
    
    # Tenta input[type=file] diretamente (mais robusto)
    file_set = False
    try:
        file_input = page.locator("input[type='file']").first
        if file_input.count() > 0:
            file_input.scroll_into_view_if_needed(timeout=3000)
            file_input.set_input_files(file_path, timeout=8000)
            file_set = True
    except Exception:
        pass

    if not file_set:
        # fallback: clicar botão e usar file chooser
        try:
            with page.expect_file_chooser(timeout=8000) as fc_info:
                SEL_UPLOAD_BTN(page).click()
            fc = fc_info.value
            fc.set_files(file_path)
            file_set = True
        except Exception as e:
            registrar_screenshot(page, "upload_falha")
            raise RuntimeError(f"Falha ao anexar arquivo (file chooser): {e}")

    # aguarda upload completar antes de salvar
    page.wait_for_timeout(PRE_SAVE_WAIT_MS)

    # espera botão Salvar habilitar e clica (retries)
    clicked_save = False
    modal_clicked = False
    for _ in range(3):
        try:
            salvar_btn = SEL_SALVAR(page)
            salvar_btn.wait_for(state="visible", timeout=6000)
            try:
                salvar_btn.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            try:
                salvar_btn.wait_for(state="enabled", timeout=5000)
            except Exception:
                pass
            salvar_btn.click(timeout=8000, force=True)
            clicked_save = True
        except Exception:
            logging.debug(f"Tentativa {_} de clicar em SALVAR falou ou não estava habilitado.")
            page.wait_for_timeout(1000)
            
            # Checagem de "Já existe" / Modal de erro
            if page.locator("text=Documento já existe").count() > 0:
                logging.info(f"Yube informou que documento já existe para {cpf_safe}.")
                # Clica em cancelar ou fechar se precisar
                return
            continue

        # tenta modal "Criar"
        try:
            confirmar_btn = page.get_by_role("button", name=re.compile("Criar", re.I))
            confirmar_btn.wait_for(timeout=5000)
            try:
                confirmar_btn.scroll_into_view_if_needed(timeout=2000)
            except Exception:
                pass
            confirmar_btn.click(timeout=5000, force=True)
            modal_clicked = True
            try:
                page.wait_for_selector("text=Confirmação", state="hidden", timeout=8000)
            except Exception:
                pass
        except Exception:
            pass

        page.wait_for_timeout(2000)
        break

    if not clicked_save:
        logging.debug("Botão Salvar não encontrado/clicável (após tentativas).")

    # aguarda conclusão do upload/salvar e registra evidência
    page.wait_for_timeout(UPLOAD_WAIT_MS)
    registrar_screenshot(page, f"upload_pos_{cpf_safe}")
    # aguarda pós-salvar para garantir processamento
    page.wait_for_timeout(POST_SAVE_WAIT_MS)
    return "UPLOADED"


def _build_nome_tentativas(nome_hint: str | None, cpf: str | None) -> list[str]:
    tentativas: list[str] = []
    seen: set[str] = set()
    generic_terms = {
        "exame admissional",
        "aso admissional",
        "exame",
        "admissional",
    }

    def _push(value: str | None):
        raw = (value or "").strip()
        key = raw.lower()
        if not raw or key in seen:
            return
        seen.add(key)
        tentativas.append(raw)

    tokens = [t for t in re.split(r"\s+", (nome_hint or "").strip()) if t]
    suffixes = {"junior", "jr", "filho", "neto", "sobrinho"}
    nome_hint_norm = " ".join(tokens).lower()

    if nome_hint_norm not in generic_terms:
        _push(nome_hint)

    if tokens and nome_hint_norm not in generic_terms:
        if len(tokens) >= 2:
            _push(" ".join(tokens[:2]))
        if len(tokens) >= 3:
            _push(" ".join(tokens[:3]))
            _push(f"{tokens[0]} {tokens[-1]}")
            _push(f"{tokens[0]} {tokens[1]} {tokens[-1]}")

        if tokens[-1].lower() in suffixes and len(tokens) >= 2:
            base_tokens = tokens[:-1]
            _push(" ".join(base_tokens))
            if len(base_tokens) >= 2:
                _push(f"{base_tokens[0]} {base_tokens[-1]}")
            if len(base_tokens) >= 3:
                _push(" ".join(base_tokens[:3]))

    _push(cpf)
    return tentativas


def processar_arquivo(page, file_path: str):
    filename = os.path.basename(file_path)
    nome_hint = os.path.splitext(filename)[0].split(" - ")[0].strip()
    cpf = extrair_cpf_do_nome(filename)
    cpf_safe = _cpf_masked(cpf)

    emit_terminal("INFO", "Processando arquivo na Yube.", step="processamento", extra={"arquivo": filename, "cpf": cpf_safe})
    
    # Mover arquivo para "em processamento" antes de iniciar
    arquivo_em_processamento = None
    try:
        if os.path.exists(file_path):
            arquivo_em_processamento = os.path.join(PASTA_EM_PROCESSAMENTO, filename)
            # Se já existe na pasta de processamento, usa ele
            if not os.path.exists(arquivo_em_processamento):
                shutil.move(file_path, arquivo_em_processamento)
            file_path = arquivo_em_processamento  # Atualiza referência
    except Exception as e:
        logging.warning(f"Erro ao mover para em processamento: {e}")
        # Continua com o arquivo original se não conseguir mover
    
    if not cpf:
        logging.error(f"CPF n?o encontrado no nome do arquivo: {filename}")
        registrar_log("", file_path, "erro", "CPF nao encontrado no nome")
        destino_erro = None
        try:
            destino_erro = os.path.join(PASTA_ERROS, filename)
            if os.path.exists(file_path):
                shutil.move(file_path, destino_erro)
        except Exception as e:
            logging.error(f"Erro ao mover para erros: {e}")
        return (False, "CPF_NOT_FOUND", destino_erro or file_path)

    # 1. Busca por nome com variacoes + fallback por CPF
    nome_tentativas = _build_nome_tentativas(nome_hint, cpf)
    
    ok_busca = False
    for tentativa in nome_tentativas:
        if not tentativa: continue
        logging.info(f"INFO: Buscando por: '{tentativa}'")
        nome_busca = None if (cpf and tentativa == cpf) else tentativa
        if pesquisar_funcionario_robusto(page, cpf, nome_hint=nome_busca):
            ok_busca = True
            break
        else:
            logging.warning(f"Busca falhou para '{tentativa}'.")

    if not ok_busca:
        emit_terminal("WARNING", "Funcionario nao encontrado na busca.", step="processamento", extra={"cpf": cpf_safe, "tentativas": nome_tentativas})
        # Se falhou tudo, salva erro e vai pro pr?ximo
        registrar_log(cpf, file_path, "erro", "Funcionario nao encontrado na busca")
        destino_erro = None
        try:
            destino_erro = os.path.join(PASTA_ERROS, filename)
            if os.path.exists(file_path):
                shutil.move(file_path, destino_erro)
        except Exception as e:
            logging.error(f"Erro ao mover para erros: {e}")
        return (False, "NOT_FOUND", destino_erro or file_path)
    try:
        links_vp = SEL_VER_PROCESSO(page)
        total_links = links_vp.count()
        if total_links == 0:
            raise RuntimeError("Nenhum link 'Ver processo' encontrado.")

        resultado_upload = None
        for i in range(total_links):
            link = links_vp.nth(i)
            resultado_upload = tentar_processo(page, link, os.path.abspath(file_path), cpf, i)
            if resultado_upload:
                break

        if not resultado_upload:
            raise RuntimeError("Nenhum processo abriu ou permitiu anexar.")
    except Exception as e:
        logging.exception(f"Falha no fluxo para {cpf_safe}: {e}")
        registrar_log(cpf, file_path, "erro", str(e))
        destino_erro = None
        try:
            destino_erro = os.path.join(PASTA_ERROS, filename)
            if os.path.exists(file_path):
                shutil.move(file_path, destino_erro)
            else:
                logging.error(f"Arquivo original sumiu antes da c?pia para erro: {file_path}")
        except Exception as e2:
            logging.error(f"Erro ao mover para erros: {e2}")
        
        # tentativa de screenshot para depurar
        try:
            page.screenshot(path=os.path.join(PASTA_LOGS_RPA, f"error_{cpf_safe}_{int(time.time())}.png"))
        except Exception:
            pass
        return (False, "ERROR", destino_erro or file_path)

    if resultado_upload == "UPLOADED":
        registrar_log(cpf, file_path, "sucesso", "arquivo anexado")
    else:
        skip_msgs = {
            "ALREADY_APPROVED": "Documento ja aprovado no Yube",
            "IN_REVIEW": "Documento em validacao no Yube",
            "EDIT_EXISTS": "Documento ja existente (Editar documento)",
            "DOC_ALREADY_EXISTS": "Documento ja existe no Yube",
        }
        registrar_log(cpf, file_path, "pulado", skip_msgs.get(resultado_upload, "Documento pulado no Yube"))
    try:
        destino_processado = os.path.join(PASTA_PROCESSADOS, filename)
        if os.path.exists(file_path):
            shutil.move(file_path, destino_processado)
        else:
            logging.error(f"Arquivo original sumiu antes da cópia para processados: {file_path}")
    except Exception as e:
        logging.error(f"Erro ao mover para processados: {e}")

    # ==================================================
    # VOLTAR PARA A BUSCA - ESTRATÉGIA OTIMIZADA
    # ==================================================
    try:
        # 0. Tenta fechar Drawer/Modal (X) - Solução para tela de upload
        # Seletores comuns de fechar no AntDesign/Yube
        close_btn = page.locator("button[aria-label='Close'], .ant-drawer-close, button.ant-modal-close, svg[data-icon='close']")
        if close_btn.count() > 0 and close_btn.first.is_visible():
             logging.info("Fechando modal/drawer pelo 'X'.")
             close_btn.first.click(timeout=3000, force=True)
             page.wait_for_timeout(1000)

        # 1. Tenta botão 'Voltar' clássico (rápido)
        voltar_btn = SEL_VOLTAR(page)
        if voltar_btn.count() > 0 and voltar_btn.is_visible():
            voltar_btn.click(timeout=3000, force=True)
        
        # 2. Se não clicou ou não achou, tenta link 'Início'
        elif page.get_by_role("link", name="Início").count() > 0:
            logging.info("Usando botão 'Início' para voltar.")
            page.get_by_text("Início").first.click(timeout=3000, force=True)
            
        # 3. Se nada funcionou, browser back
        else:
            logging.info("Indo via Browser Back.")
            page.go_back(timeout=5000)

        # Aguarda campo de busca para confirmar que voltou
        SEL_BUSCA(page).wait_for(timeout=8000)

    except Exception as e:
        logging.warning(f"Navegação de volta falhou ({e}). Forçando reload da Home.")
        # fallback de segurança (último caso)
        page.goto(YUBE_URL, timeout=20000, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)
        filtrar_todas_obras(page)
        SEL_BUSCA(page).wait_for(timeout=15000)

    if resultado_upload == "UPLOADED":
        return (True, None, None)
    return (False, f"SKIPPED_{resultado_upload}", None)



def process_folder(base_path, headless=False, max_files=None, specific_files=None):
    start_time = time.time()
    stats = {
        'total': 0,
        'sucessos': [],
        'pulados': [],
        'erros': [],
        'tempo_total': ''
    }

    with sync_playwright() as p:
        # ... (setup do browser) ...
        try:
            browser = p.chromium.launch(
                headless=headless,
                args=[
                    "--ignore-certificate-errors",
                    "--disable-http2",
                    "--disable-features=AllowInsecureLocalhost,SSLVersionFallback",
                    "--no-sandbox",
                    "--disable-web-security",
                ],
            )
        except Exception as e:
            msg = str(e)
            if "Executable doesn't exist" in msg or "playwright install" in msg:
                logging.warning("Chromium do Playwright nao encontrado. Instalando automaticamente...")
                if _ensure_playwright_chromium_installed():
                    browser = p.chromium.launch(
                        headless=headless,
                        args=[
                            "--ignore-certificate-errors",
                            "--disable-http2",
                            "--disable-features=AllowInsecureLocalhost,SSLVersionFallback",
                            "--no-sandbox",
                            "--disable-web-security",
                        ],
                    )
                else:
                    raise
            else:
                raise
        context = browser.new_context(ignore_https_errors=True)
        page = context.new_page()
        page = login(page)
        filtrar_todas_obras(page)

        if specific_files:
            files = specific_files
            logging.info(f"Processando {len(files)} arquivos ESPECÍFICOS enviados pelo main.")
        else:
            files = [str(p) for p in Path(base_path).glob("*.pdf")]
            logging.info(f"{len(files)} arquivos encontrados em {base_path} (modo varredura).")
        
        if max_files:
            files = files[:max_files]

        stats['total'] = len(files)

        retry_queue = []

        def _normalize_result(result):
            if isinstance(result, tuple) and len(result) == 3:
                return result
            return (bool(result), None, None)

        def _erro_msg(reason):
            if reason == "NOT_FOUND":
                return "Funcionario nao encontrado na busca"
            if reason == "CPF_NOT_FOUND":
                return "CPF nao encontrado no nome"
            if reason == "SKIPPED_ALREADY_APPROVED":
                return "Documento ja aprovado no Yube"
            if reason == "SKIPPED_IN_REVIEW":
                return "Documento em validacao no Yube"
            if reason == "SKIPPED_EDIT_EXISTS":
                return "Documento ja existente (Editar documento)"
            if reason == "SKIPPED_DOC_ALREADY_EXISTS":
                return "Documento ja existe no Yube"
            return "Falha na busca ou anexo (vide logs)"

        for f in files:
            filename = os.path.basename(f)
            try:
                ok, reason, retry_path = _normalize_result(processar_arquivo(page, f))
                if ok:
                    stats["sucessos"].append(filename)
                elif reason and str(reason).startswith("SKIPPED_"):
                    stats["pulados"].append({"arquivo": filename, "motivo": _erro_msg(reason)})
                else:
                    stats["erros"].append({"arquivo": filename, "erro": _erro_msg(reason)})
                    if reason == "NOT_FOUND":
                        retry_queue.append(retry_path or f)
            except Exception as e:
                logging.exception(f"Erro ao processar arquivo {f}: {e}")
                stats["erros"].append({"arquivo": filename, "erro": str(e)})

        if retry_queue and RETRY_NOT_FOUND > 0:
            for attempt in range(RETRY_NOT_FOUND):
                logging.info(f"Retry busca (nao encontrados) {attempt + 1}/{RETRY_NOT_FOUND}: {len(retry_queue)} arquivos")
                if RETRY_NOT_FOUND_DELAY_SEC > 0:
                    time.sleep(RETRY_NOT_FOUND_DELAY_SEC)
                try:
                    page.goto(YUBE_URL, timeout=20000, wait_until="domcontentloaded")
                    page.wait_for_timeout(2000)
                    filtrar_todas_obras(page)
                    SEL_BUSCA(page).wait_for(timeout=15000)
                except Exception as e:
                    logging.warning(f"Falha ao preparar retry: {e}")

                next_retry = []
                for f in retry_queue:
                    filename = os.path.basename(f)
                    try:
                        ok, reason, retry_path = _normalize_result(processar_arquivo(page, f))
                        if ok:
                            if filename not in stats["sucessos"]:
                                stats["sucessos"].append(filename)
                            stats["erros"] = [e for e in stats["erros"] if e.get("arquivo") != filename]
                            stats["pulados"] = [e for e in stats["pulados"] if e.get("arquivo") != filename]
                        elif reason and str(reason).startswith("SKIPPED_"):
                            if not any(e.get("arquivo") == filename for e in stats["pulados"]):
                                stats["pulados"].append({"arquivo": filename, "motivo": _erro_msg(reason)})
                            stats["erros"] = [e for e in stats["erros"] if e.get("arquivo") != filename]
                        else:
                            if reason == "NOT_FOUND":
                                next_retry.append(retry_path or f)
                    except Exception as e:
                        logging.exception(f"Erro ao reprocessar arquivo {f}: {e}")
                retry_queue = next_retry
                if not retry_queue:
                    break
        if KEEP_BROWSER_OPEN:
            emit_terminal("INFO", "Navegador permanecera aberto (KEEP_BROWSER_OPEN=1).", step="encerramento")
        else:
            browser.close()

    end_time = time.time()
    elapsed = end_time - start_time
    stats['tempo_total'] = str(timedelta(seconds=int(elapsed)))
    
    return stats


# função para integrar com main.py
def run_from_main(base_path, files_to_process=None):
    return process_folder(base_path, headless=False, specific_files=files_to_process)


# Se executado direto:
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        emit_terminal("INFO", "Uso: python rpa_yube.py <base_path>", step="cli")
        emit_terminal("INFO", "Ex: python rpa_yube.py \"P:\\ASO\\Obra_999\\2025-12-11\"", step="cli")
        sys.exit(1)
    base = sys.argv[1]
    process_folder(base, headless=False)
