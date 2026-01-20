import os
import re
import time
import shutil
import logging
import csv
from pathlib import Path
from datetime import datetime, timedelta
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

load_dotenv()  # opcional: carregar YUBE_USER/YUBE_PASS se quiser

# ---------- CONFIGURAÇÃO ----------
YUBE_URL = os.getenv("YUBE_URL", "https://yube.com.br/")
YUBE_USER = os.getenv("YUBE_USER", "priscilla.silva@enesa.com.br")
YUBE_PASS = os.getenv("YUBE_PASS", "Enesa@2024")
NAV_TIMEOUT = int(os.getenv("YUBE_NAV_TIMEOUT", "10000"))
KEEP_BROWSER_OPEN = os.getenv("KEEP_BROWSER_OPEN", "1") == "1"
UPLOAD_WAIT_MS = int(os.getenv("UPLOAD_WAIT_MS", "4000"))
POST_SAVE_WAIT_MS = int(os.getenv("POST_SAVE_WAIT_MS", "3000"))
PRE_SAVE_WAIT_MS = int(os.getenv("PRE_SAVE_WAIT_MS", "5000"))

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
        logging.FileHandler(os.path.join(PASTA_LOGS_RPA, "rpa_yube_debug.log"), encoding='utf-8'),
        logging.StreamHandler()
    ]
)

LOG_CSV = os.path.join(PASTA_LOGS_RPA, "rpa_log.csv")

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
SEL_EXAME_ADM = lambda page: page.get_by_role("button", name=re.compile(r"(Exame.*Admissional|Sa.de.*Ocupacional|ASO)", re.I))
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


def registrar_log(cpf, file_path, status, message=""):
    header = ["timestamp", "cpf", "file", "status", "message"]
    row = [datetime.now().isoformat(), cpf, file_path, status, message]
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
        print("Não consegui abrir a página automaticamente. Abra manualmente o login da Yube nesta janela e pressione Enter.")
        input()

    target_page = page

    target_page = page
    
    # Verifica se JÁ estamos na tela de login (inputs visíveis)
    try:
        if SEL_EMAIL(page).count() > 0:
             print("ℹ️  Já estamos na tela de login. Prosseguindo...")
    except:
        pass

    # Se NÃO tem input de senha, força a navegação para login direto (evita nova aba)
    if SEL_EMAIL(page).count() == 0:
        try:
            print("🔄  Redirecionando para página de login direta...")
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
    print("🔑  Preenchendo credenciais...")
    try:
        # Se ainda assim não aparecer, tenta esperar
        target_page.wait_for_selector("#username", state="visible", timeout=10000)
        SEL_EMAIL(target_page).fill(YUBE_USER)
        target_page.wait_for_selector("input[name='password']", state="visible", timeout=10000)
        SEL_SENHA(target_page).fill(YUBE_PASS)
        SEL_ENTRAR(target_page).click()
        print("✅  Login submetido.")
    except Exception as e:
        # Se falhar o login, mas já estiver logado (redirecionou para home), tudo bem
        if SEL_BUSCA(target_page).count() > 0:
             print("ℹ️  Parece que já estamos logados (Busca visível).")
        else:
             logging.warning(f"Falha ao preencher login: {e}")
             print("❌  Erro ao preencher login.")

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
                return
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
         
    page.wait_for_timeout(2500)
    registrar_screenshot(page, f"busca_input_{cpf}")

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
                registrar_screenshot(page, f"busca_{cpf}")
                logging.info(f"Selecionando resultado da busca para CPF {cpf}: {texto_card or cpf}")
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
    
    registrar_screenshot(page, f"busca_falha_{cpf}")
    logging.warning(f"Busca falhou para '{termo_busca}'. Nenhum resultado exato encontrado.")
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
        return
    except Exception as e:
        registrar_screenshot(page, "ver_processo_falha")
        raise RuntimeError(f"Não encontrou/clicou em Ver processo: {e}")


def tentar_processo(page, link_locator, file_path_abs: str, cpf: str, idx: int) -> bool:
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
        anexar_exame(page, file_path_abs, cpf)
        return True
    except Exception as e:
        logging.warning(f"Falha ao processar link Ver processo #{idx} para CPF {cpf}: {e}")
        registrar_screenshot(page, f"ver_processo_falha_{cpf}_{idx}")
        try:
            page.go_back(timeout=8000)
            page.wait_for_timeout(800)
        except Exception:
            pass
        return False


def abrir_exame_admissional(page):
    # Verifica se já está aprovado/validado com texto parcial
    if page.locator("text=Esse documento já foi aprovado").count() > 0:
        logging.info("Detectado 'Esse documento já foi aprovado'.")
        return
        
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
                print("✅  Aba 'Exame Admissional' detectada/aberta.")
                registrar_screenshot(page, "exame_adm_ok")
                return

        # Tentativa final: Seletor genérico de upload visível
        if page.locator(".ant-upload").is_visible():
             print("✅  Área de upload detectada.")
             return

        logging.error("CRÍTICO: Não encontrou aba 'Exame Admissional' e contexto não match.")
        raise RuntimeError("Aba 'Exame Admissional' não encontrada ou não abriu.")

    except Exception as e:
        registrar_screenshot(page, "exame_adm_falha")
        logging.warning(f"Erro ao abrir Exame Admissional: {e}")
        raise e


def anexar_exame(page, file_path: str, cpf: str):
    if not os.path.isfile(file_path):
        raise RuntimeError(f"Arquivo não encontrado para upload: {file_path}")

    # --- VERIFICAÇÃO DE ESTADO JÁ EXISTENTE ---
    if page.locator("text=Esse documento já foi aprovado").count() > 0:
        msg = f"⏩  Pular: Documento JÁ APROVADO para {cpf}"
        print(msg)
        logging.info(msg)
        return
    
    if page.locator("text=Em validação").count() > 0:
        msg = f"⏩  Pular: Documento EM VALIDAÇÃO para {cpf}"
        print(msg)
        logging.info(msg)
        return

    if page.locator("button:has-text('Editar documento')").count() > 0:
        msg = f"⏩  Pular: Botão Editar encontrado (já existe) para {cpf}"
        print(msg)
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
                logging.info(f"Yube informou que documento já existe para {cpf}.")
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
    registrar_screenshot(page, f"upload_pos_{cpf}")
    # aguarda pós-salvar para garantir processamento
    page.wait_for_timeout(POST_SAVE_WAIT_MS)


def processar_arquivo(page, file_path: str):
    filename = os.path.basename(file_path)
    nome_hint = os.path.splitext(filename)[0].split(" - ")[0].strip()
    cpf = extrair_cpf_do_nome(filename)

    print(f"📄  Processando: {filename} (CPF: {cpf})")
    
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
        logging.error(f"CPF não encontrado no nome do arquivo: {filename}")
        registrar_log("", file_path, "erro", "CPF nao encontrado no nome")
        try:
            destino_erro = os.path.join(PASTA_ERROS, filename)
            if os.path.exists(file_path):
                shutil.move(file_path, destino_erro)
        except Exception as e:
            logging.error(f"Erro ao mover para erros: {e}")
        return False

    # 1. Busca STRICT MODE (Apenas nome hint completo)
    nome_tentativas = [nome_hint]
    
    ok_busca = False
    for tentativa in nome_tentativas:
        if not tentativa: continue
        logging.info(f"🔎 Buscando por: '{tentativa}'")
        if pesquisar_funcionario(page, cpf, nome_hint=tentativa):
            ok_busca = True
            break
        else:
            logging.warning(f"Busca falhou para '{tentativa}'.")

    if not ok_busca:
        print(f"❌  Funcionário não encontrado: {cpf} (Tentativas: {nome_tentativas})")
        # Se falhou tudo, salva erro e vai pro próximo
        registrar_log(cpf, file_path, "erro", "Funcionario nao encontrado na busca")
        try:
            destino_erro = os.path.join(PASTA_ERROS, filename)
            if os.path.exists(file_path):
                shutil.move(file_path, destino_erro)
        except Exception as e:
            logging.error(f"Erro ao mover para erros: {e}")
        return False
    try:
        links_vp = SEL_VER_PROCESSO(page)
        total_links = links_vp.count()
        if total_links == 0:
            raise RuntimeError("Nenhum link 'Ver processo' encontrado.")

        sucesso = False
        for i in range(total_links):
            link = links_vp.nth(i)
            if tentar_processo(page, link, os.path.abspath(file_path), cpf, i):
                sucesso = True
                break

        if not sucesso:
            raise RuntimeError("Nenhum processo abriu ou permitiu anexar.")
    except Exception as e:
        logging.exception(f"Falha no fluxo para {cpf}: {e}")
        registrar_log(cpf, file_path, "erro", str(e))
        try:
            destino_erro = os.path.join(PASTA_ERROS, filename)
            if os.path.exists(file_path):
                shutil.move(file_path, destino_erro)
            else:
                logging.error(f"Arquivo original sumiu antes da cópia para erro: {file_path}")
        except Exception as e2:
            logging.error(f"Erro ao mover para erros: {e2}")
        
        # tentativa de screenshot para depurar
        try:
            page.screenshot(path=os.path.join(PASTA_LOGS_RPA, f"error_{cpf}_{int(time.time())}.png"))
        except Exception:
            pass
        return False

    registrar_log(cpf, file_path, "sucesso", "arquivo anexado")
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

    return True



def process_folder(base_path, headless=False, max_files=None, specific_files=None):
    start_time = time.time()
    stats = {
        'total': 0,
        'sucessos': [],
        'erros': [],
        'tempo_total': ''
    }

    with sync_playwright() as p:
        # ... (setup do browser) ...
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

        for f in files:
            filename = os.path.basename(f)
            try:
                ok = processar_arquivo(page, f)
                if ok:
                    stats['sucessos'].append(filename)
                else:
                    # Se processar_arquivo retornou False, assumimos que houve falha (ou não encontrado)
                    # O ideal seria processar_arquivo retornar (bool, reason), mas por simplicidade:
                    stats['erros'].append({'arquivo': filename, 'erro': 'Falha na busca ou anexo (vide logs)'})
            except Exception as e:
                logging.exception(f"Erro ao processar arquivo {f}: {e}")
                stats['erros'].append({'arquivo': filename, 'erro': str(e)})

        if KEEP_BROWSER_OPEN:
            print("Navegador permanecerá aberto (KEEP_BROWSER_OPEN=1).")
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
        print("Uso: python rpa_yube.py <base_path>")
        print("Ex: python rpa_yube.py \"P:\\ASO\\Obra_999\\2025-12-11\"")
        sys.exit(1)
    base = sys.argv[1]
    process_folder(base, headless=False)