import win32com.client as win32
import os
import re
import shutil
from dotenv import load_dotenv

# Carrega variáveis de ambiente do arquivo .env (se existir)
load_dotenv()
import pytesseract
from pdf2image import convert_from_path
from datetime import datetime, timedelta
import traceback
import hashlib
from rpa_yube import run_from_main

# Módulos customizados
from custom_logger import RpaLogger
from reporting import ReportGenerator
from notification import enviar_resumo_email

# ----------------------------
# CONFIGURAÇÕES
# ----------------------------

def find_tesseract():
    candidate = None
    
    # 1. Tenta pegar do .env
    env_path = os.getenv("TESSERACT_PATH")
    if env_path:
        candidate = env_path

    # Separado: se o candidato existe...
    if candidate and os.path.exists(candidate):
        if os.path.isdir(candidate):
             return os.path.join(candidate, "tesseract.exe")
        return candidate
        
    # 2. Tenta caminhos comuns
    common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
    ]
    for p in common_paths:
        if os.path.exists(p):
            return p
            
    # 3. Tenta pelo PATH do sistema
    which_tess = shutil.which("tesseract")
    if which_tess:
        return which_tess
        
    return None

tess_path = find_tesseract()
if tess_path:
    print(f"DEBUG: Configured Tesseract Path: '{tess_path}'")
    pytesseract.pytesseract.tesseract_cmd = tess_path
else:
    print("AVISO: Tesseract não encontrado automaticamente. Verifique se está instalado ou configure TESSERACT_PATH no arquivo .env")
    # Mantém o default da biblioteca ou define um fallback se preferir, 
    # mas o aviso ajuda o usuário a saber o que houve.

# Poppler também pode vir do .env
POPPLER_PATH = os.getenv("POPPLER_PATH", r"P:\ASO\Release-24.08.0-0\poppler-24.08.0\Library\bin")

PASTA_BASE = r"P:\ProcessoASO"
PASTA_PROCESSADOS = os.path.join(PASTA_BASE, "processados")
PASTA_EM_PROCESSAMENTO = os.path.join(PASTA_BASE, "em processamento")
PASTA_ERROS = os.path.join(PASTA_BASE, "erros")
PASTA_LOGS = os.path.join(PASTA_BASE, "logs")
EMAIL_DESEJADO = os.getenv("ASO_EMAIL_ACCOUNT", "aso@enesa.com.br")
MAILBOX_NAME = os.getenv("ASO_MAILBOX_NAME", "Aso")  # nome exibido na árvore

os.makedirs(PASTA_PROCESSADOS, exist_ok=True)
os.makedirs(PASTA_EM_PROCESSAMENTO, exist_ok=True)
os.makedirs(PASTA_ERROS, exist_ok=True)
os.makedirs(PASTA_LOGS, exist_ok=True)
PASTA_RELATORIOS = os.path.join(PASTA_BASE, "relatorios")
os.makedirs(PASTA_RELATORIOS, exist_ok=True)

# Inicializa logger e reporter
logger = RpaLogger(PASTA_LOGS)
reporter = ReportGenerator(PASTA_RELATORIOS)

TARGET_ACCOUNT = os.getenv("ASO_EMAIL_ACCOUNT", "aso@enesa.com.br")


# ====================================================================
# FUNÇÃO DE LOG (Wrapper para o logger estruturado)
# ====================================================================
def registrar_log(msg):
    # Mantém compatibilidade com chamadas existentes
    logger.info(msg)


# ====================================================================
# HASH DE ARQUIVO
# ====================================================================
def calcular_hash_arquivo(caminho):
    hash_md5 = hashlib.md5()
    try:
        with open(caminho, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception:
        return None


def eh_aso(texto_ocr):
    padroes_aso = [
        r"\bASO\b",
        r"Atestado de Saúde Ocupacional",
        r"ATESTADO DE SAÚDE OCUPACIONAL",
        r"Atestado\s+de\s+Sa[úu]de\s+Ocupacional",
    ]
    for p in padroes_aso:
        if re.search(p, texto_ocr, re.IGNORECASE):
            return True
    return False


# ====================================================================
# FUNÇÃO - EXTRAI DADOS COMPLETOS (OCR)
# ====================================================================
def extrair_dados_completos(img):
    """
    Retorna: nome, cpf, data_aso, funcao_cargo
    """
    try:
        texto = pytesseract.image_to_string(img, lang="por+eng")
        # Checagem de Rascunho
        if "Rascunho" in texto and texto.upper().count("RASCUNHO") > 3:
             print("ℹ️  Arquivo identificado como RASCUNHO. Ignorando.")
             return "RASCUNHO", "Ignorar", "", ""
    except Exception as e:
        texto = ""
        registrar_log(f"Erro no pytesseract: {e}")

    # Normalizar um pouco o texto para buscas
    texto_compacto = re.sub(r"\s+", " ", texto)

    # DEBUG: Mostrar primeiros caracteres para verificar qualidade
    if len(texto) > 0:
        clean_debug = texto_compacto[:300].encode('ascii', 'ignore').decode('ascii')
        print(f"DEBUG OCR RAW (start): {clean_debug}...")
    else:
        print("DEBUG OCR RAW: Vazio!")

    # ---------------- CPF ----------------
    cpf = "CPF_Desconhecido"

    # Tentativa 1: Com label "CPF"
    cpf_match = re.search(r"CPF[:\s]*([\d.\- ]{11,20})", texto, re.IGNORECASE)
    if cpf_match:
        numeros = re.sub(r"\D", "", cpf_match.group(1))
        if len(numeros) >= 11:
            numeros = numeros[-11:]
            cpf = f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"

    # Tentativa 2: Busca direta pelo formato XXX.XXX.XXX-XX (sem label)
    if cpf == "CPF_Desconhecido":
        cpf_strict = re.search(r"(\d{3}\.\d{3}\.\d{3}-\d{2})", texto)
        if cpf_strict:
            cpf = cpf_strict.group(1)

    if cpf == "CPF_Desconhecido":
        # Tentativa 3: Qualquer sequência de 11 digitos
        digitos = re.findall(r"\d{11}", texto)
        if digitos:
            numeros = digitos[0]
            cpf = f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"

    # ---------------- NOME (robusto p/ múltiplos modelos) ----------------
    nome = "Desconhecido"

    padroes_nome = [
        # Nome explícito
        r"Nome\s*Completo[:\s\-]*([A-Z\u00C0-\u00DD][A-Z\u00C0-\u00DD \-\.]{3,150})",
        r"Nome\s*Comp[li1]eto[:\s\-]*([A-Z\u00C0-\u00DD][A-Z\u00C0-\u00DD \-\.]{3,150})",
        r"Nome[:\s\-]*([A-Z\u00C0-\u00DD][A-Z\u00C0-\u00DD \-\.]{3,150})",

        # Funcionário na mesma linha (permissivo)
        r"Funcion.*rio[:\s\-]*([A-Z\u00C0-\u00DD][A-Z\u00C0-\u00DD \-\.]{3,150})",

        # Funcionário COM nome NA LINHA ABAIXO (seguro)
        r"(?m)^Funcion.*rio[^\n]*\n\s*([A-Z\u00C0-\u00DD][A-Z\u00C0-\u00DD \-\.]{3,150})",
    ]

    for padrao in padroes_nome:
        match = re.search(padrao, texto, re.IGNORECASE)
        # Ignorar matches que sejam apenas "M" ou "F" isolados (sexo)
        if match:
            candidato = match.group(1).strip()
            if len(candidato) > 2: 
                nome = candidato
                break
            
    # Debug para casos falhos
    if nome == "Desconhecido":
        registrar_log(f"DEBUG OCR FALHO (CPF={cpf}): Texto parcial: {texto[:200].replace(chr(10), ' ')}")
        # Fallback: nome imediatamente antes do CPF
        bloco = re.search(
            r"([A-ZÀ-Ý][A-ZÀ-Ý \-]{5,150})\s+(?:CPF|C\.P\.F)",
            texto
        )
        if bloco:
            nome = bloco.group(1).strip()

    # >>> HEURÍSTICA: Se ainda desconhecido, procurar nas linhas acima do CPF
    if nome == "Desconhecido":
        linhas = texto.splitlines()
        idx_cpf = -1
        # Achar linha do CPF
        for i, line in enumerate(linhas):
            if re.search(r"CPF[:\s]*[\d\.]", line, re.IGNORECASE) or re.search(r"\d{3}\.\d{3}\.\d{3}\-\d{2}", line):
                idx_cpf = i
                
                # Tentar pegar nome na MESMA linha do CPF (ex: "JOAO DA SILVA CPF 123...")
                # Remove o CPF e "CPF:" da linha
                line_clean = re.sub(r"CPF[:\s]*[\d\.\-]+", "", line, flags=re.IGNORECASE)
                line_clean = re.sub(r"DATA.*", "", line_clean, flags=re.IGNORECASE).strip()
                if len(line_clean) > 5 and re.match(r"^[A-Z\u00C0-\u00DD][A-Z\u00C0-\u00DD \-\.]{3,}$", line_clean):
                    nome = line_clean
                break
        
        # Olhar até 3 linhas acima
        if nome == "Desconhecido" and idx_cpf > 0:
            for offset in range(1, 4):
                if idx_cpf - offset < 0: break
                candidate_line = linhas[idx_cpf - offset].strip()
                # Remove lixo comum
                candidate_line = re.sub(r"DATA.*", "", candidate_line, flags=re.IGNORECASE).strip()
                
                # BLACKLIST DE PALAVRAS QUE NÃO SÃO NOME
                # Se a linha contiver esses termos, provavelmente é cabeçalho ou label
                blacklist = ["CARGO", "FUNCAO", "FUNÇÃO", "SETOR", "DEPARTAMENTO", "ADMISSAO", "DEMISSAO", "ASO", "SAUDE", "OCUPACIONAL", "COODENADOR", "MÉDICO", "MEDICO", "EXAMINADOR"]
                if any(bad in candidate_line.upper() for bad in blacklist):
                    continue

                # Critério: deve ter pelo menos 2 palavras (ex: "JOAO SILVA"), letras maiúsculas, sem números excessivos
                # FRANCICLE pode ser um nome único se for monônimo, mas raro. Vamos exigir tamanho > 3.
                if len(candidate_line) > 3 and not re.search(r"\d", candidate_line):
                     # Verificar se parece nome (Maiúsculas e pelo menos um espaço, ou palavra longa única)
                     # Ex: "FRANCICLE SILVA" -> OK. "FRANCICLE" -> Talvez OK se não tiver nada melhor.
                     if re.match(r"^[A-Z\u00C0-\u00DD][A-Z\u00C0-\u00DD \-\.]+$", candidate_line):
                         # Evita pegar palavras soltas curtas
                         if " " not in candidate_line.strip() and len(candidate_line) < 8:
                             # Ex: "NOME" sozinho, ou "ESTADO"
                             continue
                             
                         nome = candidate_line
                         break

    # Limpeza final
    nome = re.sub(r"\b(PCD|SIM|NAO|NÃO|APTO|INAPTO)\b", "", nome, flags=re.IGNORECASE)
    nome = re.sub(r"\s{2,}", " ", nome).strip()

    if not nome:
        nome = "Desconhecido"


    # Fallback forte: texto antes do CPF (tentativa regex final)
    if nome == "Desconhecido":
        bloco = re.search(
            r"([A-ZÀ-Ý][A-ZÀ-Ý \-]{3,150})\s+(?:CPF|C\.P\.F)",
            texto
        )
        if bloco:
            nome = bloco.group(1).strip()

    # Limpeza final
    nome = re.sub(r"[^A-Za-zÀ-ÿ\s\-]", "", nome).strip()
    nome = re.sub(r"\bCPF\b", "", nome, flags=re.IGNORECASE).strip()

    if not nome:
        nome = "Desconhecido"
        
    # --- DIAGNÓSTICO PARA O USUÁRIO ---
    if cpf == "CPF_Desconhecido":
        print(f"⚠️  ALERTA DE LEITURA: Não foi possível ler o CPF neste arquivo.")
        registrar_log(f"DEBUG OCR: Falha CPF. Texto parcial: {texto[:100]}")
    elif nome == "Desconhecido":
        print(f"⚠️  ALERTA DE LEITURA: CPF encontrato ({cpf}), mas NOME não identificado.")
        registrar_log(f"DEBUG OCR: Falha Nome (CPF={cpf}). Texto parcial: {texto[:100]}")
    else:
        # Sucesso parcial (debug)
        pass # print(f"   (Leitura OK: {nome} | {cpf})")

    # ---------------- DATA ASO (mais robusta) ----------------
    data_aso = "Desconhecida"

    # Tenta pegar datas perto de palavras-chave
    padrao_data_contexto = re.search(
        r"(DATA\s*(do\s*ASO|exame)?[:\s\-]*)([0-3]?\d/[0-1]?\d/\d{4})",
        texto_compacto,
        re.IGNORECASE
    )
    if padrao_data_contexto:
        data_aso = padrao_data_contexto.group(3)
    else:
        # fallback: primeira data que aparecer no texto
        data_match = re.search(r"([0-3]?\d/[0-1]?\d/\d{4})", texto)
        if data_match:
            data_aso = data_match.group(1)

    # ---------------- FUNÇÃO / CARGO / GHE (mais precisa) ----------------
    funcao_cargo = "Desconhecida"

    # 1) Função com OCR bugado: Fungao, Funcéo, Funç5o, etc.
    padrao_funcao = re.search(
        r"Fun[cç5g][aãaõo0eéê]{1,3}o[:/\s\-]*([A-ZÀ-Ý0-9][A-ZÀ-Ý0-9 \-\._]{3,150})",
        texto_compacto,
        re.IGNORECASE
    )

    # 3) GHE:
    padrao_ghe = re.search(
        r"GHE[:/\s\-]*([0-9]{1,3}\s*\-\s*[A-ZÀ-Ý0-9 \-]{3,150})",
        texto_compacto,
        re.IGNORECASE
    )


    # 2) Cargo:
    padrao_cargo = re.search(
        r"Cargo[:/\s\-]*([A-ZÀ-Ý0-9][A-ZÀ-Ý0-9 \-\._]{3,150})",
        texto_compacto,
        re.IGNORECASE
    )

    # 2) Setor:
    padrao_setor = re.search(
        r"Setor[:/\s\-]*([A-ZÀ-Ý0-9][A-ZÀ-Ý0-9 \-\._]{3,150})",
        texto_compacto,
        re.IGNORECASE
    )

    
    if padrao_funcao:
        funcao_cargo = padrao_funcao.group(1).strip()
    elif padrao_cargo:
        funcao_cargo = padrao_cargo.group(1).strip()
    elif padrao_ghe:
        funcao_cargo = "GHE " + padrao_ghe.group(1).strip()
    elif padrao_setor:
        funcao_cargo = "GHE " + padrao_setor.group(1).strip()

    # alguns textos claramente não são função, filtrar
    lixo_patterns = [
        r"QUE EXERCE OU IRAEXERCER",
        r"QUE EXERCE OU IRÁ EXERCER",
        r"PULMONAR COMPLETA",
        r"EXAME",
    ]
    for lp in lixo_patterns:
        if re.search(lp, funcao_cargo, re.IGNORECASE):
            funcao_cargo = ""

    # Remover "RG" no final
    funcao_cargo = re.sub(r"\bRG\b", "", funcao_cargo, flags=re.IGNORECASE).strip()

    # Limpeza final
    funcao_cargo = re.sub(r"[^A-Za-zÀ-ÿ0-9\-\s\./]", "", funcao_cargo).strip()
    funcao_cargo = re.sub(r"\s{2,}", " ", funcao_cargo)

    if not funcao_cargo:
        funcao_cargo = "Desconhecida"

    return nome, cpf, data_aso, funcao_cargo

# ====================================================================
# SALVA PDFs SEPARADOS E GERA O TXT POR ANEXO
# ====================================================================
def salvar_paginas_individualmente(pdf_path, pasta_destino, numero_obra, lista_novos_arquivos=None):
    try:
        imagens = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
    except Exception as e:
        registrar_log(f"Erro ao converter PDF '{pdf_path}': {e}")
        return

    # >>>>> NOVO: TXT exclusivo para cada anexo
    nome_txt = f"OCR_{os.path.basename(pdf_path).replace('.pdf', '')}.txt"
    txt_path = os.path.join(pasta_destino, nome_txt)

    for i, img in enumerate(imagens, start=1):
        try:
            nome, cpf, dataaso, funcao_cargo = extrair_dados_completos(img)

            # OCR bruto da página para identificar ASO
            try:
                texto_ocr = pytesseract.image_to_string(img, lang="por+eng")
            except:
                texto_ocr = ""

            is_aso = eh_aso(texto_ocr)


            nome_limpo = re.sub(r"[^\w\s\-]", "", nome).strip()
            if not nome_limpo:
                nome_limpo = "FuncionarioDesconhecido"

            nome_final = f"{nome_limpo} - {cpf}.pdf"
            caminho_final = os.path.join(pasta_destino, nome_final)
            
            arquivo_salvo = False

            if not os.path.exists(caminho_final):
                img.save(caminho_final, "PDF", resolution=300.0)
                registrar_log(f"Arquivo salvo: {caminho_final}")
                arquivo_salvo = True
            else:
                # Alteração solicitada: NÃO sobrescrever se já existe.
                # Mas marcamos como True para garantir que o RPA verifique se este arquivo (já existente) foi processado.
                registrar_log(f"Arquivo já existe na pasta (não sobrescrito): {caminho_final}")
                arquivo_salvo = True
            
            if arquivo_salvo and lista_novos_arquivos is not None:
                lista_novos_arquivos.append(caminho_final)


            # >>>>> ESCREVE NO TXT DO ANEXO
            try:
                with open(txt_path, "a", encoding="utf-8") as txt:
                    txt.write("\n======================================\n")
                    txt.write(f"Obra: {numero_obra}\n")
                    txt.write(f"Arquivo gerado: {nome_final}\n")
                    txt.write(f"Nome: {nome}\n")
                    txt.write(f"CPF: {cpf}\n")
                    txt.write(f"Data ASO: {dataaso}\n")
                    txt.write(f"Função/Cargo: {funcao_cargo}\n")
                    txt.write("======================================\n")
            except Exception as e:
                registrar_log(f"Erro ao escrever TXT: {e}")

        except Exception as e:
            registrar_log(f"Erro na página {i} do PDF '{pdf_path}': {e}")


# ====================================================================
# CAPTA EMAIL E PROCESSA ANEXOS
# ====================================================================
def captar_emails(limit=200):
    registrar_log("Iniciando leitura do Outlook...")
    anexos_processados = set()



    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        conta_destino = None

        # Tentativa 1: localizar pela conta do Outlook
        for acc in outlook.Accounts:
            try:
                if acc.DisplayName.lower() == EMAIL_DESEJADO.lower():
                    conta_destino = acc
                    break
            except:
                pass

        # Tentativa 2: abrir mailbox direto pelo email
        if not conta_destino:
            try:
                conta_destino = outlook.Folders(EMAIL_DESEJADO)
            except:
                conta_destino = None

        # Tentativa 3: abrir mailbox pelo nome visível
        if not conta_destino and MAILBOX_NAME:
            try:
                conta_destino = outlook.Folders(MAILBOX_NAME)
            except:
                conta_destino = None

        # Se falhar tudo
        if not conta_destino:
            registrar_log(f"Mailbox não encontrada: {EMAIL_DESEJADO} / {MAILBOX_NAME}")
            return


        registrar_log(f"Usando mailbox: {getattr(conta_destino, 'Name', 'Desconhecida')}")

    except Exception as e:
        registrar_log(f"Erro ao conectar no Outlook: {e}")
        return

    try:
        # Caixa de entrada da conta selecionada
        try:
            inbox = conta_destino.Folders("Caixa de Entrada")
        except:
            inbox = conta_destino.Folders("Inbox")

        mensagens = inbox.Items
        mensagens.Sort("ReceivedTime", True)

    except Exception as e:
        registrar_log(f"Erro ao acessar caixa de entrada: {e}")
        return



    processados = 0
    encontrados_hoje = 0

    inicio_hoje = datetime.now().replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    inicio_amanha = inicio_hoje + timedelta(days=1)
    # Padrão: busca apenas emails de hoje (pode ser configurado via ASO_DAYS_BACK)
    days_back_env = os.getenv("ASO_DAYS_BACK")
    if days_back_env is None or days_back_env == "":
        days_back = 0  # Padrão: apenas hoje
    else:
        try:
            days_back = int(days_back_env)
            if days_back < 0:
                days_back = 0  # Se for negativo, usa 0 (apenas hoje)
        except ValueError:
            days_back = 0  # Se não for número válido, usa 0 (apenas hoje)
    
    inicio_janela = inicio_hoje - timedelta(days=days_back)
    
    if days_back == 0:
        registrar_log(f"Janela de busca: apenas hoje ({inicio_hoje.date()})")
    else:
        registrar_log(f"Janela de busca: últimos {days_back} dias (de {inicio_janela.date()} até {inicio_amanha.date()})")
        if days_back_env:
            registrar_log(f"  (Configurado via ASO_DAYS_BACK={days_back_env})")
    
    # ESTATÍSTICAS GERAIS ACUMULADAS
    stats_gerais = {
        'total': 0,
        'sucessos': [],
        'erros': [],
        'tempo_total': '' # será calculado no final aproximado
    }
    
    start_time_total = datetime.now()

    # Collect indices to iterate in reverse, to avoid issues with deleting items if that were ever implemented
    # For now, it just ensures consistent iteration order if new items arrive during processing
    indices = list(range(1, min(limit, mensagens.Count) + 1))
    
    registrar_log(f"Total de mensagens na caixa de entrada: {mensagens.Count}")
    
    for i in reversed(indices): # Changed from `for i in range(1, ...)` to `for i in reversed(indices)`
        try:
            msg = mensagens.Item(i)

            # Apenas emails
            if getattr(msg, "Class", None) != 43:
                continue

            recebido = msg.ReceivedTime.replace(tzinfo=None)
            assunto = msg.Subject or ""

            if not (inicio_janela <= recebido < inicio_amanha):
                continue

            if recebido.date() == inicio_hoje.date():
                encontrados_hoje += 1

            # Padrão mais flexível: aceita variações no formato
            m = re.search(
                r"ASO\s+ADMISSIONAL\s*-\s*([A-Za-z0-9]+)\s*-\s*([0-3]?\d/[0-1]?\d/\d{2,4})(?:\s*-\s*.*)?",
                assunto,
                re.IGNORECASE
            )

            if not m:
                continue

            numero_obra = m.group(1)
            registrar_log(f"✓ Email compatível encontrado - Obra: {numero_obra} | Assunto: {assunto[:60]}...")

            anexos_pdf = []
            total_attachments = msg.Attachments.Count
            for ai in range(1, total_attachments + 1):
                att = msg.Attachments.Item(ai)
                if att.FileName.lower().endswith(".pdf"):
                    anexos_pdf.append(att)

            if not anexos_pdf:
                registrar_log(f"  ⚠ Nenhum anexo PDF encontrado neste email")
                continue

            pasta_obra = os.path.join(PASTA_BASE, f"Obra_{numero_obra}")
            os.makedirs(pasta_obra, exist_ok=True)

            data_atual = datetime.now().strftime("%Y-%m-%d")
            pasta_data = os.path.join(pasta_obra, data_atual)
            os.makedirs(pasta_data, exist_ok=True)

            # LISTA DE ARQUIVOS GERADOS NESTA EXECUÇÃO PARA O RPA
            arquivos_para_rpa = []

            registrar_log(f"Pasta destino: {pasta_data}")

            for idx, anexo in enumerate(anexos_pdf, start=1):
                try:
                    temp_pdf = os.path.join(pasta_data, f"temp_{idx}.pdf")
                    anexo.SaveAsFile(temp_pdf)

                    hash_atual = calcular_hash_arquivo(temp_pdf)
                    if hash_atual in anexos_processados:
                        os.remove(temp_pdf)
                        continue

                    anexos_processados.add(hash_atual)

                    # Passamos a lista para coletar os novos arquivos
                    salvar_paginas_individualmente(temp_pdf, pasta_data, numero_obra, lista_novos_arquivos=arquivos_para_rpa)

                    os.remove(temp_pdf)

                except Exception as e:
                    registrar_log(f"Erro ao processar anexo {idx}: {e}")
                    stats_gerais['erros'].append({'arquivo': f"AnexoEmail_{idx}", 'erro': f"Erro extração PDF: {e}"})
            
            # ==================================================
            # CHAMAR RPA YUBE PARA A PASTA GERADA (APENAS NOVOS)
            # ==================================================
            if arquivos_para_rpa:
                try:
                    registrar_log(f"Iniciando RPA Yube para {len(arquivos_para_rpa)} arquivos novos...")
                    # Passamos a lista explícita para evitar processar lixo antigo
                    # E capturamos as estatísticas de retorno
                    stats_rpa = run_from_main(pasta_data, files_to_process=arquivos_para_rpa)
                    
                    if stats_rpa:
                        # ACUMULA RESULTADOS
                        stats_gerais['total'] += stats_rpa.get('total', 0)
                        stats_gerais['sucessos'].extend(stats_rpa.get('sucessos', []))
                        stats_gerais['erros'].extend(stats_rpa.get('erros', []))

                except Exception as e:
                    registrar_log(f"Erro ao executar RPA Yube: {e}")
                    stats_gerais['erros'].append({'arquivo': 'RPA_CRASH', 'erro': str(e)})
            else:
                registrar_log("Nenhum arquivo novo para processar no RPA.")

            processados += 1

        except Exception as e:
            registrar_log(f"Erro inesperado: {e}")
            try:
                erro_id = f"erro_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                pasta_erro = os.path.join(PASTA_ERROS, erro_id)
                os.makedirs(pasta_erro, exist_ok=True)
                with open(os.path.join(pasta_erro, "erro.txt"), "w", encoding="utf-8") as f:
                    f.write(traceback.format_exc())
            except:
                pass
            continue

    # ==================================================
    # ENVIO DO RESUMO CONSOLIDADO
    # ==================================================
    elapsed_total = datetime.now() - start_time_total
    stats_gerais['tempo_total'] = str(timedelta(seconds=int(elapsed_total.total_seconds())))

    # Só envia email se houver algo processado (sucesso ou erro)
    if stats_gerais['total'] > 0 or stats_gerais['erros']:
        registrar_log("Gerando relatório e enviando email...")
        
        # 1. Salvar Relatório JSON
        reporter.save_report(stats_gerais)
        
        # 2. Enviar Email
        try:
           enviar_resumo_email(TARGET_ACCOUNT, stats_gerais)
        except Exception as e:
           registrar_log(f"Erro ao enviar email final: {e}")
    else:
        registrar_log("Nada processado, email de resumo não enviado.")

    registrar_log(f"Mensagens verificadas hoje: {encontrados_hoje}; mensagens processadas: {processados}")



# ====================================================================
# MAIN
# ====================================================================
if __name__ == "__main__":
    registrar_log("===== INÍCIO DO PROCESSAMENTO DIÁRIO DE ASO =====")
    try:
        captar_emails(limit=500)
    except Exception as e:
        registrar_log(f"Erro fatal: {e}")
        try:
            with open(os.path.join(PASTA_ERROS, "erro_fatal.txt"), "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except:
            pass
    registrar_log("===== FIM DO PROCESSAMENTO =====")
