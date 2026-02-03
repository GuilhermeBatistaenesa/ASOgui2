import win32com.client as win32
import pywintypes
import pythoncom
import os
import re
import shutil
import uuid
import json
from dotenv import load_dotenv
import sys
import html as html_lib
import urllib.request
import urllib.parse
import http.cookiejar

# Carrega variaveis de ambiente do arquivo .env (se existir)
_env_candidates = []
if getattr(sys, "frozen", False):
    _env_candidates.append(os.path.join(os.path.dirname(sys.executable), ".env"))
_env_candidates.append(os.path.join(os.getcwd(), ".env"))
_env_candidates.append(os.path.join(os.path.dirname(__file__), ".env"))
for _env_path in _env_candidates:
    if _env_path and os.path.exists(_env_path):
        load_dotenv(_env_path)
        break
else:
    load_dotenv()
import pytesseract
from pdf2image import convert_from_path
from datetime import datetime, timedelta
import traceback
import hashlib
import time
from PIL import ImageOps, ImageFilter
from rpa_yube import run_from_main

# Módulos customizados
from custom_logger import RpaLogger
from reporting import ReportGenerator
from notification import enviar_resumo_email
from outcomes import (
    SUCCESS,
    ERROR,
    SKIPPED_DUPLICATE,
    SKIPPED_DRAFT,
    SKIPPED_NON_ASO,
)
from utils_masking import mask_cpf, mask_cpf_in_text, mask_pii_in_obj
from idempotency import should_skip_duplicate

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
def find_poppler():
    candidates = []
    env_path = os.getenv("POPPLER_PATH")
    if env_path:
        candidates.append(env_path)

    here = os.path.dirname(__file__)
    candidates.extend([
        os.path.join(os.getcwd(), "vendor", "poppler", "bin"),
        os.path.join(here, "vendor", "poppler", "bin"),
        os.path.join(here, "Release-24.08.0-0", "poppler-24.08.0", "Library", "bin"),
        r"C:\Program Files\poppler\bin",
        r"C:\Program Files\poppler-24.08.0\Library\bin",
        r"C:\Program Files (x86)\poppler\bin",
    ])

    for p in candidates:
        if not p:
            continue
        candidate = p
        if os.path.isfile(candidate) and candidate.lower().endswith("pdftoppm.exe"):
            candidate = os.path.dirname(candidate)
        if os.path.isdir(candidate):
            exe_path = os.path.join(candidate, "pdftoppm.exe")
            if os.path.exists(exe_path):
                return candidate
    return None

POPPLER_PATH = find_poppler()
if POPPLER_PATH:
    print(f"DEBUG: Configured Poppler Path: '{POPPLER_PATH}'")
else:
    print("AVISO: Poppler (pdftoppm.exe) nao encontrado. Configure POPPLER_PATH no .env")

PASTA_BASE = os.getenv("PROCESSO_ASO_BASE", r"P:\ProcessoASO")
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
GDRIVE_NAME_FILTER = os.getenv("ASO_GDRIVE_NAME_FILTER", "asos enesa").strip().lower()
GDRIVE_TIMEOUT_SEC = int(os.getenv("ASO_GDRIVE_TIMEOUT_SEC", "60"))


# ====================================================================
# FUNCAO - EXTRAI DADOS COMPLETOS (OCR)
# ====================================================================
def _score_ocr(texto):
    if not texto:
        return 0
    score = 0
    t = texto.upper()
    if "CPF" in t:
        score += 3
    if "ASO" in t or "SAUDE OCUPACIONAL" in t or "SA?DE OCUPACIONAL" in t:
        score += 2
    if re.search(r"\d{3}\.\d{3}\.\d{3}-\d{2}", texto):
        score += 3
    if re.search(r"\d{11}", texto):
        score += 2
    if "FUNCION" in t:
        score += 1
    return score


def _preprocess_img(img, scale=1.5):
    try:
        gray = ImageOps.grayscale(img)
        if scale and scale != 1.0:
            w, h = gray.size
            gray = gray.resize((int(w * scale), int(h * scale)))
        gray = ImageOps.autocontrast(gray)
        gray = gray.filter(ImageFilter.MedianFilter(size=3))
        gray = gray.filter(ImageFilter.SHARPEN)
        bw = gray.point(lambda x: 0 if x < 180 else 255, mode="1")
        return bw
    except Exception:
        return img


def ocr_with_fallback(img, force_full=False):
    # Primeira tentativa rapida
    try:
        base = pytesseract.image_to_string(img, lang="por+eng", config="--oem 3 --psm 6")
    except Exception:
        base = ""

    # Se a leitura base ja tem sinais bons, retorna sem fallback
    if not force_full and _score_ocr(base) >= 5:
        return base

    # Fallbacks so quando falhar
    configs = [
        "--oem 3 --psm 11",
        "--oem 3 --psm 4",
    ]
    texts = [base]
    for cfg in configs:
        try:
            t = pytesseract.image_to_string(img, lang="por+eng", config=cfg)
            texts.append(t)
        except Exception:
            texts.append("")
    pre = _preprocess_img(img)
    for cfg in configs:
        try:
            t = pytesseract.image_to_string(pre, lang="por+eng", config=cfg)
            texts.append(t)
        except Exception:
            texts.append("")

    best = ""
    best_score = -1
    for t in texts:
        s = _score_ocr(t)
        if s > best_score:
            best_score = s
            best = t
    return best

# ====================================================================
def registrar_log(msg, context=None):
    # Mantém compatibilidade com chamadas existentes
    logger.info(msg, extra=context)

def salvar_diagnostico_resumo(stats, manifest_path=None, report_paths=None, extra=None):
    try:
        os.makedirs(PASTA_LOGS, exist_ok=True)
        path = os.path.join(PASTA_LOGS, "diagnostico_ultima_execucao.txt")
        lines = []
        lines.append(f"execution_id: {stats.get('execution_id')}")
        lines.append(f"started_at: {stats.get('started_at')}")
        lines.append(f"finished_at: {stats.get('finished_at')}")
        lines.append(f"run_status: {stats.get('run_status')}")
        lines.append(f"total_detected: {stats.get('total_detected')}")
        lines.append(f"total_processed: {stats.get('total_processed')}")
        lines.append(f"success: {stats.get('success')}")
        lines.append(f"error: {stats.get('error')}")
        lines.append(f"skipped_duplicate: {stats.get('skipped_duplicate')}")
        lines.append(f"skipped_draft: {stats.get('skipped_draft')}")
        lines.append(f"skipped_non_aso: {stats.get('skipped_non_aso')}")
        if manifest_path:
            lines.append(f"manifest: {manifest_path}")
        if report_paths:
            lines.append(f"report_json: {report_paths.get('json')}")
            lines.append(f"report_md: {report_paths.get('md')}")
        if extra:
            for k, v in extra.items():
                lines.append(f"{k}: {v}")
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    except Exception:
        pass

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
def extrair_dados_completos(img, texto_ocr=None, _retry=False):
    """
    Retorna: nome, cpf, data_aso, funcao_cargo, texto_ocr
    """
    if texto_ocr is None:
        try:
            texto_ocr = ocr_with_fallback(img)
        except Exception as e:
            texto_ocr = ""
            registrar_log(f"Erro no pytesseract: {e}")

    texto = texto_ocr or ""

    # Checagem de Rascunho
    if "RASCUNHO" in texto.upper() and texto.upper().count("RASCUNHO") > 3:
         print("INFO: Arquivo identificado como RASCUNHO. Ignorando.")
         return "RASCUNHO", "Ignorar", "", "", texto_ocr

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
            
    # Fallback: nome imediatamente antes do CPF
    if nome == "Desconhecido":
        bloco = re.search(
            r"([A-ZÀ-Ý][A-ZÀ-Ý \-]{5,150})\s+(?:CPF|C\.P\.F)",
            texto
        )
        if bloco:
            nome = bloco.group(1).strip()

    # Se nome ainda desconhecido, tentar um OCR mais agressivo uma unica vez
    if nome == "Desconhecido" and not _retry:
        try:
            texto2 = ocr_with_fallback(img, force_full=True)
        except Exception:
            texto2 = ""
        if texto2 and texto2 != texto_ocr:
            return extrair_dados_completos(img, texto_ocr=texto2, _retry=True)

    # Debug para casos falhos (apos fallback)
    if nome == "Desconhecido":
        registrar_log(f"DEBUG OCR FALHO (CPF={mask_cpf(cpf)}): Texto parcial: {texto[:200].replace(chr(10), ' ')}")

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
        print("ALERTA DE LEITURA: Nao foi possivel ler o CPF neste arquivo.")
        registrar_log(f"DEBUG OCR: Falha CPF", context={"partial_text": texto[:100], "status": "CPF_MISSING"})
    elif nome == "Desconhecido":
        print(f"ALERTA DE LEITURA: CPF encontrado ({mask_cpf(cpf)}), mas NOME nao identificado.")
        registrar_log(f"DEBUG OCR: Falha Nome", context={"cpf": cpf, "partial_text": texto[:100], "status": "NAME_MISSING"})
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
    # 1) Função com OCR bugado: Fungao, Funcéo, Funç5o, etc.
    # Stop capturing if we hit another label like Setor, Cargo, GHE, Riscos, CPF, Data
    padrao_funcao = re.search(
        r"Fun[cç5g][aãaõo0eéê]{1,3}o[:/\s\-]*([A-ZÀ-Ý0-9][A-ZÀ-Ý0-9 \-\._]{3,150}?)(?=\s+(?:Setor|Cargo|GHE|Riscos|CPF|Data)|$)",
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

    return nome, cpf, data_aso, funcao_cargo, texto_ocr

# ====================================================================
# SALVA PDFs SEPARADOS E GERA O TXT POR ANEXO
# ====================================================================
def salvar_paginas_individualmente(pdf_path, pasta_destino, numero_obra, lista_novos_arquivos=None, stats=None, manifest_items=None):
    try:
        imagens = convert_from_path(pdf_path, dpi=300, poppler_path=POPPLER_PATH)
    except Exception as e:
        registrar_log(f"Erro ao converter PDF '{pdf_path}': {e}")
        if stats is not None:
            stats["error"] += 1
            stats["erros"].append({"arquivo": os.path.basename(pdf_path), "erro": f"Erro conversao PDF: {e}"})
        return

    nome_txt = f"OCR_{os.path.basename(pdf_path).replace('.pdf', '')}.txt"
    txt_path = os.path.join(pasta_destino, nome_txt)

    for i, img in enumerate(imagens, start=1):
        try:
            try:
                texto_ocr = ocr_with_fallback(img)
            except Exception as e:
                texto_ocr = ""
                registrar_log(f"Erro no pytesseract: {e}")

            nome, cpf, dataaso, funcao_cargo, texto_ocr = extrair_dados_completos(img, texto_ocr=texto_ocr)
            is_aso = eh_aso(texto_ocr)

            if stats is not None:
                stats["total_detected"] += 1

            outcome = None
            outcome_msg = None
            if nome == "Desconhecido" or cpf == "CPF_Desconhecido":
                outcome = ERROR
                outcome_msg = "Falha OCR (nome/cpf)"
                if stats is not None:
                    stats["error"] += 1
                    stats["erros"].append({"arquivo": f"{os.path.basename(pdf_path)}#pg{i}", "erro": outcome_msg})
                    stats["ocr_failures"].append({
                        "arquivo": f"{os.path.basename(pdf_path)}#pg{i}",
                        "cpf": mask_cpf(cpf),
                        "nome": nome
                    })
            if outcome is None and (nome == "RASCUNHO" or cpf == "Ignorar"):
                outcome = SKIPPED_DRAFT
                outcome_msg = "Rascunho detectado"
                if stats is not None:
                    stats["skipped_draft"] += 1
                    stats["skipped_items"].append(f"{outcome}: {mask_cpf_in_text(os.path.basename(pdf_path))}#pg{i}")
            elif outcome is None and not is_aso:
                outcome = SKIPPED_NON_ASO
                outcome_msg = "Nao identificado como ASO"
                if stats is not None:
                    stats["skipped_non_aso"] += 1
                    stats["skipped_items"].append(f"{outcome}: {mask_cpf_in_text(os.path.basename(pdf_path))}#pg{i}")

            if outcome:
                if manifest_items is not None:
                    manifest_items.append({
                        "file_display": mask_cpf_in_text(os.path.basename(pdf_path)),
                        "page": i,
                        "cpf_masked": mask_cpf(cpf),
                        "outcome": outcome,
                        "message": outcome_msg,
                    })

                try:
                    with open(txt_path, "a", encoding="utf-8") as txt:
                        txt.write("\n======================================\n")
                        txt.write(f"Obra: {numero_obra}\n")
                        txt.write(f"Outcome: {outcome}\n")
                        txt.write(f"Arquivo origem: {os.path.basename(pdf_path)}\n")
                        txt.write(f"Nome: {nome}\n")
                        txt.write(f"CPF: {mask_cpf(cpf)}\n")
                        txt.write(f"Data ASO: {dataaso}\n")
                        txt.write(f"Funcao/Cargo: {funcao_cargo}\n")
                        txt.write("======================================\n")
                except Exception as e:
                    registrar_log(f"Erro ao escrever TXT: {e}")
                continue

            nome_limpo = re.sub(r"[^\w\s\-]", "", nome).strip()
            if not nome_limpo:
                nome_limpo = "FuncionarioDesconhecido"

            nome_final = f"{nome_limpo} - {cpf}.pdf"
            caminho_final = os.path.join(pasta_destino, nome_final)

            if should_skip_duplicate(caminho_final):
                registrar_log(f"Arquivo ja existe na pasta (nao sobrescrito): {caminho_final}")
                if stats is not None:
                    stats["skipped_duplicate"] += 1
                    stats["skipped_items"].append(f"{SKIPPED_DUPLICATE}: {mask_cpf_in_text(nome_final)}")
                if manifest_items is not None:
                    manifest_items.append({
                        "file_display": mask_cpf_in_text(nome_final),
                        "cpf_masked": mask_cpf(cpf),
                        "outcome": SKIPPED_DUPLICATE,
                        "message": "Arquivo ja existente",
                    })
                continue

            img.save(caminho_final, "PDF", resolution=300.0)
            registrar_log(f"Arquivo salvo: {caminho_final}")

            if lista_novos_arquivos is not None:
                lista_novos_arquivos.append(caminho_final)

            try:
                with open(txt_path, "a", encoding="utf-8") as txt:
                    txt.write("\n======================================\n")
                    txt.write(f"Obra: {numero_obra}\n")
                    txt.write(f"Arquivo gerado: {nome_final}\n")
                    txt.write(f"Nome: {nome}\n")
                    txt.write(f"CPF: {mask_cpf(cpf)}\n")
                    txt.write(f"Data ASO: {dataaso}\n")
                    txt.write(f"Funcao/Cargo: {funcao_cargo}\n")
                    txt.write("======================================\n")
            except Exception as e:
                registrar_log(f"Erro ao escrever TXT: {e}")

        except Exception as e:
            registrar_log(f"Erro na pagina {i} do PDF '{pdf_path}': {e}")
            if stats is not None:
                stats["error"] += 1
                stats["erros"].append({"arquivo": os.path.basename(pdf_path), "erro": f"Erro pagina {i}: {e}"})


# ====================================================================
# CAPTA EMAIL E PROCESSA ANEXOS
# ====================================================================


def salvar_manifest(manifest, report_dir, filepath=None, execution_id=None):
    os.makedirs(report_dir, exist_ok=True)
    if not filepath:
        if execution_id:
            filename = f"manifest_{execution_id}.json"
        else:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"manifest_{ts}.json"
        filepath = os.path.join(report_dir, filename)
    safe_manifest = mask_pii_in_obj(manifest)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(safe_manifest, f, indent=2, ensure_ascii=False)
    return filepath


def _extract_gdrive_file_ids(text):
    if not text:
        return []
    text = html_lib.unescape(text)
    ids = set()
    patterns = [
        r"https?://drive\.google\.com/file/d/([a-zA-Z0-9_-]{10,})",
        r"https?://drive\.google\.com/open\?id=([a-zA-Z0-9_-]{10,})",
        r"https?://drive\.google\.com/uc\?[^ \t\r\n\"'>]+?id=([a-zA-Z0-9_-]{10,})",
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            ids.add(m.group(1))
    return list(ids)


def _parse_filename_from_cd(content_disposition):
    if not content_disposition:
        return None
    m = re.search(r"filename\*=UTF-8''([^;]+)", content_disposition, re.IGNORECASE)
    if m:
        return urllib.parse.unquote(m.group(1)).strip()
    m = re.search(r'filename="?([^";]+)"?', content_disposition, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None


def _parse_filename_from_html(html_text):
    if not html_text:
        return None
    m = re.search(r"<title>\s*(.*?)\s*- Google Drive\s*</title>", html_text, re.IGNORECASE)
    if m:
        return html_lib.unescape(m.group(1)).strip()
    m = re.search(r'class="uc-name-size"[^>]*>\s*([^<]+)\s*<', html_text, re.IGNORECASE)
    if m:
        return html_lib.unescape(m.group(1)).strip()
    return None


def _parse_confirm_token(html_text):
    if not html_text:
        return None
    m = re.search(r"confirm=([0-9A-Za-z_]+)", html_text)
    if m:
        return m.group(1)
    return None


def _safe_filename(filename):
    if not filename:
        return None
    filename = os.path.basename(filename)
    filename = re.sub(r"[\\/:*?\"<>|]+", "_", filename)
    return filename.strip()


def _unique_path(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    for i in range(1, 1000):
        candidate = f"{base}_{i}{ext}"
        if not os.path.exists(candidate):
            return candidate
    return f"{base}_{int(time.time())}{ext}"


def _gdrive_name_matches(filename):
    if not GDRIVE_NAME_FILTER:
        return True
    return GDRIVE_NAME_FILTER in (filename or "").lower()


def _stream_download(resp, dest_dir, filename):
    safe_name = _safe_filename(filename) or f"gdrive_{int(time.time())}.bin"
    full_path = _unique_path(os.path.join(dest_dir, safe_name))
    with open(full_path, "wb") as f:
        while True:
            chunk = resp.read(1024 * 1024)
            if not chunk:
                break
            f.write(chunk)
    try:
        resp.close()
    except Exception:
        pass
    return full_path


def download_gdrive_file(file_id, dest_dir):
    cookiejar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookiejar))

    def _open(url):
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "*/*",
            },
        )
        return opener.open(req, timeout=GDRIVE_TIMEOUT_SEC)

    base_url = "https://drive.google.com/uc?export=download"
    url = f"{base_url}&id={urllib.parse.quote(file_id)}"

    resp = _open(url)
    cd = resp.headers.get("Content-Disposition")
    filename = _parse_filename_from_cd(cd)
    if filename:
        safe_name = _safe_filename(filename)
        if not _gdrive_name_matches(safe_name):
            resp.close()
            return None
        if not safe_name.lower().endswith(".pdf"):
            ctype = (resp.headers.get("Content-Type") or "").lower()
            if "pdf" not in ctype:
                resp.close()
                return None
        return _stream_download(resp, dest_dir, safe_name)

    html_text = resp.read(1024 * 1024).decode("utf-8", errors="ignore")
    resp.close()
    filename = _parse_filename_from_html(html_text)
    confirm = _parse_confirm_token(html_text)

    if not confirm:
        return None

    url2 = f"{base_url}&confirm={urllib.parse.quote(confirm)}&id={urllib.parse.quote(file_id)}"
    resp2 = _open(url2)
    cd2 = resp2.headers.get("Content-Disposition")
    filename2 = _parse_filename_from_cd(cd2) or filename or f"gdrive_{file_id}.pdf"
    safe_name2 = _safe_filename(filename2)
    if not _gdrive_name_matches(safe_name2):
        resp2.close()
        return None
    if not safe_name2.lower().endswith(".pdf"):
        ctype2 = (resp2.headers.get("Content-Type") or "").lower()
        if "pdf" not in ctype2:
            resp2.close()
            return None
    return _stream_download(resp2, dest_dir, safe_name2)



RPC_E_CALL_REJECTED = -2147418111
_OUTLOOK_PREV_FILTER = None
_OUTLOOK_COM_INIT = False


try:
    _IID_IOleMessageFilter = pythoncom.IID_IOleMessageFilter
except AttributeError:
    try:
        _IID_IOleMessageFilter = pywintypes.IID("{00000016-0000-0000-C000-000000000046}")
    except Exception:
        _IID_IOleMessageFilter = None


class MessageFilter:
    _com_interfaces_ = [_IID_IOleMessageFilter]
    _public_methods_ = ["HandleInComingCall", "RetryRejectedCall", "MessagePending"]

    def __init__(self, max_wait_sec=60):
        self._start = time.time()
        self._max_wait = max_wait_sec
        self._delays = [250, 500, 1000, 2000]
        self._idx = 0

    def HandleInComingCall(self, callType, taskCaller, tickCount, interfaceInfo):
        return 0  # SERVERCALL_ISHANDLED

    def RetryRejectedCall(self, taskCaller, rejectType, tickCount):
        try:
            retry_later = pythoncom.SERVERCALL_RETRYLATER
        except Exception:
            retry_later = 2
        elapsed = time.time() - self._start
        if elapsed >= self._max_wait:
            return -1  # CANCELCALL
        if rejectType == retry_later:
            delay = self._delays[self._idx % len(self._delays)]
            self._idx += 1
            return delay
        return -1

    def MessagePending(self, taskCaller, tickCount, pendingType):
        return 2  # PENDINGMSG_WAITDEFPROCESS


def _register_message_filter(timeout_sec):
    global _OUTLOOK_PREV_FILTER
    if not hasattr(pythoncom, "CoRegisterMessageFilter"):
        return False
    if _IID_IOleMessageFilter is None:
        return False
    filt = MessageFilter(max_wait_sec=timeout_sec)
    _OUTLOOK_PREV_FILTER = pythoncom.CoRegisterMessageFilter(filt)
    return True


def _unregister_message_filter():
    global _OUTLOOK_PREV_FILTER
    if not hasattr(pythoncom, "CoRegisterMessageFilter"):
        _OUTLOOK_PREV_FILTER = None
        return
    try:
        pythoncom.CoRegisterMessageFilter(_OUTLOOK_PREV_FILTER)
    finally:
        _OUTLOOK_PREV_FILTER = None


def cleanup_outlook_com():
    global _OUTLOOK_COM_INIT
    try:
        _unregister_message_filter()
    except Exception:
        pass
    if _OUTLOOK_COM_INIT:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        _OUTLOOK_COM_INIT = False


def get_outlook_namespace_robusto(timeout_sec=60):
    global _OUTLOOK_COM_INIT
    pythoncom.CoInitialize()
    _OUTLOOK_COM_INIT = True
    filter_enabled = _register_message_filter(timeout_sec)

    try:
        try:
            app = win32.GetActiveObject("Outlook.Application")
        except Exception:
            app = win32.DispatchEx("Outlook.Application")

        ns = app.GetNamespace("MAPI")
        try:
            ns.Logon("", "", False, False)
        except Exception:
            pass

        start = time.time()
        attempt = 0
        while True:
            try:
                _ = ns.Folders.Count
                registrar_log("Outlook disponivel.")
                return app, ns
            except pywintypes.com_error as e:
                hr = e.args[0] if e.args else None
                elapsed = int(time.time() - start)
                if hr == RPC_E_CALL_REJECTED:
                    attempt += 1
                    registrar_log(f"Aguardando Outlook ficar disponivel... tentativa {attempt} ({elapsed}s)")
                    if elapsed >= timeout_sec:
                        raise RuntimeError("Outlook ocupado ou com prompt aberto")
                    time.sleep(1)
                    continue
                raise
    except Exception:
        raise

def captar_emails(limit=200, execution_id=None, started_at=None, manifest=None):
    def _get_msg_datetime(msg):
        for attr in ("ReceivedTime", "SentOn", "CreationTime"):
            try:
                t = getattr(msg, attr)
                if isinstance(t, datetime):
                    if t.tzinfo:
                        t = t.astimezone().replace(tzinfo=None)
                    return t
            except Exception:
                continue
        return None

    def _get_shared_inbox(ns, smtp):
        try:
            recip = ns.CreateRecipient(smtp)
            recip.Resolve()
            if not getattr(recip, "Resolved", False):
                return None
            # 6 = olFolderInbox
            return ns.GetSharedDefaultFolder(recip, 6)
        except Exception:
            return None

    def _capta_core():
        registrar_log("Iniciando leitura do Outlook...")
        anexos_processados = set()

        try:
            app, outlook = get_outlook_namespace_robusto()

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

            # Tentativa 3: abrir mailbox pelo nome visivel
            if not conta_destino and MAILBOX_NAME:
                try:
                    conta_destino = outlook.Folders(MAILBOX_NAME)
                except:
                    conta_destino = None

            # Se falhar tudo
            if not conta_destino:
                registrar_log(f"Mailbox nao encontrada: {EMAIL_DESEJADO} / {MAILBOX_NAME}")
                return

            registrar_log(f"Usando mailbox: {getattr(conta_destino, 'Name', 'Desconhecida')}")

        except Exception as e:
            registrar_log(f"Erro ao conectar no Outlook: {e}")
            registrar_log("Verifique se o Outlook esta aberto e sem prompts. Sugestao: fechar e abrir manualmente, testar outlook /safe.")
            return

        try:
            # Caixa de entrada da conta selecionada
            try:
                inbox = conta_destino.Folders("Caixa de Entrada")
            except:
                inbox = conta_destino.Folders("Inbox")

            mensagens = inbox.Items
            try:
                mensagens.Sort("[ReceivedTime]", True)
            except Exception:
                mensagens.Sort("ReceivedTime", True)

        except Exception as e:
            registrar_log(f"Erro ao acessar caixa de entrada: {e}")
            # Fallback: tentar inbox compartilhado via recipient
            inbox = _get_shared_inbox(outlook, EMAIL_DESEJADO)
            if inbox is None:
                return
            try:
                mensagens = inbox.Items
                try:
                    mensagens.Sort("[ReceivedTime]", True)
                except Exception:
                    mensagens.Sort("ReceivedTime", True)
                registrar_log(f"Usando inbox compartilhado: {getattr(inbox, 'FolderPath', 'Desconhecido')}")
            except Exception as e2:
                registrar_log(f"Erro ao acessar inbox compartilhado: {e2}")
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
        
        # ESTATISTICAS GERAIS ACUMULADAS
        stats_gerais = {
            'execution_id': execution_id,
            'started_at': started_at.isoformat() if started_at else None,
            'total_detected': 0,
            'total_processed': 0,
            'success': 0,
            'error': 0,
            'skipped_duplicate': 0,
            'skipped_draft': 0,
            'skipped_non_aso': 0,
            'ocr_failures': [],
            'sucessos': [],
            'erros': [],
            'skipped_items': [],
            'tempo_total': ''
        }
        last_error = None
        start_time_total = datetime.now()
        
        # Collect indices to iterate in reverse, to avoid issues with deleting items if that were ever implemented
        # For now, it just ensures consistent iteration order if new items arrive during processing
        indices = list(range(1, min(limit, mensagens.Count) + 1))
        
        registrar_log(f"Total de mensagens na caixa de entrada: {mensagens.Count}")
        # Se a caixa atual nao tiver emails de hoje, tenta inbox compartilhado
        if EMAIL_DESEJADO:
            try:
                today_count = 0
                for di in range(1, min(50, mensagens.Count) + 1):
                    try:
                        dmsg = mensagens.Item(di)
                        drec = _get_msg_datetime(dmsg)
                        if drec and drec.date() == inicio_hoje.date():
                            today_count += 1
                    except Exception:
                        continue
                if today_count == 0:
                    shared_inbox = _get_shared_inbox(outlook, EMAIL_DESEJADO)
                    if shared_inbox and shared_inbox is not inbox:
                        shared_items = shared_inbox.Items
                        try:
                            shared_items.Sort("[ReceivedTime]", True)
                        except Exception:
                            shared_items.Sort("ReceivedTime", True)
                        inbox = shared_inbox
                        mensagens = shared_items
                        registrar_log(f"Usando inbox compartilhado: {getattr(shared_inbox, 'FolderPath', 'Desconhecido')}")
            except Exception:
                pass
        
        for i in reversed(indices): # Changed from `for i in range(1, ...)` to `for i in reversed(indices)`
            try:
                msg = mensagens.Item(i)
        
                # Apenas emails
                if getattr(msg, "Class", None) != 43:
                    continue
        
                recebido = _get_msg_datetime(msg)
                if not recebido:
                    continue
                assunto = msg.Subject or ""
        
                if not (inicio_janela <= recebido < inicio_amanha):
                    continue
        
                if recebido.date() == inicio_hoje.date():
                    encontrados_hoje += 1
        
                # Padrao mais flexivel: aceita prefixos (ENC/RE/FW) e pequenas variacoes
                m = re.search(
                    r"(?:ENC:|RE:|FWD:|FW:)?\s*ASO\s+ADMISSIONAL\s*-\s*([A-Za-z0-9]+)\s*-\s*([0-3]?\d/[0-1]?\d/\d{2,4})(?:\s*-\s*.*)?",
                    assunto,
                    re.IGNORECASE
                )
        
                if not m:
                    continue
        
                numero_obra = m.group(1)
                registrar_log(f"Email compativel encontrado - Obra: {numero_obra} | Assunto: {assunto[:60]}...")

                pasta_obra = os.path.join(PASTA_BASE, f"Obra_{numero_obra}")
                os.makedirs(pasta_obra, exist_ok=True)

                data_atual = datetime.now().strftime("%Y-%m-%d")
                pasta_data = os.path.join(pasta_obra, data_atual)
                os.makedirs(pasta_data, exist_ok=True)

                # LISTA DE ARQUIVOS GERADOS NESTA EXECUCAO PARA O RPA
                arquivos_para_rpa = []

                registrar_log(f"Pasta destino: {pasta_data}")
        
                anexos_pdf = []
                total_attachments = msg.Attachments.Count
                for ai in range(1, total_attachments + 1):
                    att = msg.Attachments.Item(ai)
                    if att.FileName.lower().endswith(".pdf"):
                        anexos_pdf.append(att)
        
                if anexos_pdf:
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
                            salvar_paginas_individualmente(temp_pdf, pasta_data, numero_obra, lista_novos_arquivos=arquivos_para_rpa, stats=stats_gerais, manifest_items=(manifest.get('items') if manifest else None))
                            
                            os.remove(temp_pdf)
                            
                        except Exception as e:
                            last_error = f"Erro ao processar anexo {idx}: {e}"
                            registrar_log(last_error)
                            stats_gerais['erros'].append({'arquivo': f"AnexoEmail_{idx}", 'erro': f"Erro extracao PDF: {e}"})
                else:
                    body_text = ""
                    try:
                        body_text = (msg.HTMLBody or "")
                    except Exception:
                        body_text = ""
                    try:
                        body_text = body_text + "\n" + (msg.Body or "")
                    except Exception:
                        pass
                    
                    gdrive_ids = _extract_gdrive_file_ids(body_text)
                    if not gdrive_ids:
                        registrar_log("  Aviso: Nenhum anexo PDF ou link Google Drive encontrado neste email")
                        continue
                    
                    registrar_log(f"  Nenhum anexo PDF. Links Google Drive encontrados: {len(gdrive_ids)}")
                    
                    baixados = []
                    for gid in gdrive_ids:
                        try:
                            caminho = download_gdrive_file(gid, pasta_data)
                            if not caminho:
                                registrar_log(f"  Link Google Drive ignorado pelo filtro de nome: {gid}")
                                continue
                            if not caminho.lower().endswith('.pdf'):
                                registrar_log(f"  Link Google Drive ignorado (nao PDF): {os.path.basename(caminho)}")
                                try:
                                    os.remove(caminho)
                                except Exception:
                                    pass
                                continue
                            registrar_log(f"  Baixado Google Drive: {os.path.basename(caminho)}")
                            baixados.append(caminho)
                        except Exception as e:
                            last_error = f"Erro ao baixar Google Drive ({gid}): {e}"
                            registrar_log(last_error)
                            stats_gerais['erros'].append({'arquivo': f"GoogleDrive_{gid}", 'erro': f"Erro download: {e}"})
                            stats_gerais['error'] += 1
                    
                    if not baixados:
                        registrar_log("  Nenhum arquivo valido baixado do Google Drive.")
                        continue
                    
                    for idx, temp_pdf in enumerate(baixados, start=1):
                        try:
                            hash_atual = calcular_hash_arquivo(temp_pdf)
                            if hash_atual in anexos_processados:
                                os.remove(temp_pdf)
                                continue
                            
                            anexos_processados.add(hash_atual)
                            
                            salvar_paginas_individualmente(temp_pdf, pasta_data, numero_obra, lista_novos_arquivos=arquivos_para_rpa, stats=stats_gerais, manifest_items=(manifest.get('items') if manifest else None))
                            
                            os.remove(temp_pdf)
                            
                        except Exception as e:
                            last_error = f"Erro ao processar Google Drive {idx}: {e}"
                            registrar_log(last_error)
                            stats_gerais['erros'].append({'arquivo': f"GoogleDrive_{idx}", 'erro': f"Erro extracao PDF: {e}"})
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
                            stats_gerais['sucessos'].extend([mask_cpf_in_text(s) for s in stats_rpa.get('sucessos', [])])
                            stats_gerais['erros'].extend([{"arquivo": mask_cpf_in_text(e.get('arquivo', 'Desconhecido')), "erro": e.get('erro', '')} for e in stats_rpa.get('erros', [])])
                            stats_gerais['success'] += len(stats_rpa.get('sucessos', []))
                            stats_gerais['error'] += len(stats_rpa.get('erros', []))
                            stats_gerais['total_processed'] = stats_gerais['success'] + stats_gerais['error']
                            if manifest and manifest.get('items') is not None:
                                for fname in stats_rpa.get('sucessos', []):
                                    manifest['items'].append({
                                        'file_display': mask_cpf_in_text(fname),
                                        'cpf_masked': mask_cpf(fname),
                                        'outcome': SUCCESS,
                                        'message': 'RPA sucesso'
                                    })
                                for err in stats_rpa.get('erros', []):
                                    manifest['items'].append({
                                        'file_display': mask_cpf_in_text(err.get('arquivo', 'Desconhecido')),
                                        'cpf_masked': mask_cpf(err.get('arquivo', '')),
                                        'outcome': ERROR,
                                        'message': err.get('erro', '')
                                    })
        
                    except Exception as e:
                        last_error = f"Erro ao executar RPA Yube: {e}"
                        registrar_log(last_error)
                        stats_gerais['erros'].append({'arquivo': 'RPA_CRASH', 'erro': str(e)})
                        stats_gerais['error'] += 1
                else:
                    registrar_log("Nenhum arquivo novo para processar no RPA.")
        
                processados += 1
        
            except Exception as e:
                last_error = f"Erro inesperado: {e}"
                registrar_log(last_error)
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
        run_status = "INCONSISTENT" if stats_gerais['error'] > 0 else "CONSISTENT"
        stats_gerais['tempo_total'] = str(timedelta(seconds=int(elapsed_total.total_seconds())))
        
        stats_gerais['total_processed'] = stats_gerais['success'] + stats_gerais['error']
        stats_gerais['total'] = stats_gerais['total_processed']
        stats_gerais['run_status'] = run_status
        if manifest is not None:
            manifest['finished_at'] = datetime.now().isoformat()
            manifest['duration_sec'] = int(elapsed_total.total_seconds())
            manifest['run_status'] = run_status
            manifest['totals'] = {
                'total_detected': stats_gerais['total_detected'],
                'total_processed': stats_gerais['total_processed'],
                'success': stats_gerais['success'],
                'error': stats_gerais['error'],
                'skipped_duplicate': stats_gerais['skipped_duplicate'],
                'skipped_draft': stats_gerais['skipped_draft'],
                'skipped_non_aso': stats_gerais['skipped_non_aso'],
            }
            if last_error:
                manifest['last_error'] = last_error

        manifest_path = None
        if manifest is not None:
            manifest_path = salvar_manifest(manifest, PASTA_RELATORIOS, execution_id=execution_id)
            if manifest_path:
                manifest['paths']['manifest'] = manifest_path
        
        # So envia email se houver algo processado (sucesso ou erro)
        if stats_gerais['total_detected'] > 0 or stats_gerais['error'] > 0:
            registrar_log("Gerando relatorio e enviando email...")
        
            # 1. Salvar Relatorio JSON/MD
            report_paths = reporter.save_report(stats_gerais)
            if manifest is not None:
                manifest['paths'].update({
                    'report_json': report_paths.get('json') if report_paths else None,
                    'report_md': report_paths.get('md') if report_paths else None,
                    'logs': PASTA_LOGS,
                })
        
            # 2. Enviar Email
            email_status, email_error = enviar_resumo_email(
                TARGET_ACCOUNT,
                stats_gerais,
                execution_id,
                run_status,
                report_paths=report_paths,
                manifest_path=manifest_path,
                logger=logger,
            )
            if manifest is not None:
                manifest['email_status'] = email_status
                manifest['email_error'] = email_error
                if manifest_path:
                    salvar_manifest(manifest, PASTA_RELATORIOS, filepath=manifest_path, execution_id=execution_id)
        else:
            registrar_log("Nada processado, email de resumo nao enviado.")
        
        stats_gerais['finished_at'] = datetime.now().isoformat()
        stats_gerais['run_status'] = run_status
        if last_error:
            stats_gerais['last_error'] = last_error
        salvar_diagnostico_resumo(
            stats_gerais,
            manifest_path=manifest_path,
            report_paths=report_paths if 'report_paths' in locals() else None,
            extra={"logs_dir": PASTA_LOGS}
        )

        registrar_log(f"Mensagens verificadas hoje: {encontrados_hoje}; mensagens processadas: {processados}")
        
        
        
    try:
        return _capta_core()
    finally:
        cleanup_outlook_com()
# ====================================================================
# MAIN
# ====================================================================
if __name__ == "__main__":
    execution_id = str(uuid.uuid4())
    started_at = datetime.now()
    logger.set_execution_id(execution_id)
    manifest = {
        'execution_id': execution_id,
        'started_at': started_at.isoformat(),
        'finished_at': None,
        'duration_sec': None,
        'run_status': None,
        'paths': {},
        'email_status': None,
        'email_error': None,
        'totals': {},
        'items': []
    }
    registrar_log("===== INICIO DO PROCESSAMENTO DIARIO DE ASO =====", context={"execution_id": execution_id})
    try:
        captar_emails(limit=500, execution_id=execution_id, started_at=started_at, manifest=manifest)
    except Exception as e:
        registrar_log(f"Erro fatal: {e}")
        try:
            with open(os.path.join(PASTA_ERROS, "erro_fatal.txt"), "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
        except:
            pass
    registrar_log("===== FIM DO PROCESSAMENTO =====", context={"execution_id": execution_id})
