import os
import re
from datetime import date, datetime


def parse_br_date(value):
    if value in (None, "", "Desconhecida"):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _sanitize_table_name(table_name):
    value = (table_name or "dbo.IntegracaoASO").strip()
    if not re.fullmatch(r"[\w\.\[\]]+", value):
        raise ValueError(f"Nome de tabela invalido: {value}")
    return value


def _is_enabled():
    return os.getenv("ASO_SQL_ENABLE", "").strip().lower() in ("1", "true", "yes")


def insert_aso_record(payload, log_fn=None):
    if not _is_enabled():
        return False

    conn_str = (os.getenv("ASO_SQL_CONNECTION_STRING") or "").strip()
    if not conn_str:
        if log_fn:
            log_fn("ASO_SQL_ENABLE=1, mas ASO_SQL_CONNECTION_STRING nao foi configurada.")
        return False

    data_aso = parse_br_date(payload.get("data_aso"))
    if data_aso is None:
        if log_fn:
            log_fn(
                "Registro ASO sem data valida; integracao SQL ignorada.",
                {
                    "origem": payload.get("origem_robo"),
                    "arquivo": payload.get("arquivo_origem"),
                },
            )
        return False

    try:
        import pyodbc
    except ImportError:
        if log_fn:
            log_fn("pyodbc nao instalado; integracao SQL ignorada.")
        return False

    table_name = _sanitize_table_name(os.getenv("ASO_SQL_TABLE", "dbo.IntegracaoASO"))
    timeout = int((os.getenv("ASO_SQL_TIMEOUT_SEC") or "10").strip())

    cpf = re.sub(r"\D", "", str(payload.get("cpf") or ""))
    nome = str(payload.get("nome") or "").strip() or None
    funcao = str(payload.get("funcao_cargo") or "").strip() or None
    arquivo = str(payload.get("arquivo_origem") or "").strip() or None
    referencia = str(payload.get("referencia_origem") or "").strip() or None
    origem = str(payload.get("origem_robo") or "").strip() or "ASO"

    sql = f"""
IF NOT EXISTS (
    SELECT 1
    FROM {table_name}
    WHERE OrigemRobo = ?
      AND CPF = ?
      AND DataASO = ?
      AND ISNULL(ArquivoOrigem, '') = ISNULL(?, '')
)
BEGIN
    INSERT INTO {table_name} (
        OrigemRobo,
        NomeColaborador,
        CPF,
        DataASO,
        FuncaoCargo,
        ArquivoOrigem,
        ReferenciaOrigem,
        DataCaptura
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, SYSDATETIME())
END
"""

    params = [
        origem,
        cpf,
        data_aso,
        arquivo,
        origem,
        nome,
        cpf,
        data_aso,
        funcao,
        arquivo,
        referencia,
    ]

    conn = None
    try:
        conn = pyodbc.connect(conn_str, timeout=timeout)
        cursor = conn.cursor()
        cursor.execute(sql, params)
        conn.commit()
        if log_fn:
            log_fn(
                "Registro ASO enviado para SQL.",
                {
                    "origem": origem,
                    "arquivo": arquivo,
                    "data_aso": data_aso.isoformat(),
                },
            )
        return True
    except Exception as exc:
        if log_fn:
            log_fn(
                "Falha ao gravar registro ASO no SQL.",
                {
                    "erro": str(exc),
                    "origem": origem,
                    "arquivo": arquivo,
                },
            )
        return False
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
