# Auditoria de Execucoes (Excel Central)

Este projeto grava automaticamente os dados de execucao do robo ASO em um Excel central de governanca.

**Arquivo central**
- `P:\AuditoriaRobos\Auditoria_Robos.xlsx`

**Fallback (quando o Excel estiver aberto e nao permitir salvar)**
- `P:\AuditoriaRobos\pending\Auditoria_Robos__PENDENTE__<timestamp>.xlsx`
- `P:\AuditoriaRobos\pending\run__<run_id>.json`

## Integracao no ASO
O modulo esta em `src/auditoria_excel.py` e e chamado no final de `src/main.py`.

Funcao principal:
```python
from auditoria_excel import log_run

log_run(run_data, errors=errors_list)
```

O ASO envia:
- `total_processado`, `total_sucesso`, `total_erro`
- `erros_auto_mitigados`, `erros_manuais`
- `resultado_final`
- `observacoes`
- `run_id`, `started_at`, `finished_at`

Os campos automaticos (usuario, host, versao, etc.) sao preenchidos pelo modulo caso nao sejam fornecidos.

## Requisitos
- `openpyxl` (ja adicionado em `requirements.txt`).

## Observacoes
- A planilha e criada automaticamente com as abas `RUNS`, `ERRORS`, `ROBOS` e `DASHBOARD`.
- As validacoes, formatos e graficos sao configurados no primeiro uso.
- Para ambiente, pode-se definir `ASO_ENV=PROD` ou `ASO_ENV=TEST` (padrao: `PROD`).
