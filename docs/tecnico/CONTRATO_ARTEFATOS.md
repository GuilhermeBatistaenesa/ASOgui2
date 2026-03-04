# Contrato de Artefatos - ProcessoASO

## Convencoes
- Log: execution_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.jsonl
- Manifesto: manifest_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.json
- Relatorio: relatorio_execucao_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.txt|md|json

## Campos minimos de log
- timestamp_utc
- robot_name
- robot_version
- execution_id
- run_status
- step
- event_type
- severity
- message
- source_file
- correlation_keys

## Classes de erro
- BUSINESS_VALIDATION
- EXTERNAL_DEPENDENCY
- INTEGRATION
- DATA_QUALITY
- INFRA_IO
- UNEXPECTED
