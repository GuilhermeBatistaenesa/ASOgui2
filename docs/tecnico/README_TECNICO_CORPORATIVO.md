# README Tecnico Corporativo - ProcessoASO

## Visao geral
Automacao corporativa em producao para leitura de emails de ASO, extracao por OCR, classificacao documental e encaminhamento para o fluxo RPA.

## Escopo
- Operacao do fluxo principal (`src/main.py`)
- Governanca, rastreabilidade e artefatos de execucao
- Fluxo auxiliar de ASO admissional (`src/aso_admissional_email.py`)
- Integracao com o modulo `rpa_yube`

## Fluxo operacional (alto nivel)
Outlook -> OCR -> Yube

## Base operacional
P:\ProcessoASO

## Pastas de artefatos
- em processamento/
- processados/
- erros/
- logs/
- relatorios/
- json/
- releases/

## Convencoes vigentes de artefatos
- Log tecnico: `logs/execution_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.jsonl`
- Manifesto: `json/manifest_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.json`
- Relatorio estruturado: `json/relatorio_execucao_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.json`
- Resumo legivel: `relatorios/resumo_execucao_<YYYY-MM-DD_HH-mm-ss>__<execution_id>.md`
- Indice de processados: `json/processed_index.json`

## Estado de conformidade (2026-03-03)
- Logger e artefatos alinhados ao contrato corporativo.
- Classificacao de erro normalizada para classes corporativas.
- Terminal principal padronizado no formato ENESA.
- Suite local validada com `26 passed, 2 skipped`.
- Pendencia residual: nomeacao formal de ownership e endurecimento do gate de CI.

## Versionamento
- SemVer: MAJOR.MINOR.PATCH
- Changelog obrigatorio em `CHANGELOG.md`

## Ownership
- Responsavel tecnico: pendente de nomeacao formal no documento corporativo
- Responsavel funcional: pendente de nomeacao formal no documento corporativo
