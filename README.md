# ASOgui (Automacao de ASO)

## Visao geral
Este projeto automatiza o recebimento e o cadastro de Atestados de Saude Ocupacional (ASO) a partir do Outlook. O fluxo busca emails, processa anexos PDF com OCR, extrai dados (nome, CPF, data e funcao) e integra com o bot RPA. Ao final, gera relatorios e envia notificacoes.

## Principais recursos
- Integracao com Outlook via COM (Windows)
- OCR com Tesseract + conversao PDF com Poppler
- Regras de validacao (ASO x rascunho x nao ASO)
- Relatorios JSON + resumo Markdown
- Manifesto de execucao (`manifest_*.json`) e logs estruturados
- Suporte a links do Google Drive em emails (download automatico com filtro)
- Runner/Updater para instalacao onedir com atualizacao via rede ou GitHub
- Scripts de build e atalhos .bat para uso sem terminal

## Requisitos
- Windows 10/11 com Microsoft Outlook configurado
- Python 3.10+ (para executar localmente)
- Tesseract OCR e Poppler (incluidos no pacote quando buildado)
- Acesso ao compartilhamento `P:\ProcessoASO` (default de saidas)

## Configuracao rapida
1) Copie `.env.example` para `.env` e ajuste as variaveis.
2) Instale dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3) Rode o fluxo:
   - Terminal: `python src/main.py`
   - Clique: `run_main.bat`

## Variaveis de ambiente (principais)
Veja `.env.example` para o template completo.
- `PROCESSO_ASO_BASE`: base de saida (default `P:\ProcessoASO`)
- `TESSERACT_PATH`: caminho do `tesseract.exe` (ou pasta)
- `POPPLER_PATH`: caminho do `bin` do Poppler
- `ASO_EMAIL_ACCOUNT`: conta principal no Outlook
- `ASO_MAILBOX_NAME`: nome da mailbox/caixa compartilhada
- `ASO_DAYS_BACK`: dias retroativos (0 = somente hoje)
- `ASO_NOTIFY_TO` ou `ASO_EMAIL_TO`: destinatarios do resumo
- `ASO_EMAIL_FROM`: remetente (opcional)
- `ASO_GDRIVE_NAME_FILTER`: filtro de nome para links Google Drive (default `asos enesa`)
- `ASO_GDRIVE_TIMEOUT_SEC`: timeout de download (segundos)
- `YUBE_URL`, `YUBE_USER`, `YUBE_PASS`, `YUBE_NAV_TIMEOUT`: credenciais e timeout do bot

## Pastas de saida (default)
Base: `PROCESSO_ASO_BASE` (default `P:\ProcessoASO`)
- `processados`
- `em processamento`
- `erros`
- `logs` (ex.: `execution_log_YYYY-MM-DD.jsonl`, `diagnostico_ultima_execucao.txt`)
- `relatorios` (ex.: `relatorio_*.json`, `resumo_execucao_*.md`, `manifest_*.json`)

## Atalhos .bat (sem terminal)
- `run_main.bat` -> executa `src/main.py`
- `run_tests.bat` -> executa pytest
- `build_zip.bat [patch|minor|major|1.2.3]` -> gera ZIP + latest.json + sha256
- `build_windows.bat` -> gera `dist\ASOgui` (onedir)

## Build e release
Build recomendado (gera ZIP de release):
```bash
powershell -ExecutionPolicy Bypass -File scripts\build_aso_zip.ps1
```
- O script aumenta a versao automaticamente (patch). Para `minor`/`major`, use `-Bump`.
- O ZIP e o `latest.json` sao gravados em `dist\`.

## Estrutura do projeto
- `src/`: codigo-fonte do ASOgui
- `src/main.py`: orquestracao principal
- `src/runner.py`: updater/launcher (instalacao onedir)
- `src/reporting.py`: relatorios
- `src/notification.py`: email de resumo
- `src/utils_masking.py`: mascaramento de PII (CPF)
- `scripts/`: builds e empacotamento
- `tests/`: testes automatizados

## Testes
```bash
python -m pytest
```
Ou clique em `run_tests.bat`.

## Script auxiliar: ASO admissional
O `src/aso_admissional_email.py` e um fluxo simples para baixar anexos de ASO admissional.
- Base de saida: `ASO_DEST_BASE` (default `P:\ASO_ADMISSIONAL`)
- Filtros: `ASO_SUBJECT_PREFIX`, `ASO_ATTACH_EXTS`
- Outlook: `ASO_STORE_NAME`, `ASO_MAILBOX_NAME`, `ASO_EMAIL_ACCOUNT`
- Limites: `ASO_MAX_EMAILS`, `ASO_MAPI_SCAN_DEPTH`, `ASO_DAYS_BACK`
Observacao: no script admissional, `ASO_DAYS_BACK` default e 3.

## Documentacao completa
Leia `docs/DOCUMENTACAO_OFICIAL.md` para detalhes corporativos, arquitetura, operacao e troubleshooting.
