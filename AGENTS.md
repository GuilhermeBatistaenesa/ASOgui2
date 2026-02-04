# AGENTS.md - ASOgui

Este arquivo define instrucoes locais para o Codex ao trabalhar neste repositorio.

**Project Root**
`P:\ProcessoASO\Codigo\ASOgui2`

**Resumo Do Projeto**
Automacao de ASO a partir do Outlook. Processa PDFs com OCR (Tesseract + Poppler), extrai dados e integra com bot RPA. Gera relatorios e notificacoes. Inclui Runner/Updater para instalacao onedir.

**Requisitos**
- Windows 10/11 com Outlook configurado.
- Python 3.10+ para executar localmente.
- Tesseract OCR e Poppler (para execucao local) ou `vendor\` para build.
- Playwright (usa Chromium).

**Configuracao Local**
1. Copie `.\.env.example` para `.\.env` e ajuste.
2. Instale dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3. Execute:
   - `python main.py`
   - ou `.\run_main.bat`

**Variaveis De Ambiente (principais)**
- `PROCESSO_ASO_BASE`: base de saida (default `P:\ProcessoASO`)
- `TESSERACT_PATH`: caminho do `tesseract.exe` (ou pasta)
- `POPPLER_PATH`: caminho do `bin` do Poppler
- `ASO_EMAIL_ACCOUNT`, `ASO_MAILBOX_NAME`
- `ASO_DAYS_BACK` (0 = somente hoje)
- `ASO_NOTIFY_TO` ou `ASO_EMAIL_TO`
- `ASO_EMAIL_FROM`
- `ASO_GDRIVE_NAME_FILTER`, `ASO_GDRIVE_TIMEOUT_SEC`
- `YUBE_URL`, `YUBE_USER`, `YUBE_PASS`, `YUBE_NAV_TIMEOUT`

**Pastas De Saida**
Base: `PROCESSO_ASO_BASE`
- `processados`, `em processamento`, `erros`, `logs`, `relatorios`
- Relatorios: `relatorio_*.json`, `resumo_execucao_*.md`, `manifest_*.json`
- Logs: `execution_log_YYYY-MM-DD.jsonl`, `diagnostico_ultima_execucao.txt`

**Testes**
- `python -m pytest`
- ou `.\run_tests.bat`

**Build E Release**
Release recomendado (ZIP + latest.json + sha256):
```bash
powershell -ExecutionPolicy Bypass -File scripts\build_aso_zip.ps1
```
Opcoes:
- `-Bump patch|minor|major` (padrao: patch)
- `-Version 1.2.3`

Build onedir (legacy, sem ZIP):
```bash
powershell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

Saidas:
- `.\dist\ASOgui_*.zip`
- `.\dist\latest.json`
- `.\dist\ASOgui\` (onedir)
- `.\build\` (arquivos intermediarios do PyInstaller)

**Runner/Updater**
- Script: `.\runner.py`
- Config: `.\config.json`
- Build runner (onefile):
  ```bash
  pyinstaller --onefile --noconsole --name ASOguiRunner runner.py
  ```
Campos importantes em `.\config.json`:
- `network_release_dir`, `network_latest_json`
- `github_repo`
- `install_dir`
- `prefer_network`, `allow_prerelease`
- `run_args`
- `log_level`, `ui`
Config pode ser passado com `--config "C:\caminho\config.json"`.

**Script Auxiliar (ASO admissional)**
- `python aso_admissional_email.py`
- Usa `ASO_DEST_BASE`, `ASO_SUBJECT_PREFIX`, `ASO_ATTACH_EXTS`, `ASO_MAX_EMAILS`, `ASO_MAPI_SCAN_DEPTH`

**Ferramentas (Tesseract/Poppler)**
Build usa arquivos em:
- `.\vendor\tesseract\` ou `.\vendor\tesseract\Tesseract-OCR\`
- `.\vendor\poppler\bin\`

Execucao local usa caminhos em `.\.env`:
- `TESSERACT_PATH`
- `POPPLER_PATH`

Pasta de downloads (distribuicao manual):
- `.\tools_downloads\`

**Playwright**
Se faltar browser:
```bash
python -m playwright install chromium
```
O build ZIP instala browsers em `%TEMP%` para evitar bloqueio do OneDrive e depois copia.

**Arquivos Gerados (nao editar)**
- `.\dist\`
- `.\build\`
- `.\.venv*`
- `.\.pytest_cache\`
- `.\__pycache__\`
- `.\.playwright-browsers\`
- `.\.tmp\`

**Segredos**
`.\.env` pode conter credenciais. Nao exibir nem commitar valores reais.

**Referencias**
- `.\README.md`
- `.\README_Runner.md`
- `.\docs\DOCUMENTACAO_OFICIAL.md`
