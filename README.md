# ASOgui (Automacao de ASO)

## Visao geral
Este projeto automatiza o recebimento e o cadastro de Atestados de Saude Ocupacional (ASO) a partir do Outlook. O fluxo busca emails, processa anexos PDF com OCR, extrai dados (nome, CPF, data e funcao) e integra com o bot RPA. Ao final, gera relatorios e envia notificacoes.

## Principais recursos
- Integracao com Outlook via COM (Windows)
- OCR com Tesseract + conversao PDF com Poppler
- Regras de validacao (ASO x rascunho x nao ASO)
- Relatorios JSON + resumo Markdown
- Runner/Updater para instalacao onedir com atualizacao via rede ou GitHub
- Scripts de build e atalhos .bat para uso sem terminal

## Requisitos
- Windows 10/11 com Microsoft Outlook configurado
- Python 3.10+ (para executar localmente)
- Tesseract OCR e Poppler (incluidos no pacote quando buildado)

## Configuracao rapida
1) Copie `.env.example` para `.env` e ajuste as variaveis.
2) Instale dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3) Rode o fluxo:
   - Terminal: `python main.py`
   - Clique: `run_main.bat`

## Atalhos .bat (sem terminal)
- `run_main.bat` -> executa `main.py`
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
- `main.py`: orquestracao principal
- `runner.py`: updater/launcher (instalacao onedir)
- `reporting.py`: relatorios
- `notification.py`: email de resumo
- `utils_masking.py`: mascaramento de PII (CPF)
- `scripts/`: builds e empacotamento
- `tests/`: testes automatizados

## Testes
```bash
python -m pytest
```
Ou clique em `run_tests.bat`.

## Documentacao completa
Leia `docs/DOCUMENTACAO_OFICIAL.md` para detalhes corporativos, arquitetura, operacao e troubleshooting.
