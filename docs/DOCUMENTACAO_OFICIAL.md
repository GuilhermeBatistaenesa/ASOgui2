# DOCUMENTACAO TECNICA OFICIAL - ASOgui (RPA ASO)

**Versao do documento**: 2.0
**Data**: 30/01/2026
**Projeto**: Automacao de Recebimento e Cadastro de Atestados de Saude Ocupacional (ASO)
**Publico-alvo**: Equipes de Automacao, TI, Suporte, Operacoes e Auditoria

---

## 1. Sumario executivo
O ASOgui e uma solucao corporativa de automacao que monitora a caixa de email corporativa, identifica ASOs, extrai dados via OCR e integra o resultado com o bot RPA para cadastro. O sistema reduz tempo operacional, padroniza processos e gera evidencias auditaveis (logs, relatorios e manifestos).

---

## 2. Escopo
### 2.1 O que o sistema faz
- Ler emails do Outlook (conta e caixa configuradas)
- Baixar anexos PDF
- Converter PDF em imagens
- Executar OCR com Tesseract
- Extrair Nome, CPF, Data e Funcao/Cargo
- Validar se o documento e ASO, rascunho ou nao ASO
- Salvar PDF renomeado na pasta de destino
- Acionar o bot RPA (modulo `rpa_yube`)
- Gerar relatorios (JSON + Markdown)
- Enviar email de resumo

### 2.2 O que o sistema nao faz
- Nao cadastra manualmente no sistema final (isso e papel do bot RPA)
- Nao substitui a governanca de dados (mantem apenas mascaramento basico)
- Nao executa OCR offline sem Tesseract/Poppler

---

## 3. Arquitetura e componentes
### 3.1 Componentes principais
- **ASOgui (main.py)**: orquestracao e processamento principal
- **Runner/Updater (runner.py)**: instalacao e atualizacao onedir
- **OCR Stack**: Tesseract + Poppler + pdf2image + pytesseract
- **Integracao Outlook**: pywin32 (COM)
- **RPA**: modulo proprietario `rpa_yube`
- **Relatorios**: `reporting.py` e `notification.py`

### 3.2 Fluxo de dados (alto nivel)
1. Outlook -> coleta de emails
2. Anexo PDF -> conversao em imagens
3. OCR -> texto bruto
4. Parser -> dados estruturados
5. Validacao -> ASO/rascunho/nao ASO
6. Saida -> PDF renomeado + manifest + relatorios
7. Integracao -> bot RPA
8. Notificacao -> email de resumo

### 3.3 Diagrama textual
```
Outlook -> Download PDF -> pdf2image -> OCR (Tesseract)
      -> Parser/Validacao -> Renomeia/Salva -> rpa_yube
      -> Relatorio JSON/MD -> Email resumo
```

---

## 4. Estrutura do repositorio
Principais arquivos e pastas:
- `main.py`: fluxo principal
- `runner.py`: updater/launcher
- `reporting.py`: relatorios
- `notification.py`: email resumo
- `utils_masking.py`: mascaramento de CPF
- `idempotency.py`: prevencao de duplicidade
- `outcomes.py`: enums de status
- `scripts/`: scripts de build
- `vendor/`: binarios empacotados (Tesseract/Poppler)
- `tests/`: suite de testes
- `docs/`: documentacao oficial

---

## 5. Requisitos tecnicos
### 5.1 Sistema operacional
- Windows 10/11 (obrigatorio para Outlook COM)

### 5.2 Dependencias
- Python 3.10+
- `requirements.txt` inclui: pywin32, pytesseract, pdf2image, Pillow, python-dotenv, playwright, pytest, pytest-mock

### 5.3 Dependencias externas
- **Tesseract OCR**
- **Poppler (pdf2image)**

> Observacao: ao buildar o pacote, essas dependencias sao embarcadas no `dist` e nao precisam ser instaladas na maquina final.

---

## 6. Configuracao
### 6.1 `.env`
Copie `.env.example` para `.env` e configure as variaveis:

| Variavel | Descricao | Exemplo |
| --- | --- | --- |
| `TESSERACT_PATH` | Caminho absoluto do `tesseract.exe` | `C:\Program Files\Tesseract-OCR\tesseract.exe` |
| `POPPLER_PATH` | Caminho do Poppler (bin) | `C:\Tools\poppler\bin` |
| `ASO_EMAIL_ACCOUNT` | Conta principal Outlook | `aso@empresa.com.br` |
| `ASO_MAILBOX_NAME` | Nome da caixa compartilhada | `Aso` |
| `ASO_NOTIFY_TO` | Destinatarios de resumo | `a@empresa.com.br;b@empresa.com.br` |
| `ASO_EMAIL_FROM` | Remetente (opcional) | `aso@empresa.com.br` |
| `ASO_DAYS_BACK` | Dias retroativos | `0` |

### 6.2 `config.json` (Runner)
Usado pelo updater:

| Campo | Descricao |
| --- | --- |
| `network_release_dir` | Pasta de releases na rede |
| `network_latest_json` | Caminho do `latest.json` |
| `github_repo` | Repo no GitHub (fallback) |
| `install_dir` | Diretorio de instalacao (`C:\ASOgui`) |
| `prefer_network` | Prefere canal de rede |
| `allow_prerelease` | Permite pre-release |
| `run_args` | Argumentos passados ao app |
| `log_level` | Nivel de log |

---

## 7. Build, versionamento e release
### 7.1 Scripts oficiais
- `scripts\build_aso_zip.ps1` (recomendado): gera ZIP completo + SHA256 + `latest.json`
- `scripts\build_windows.ps1`: gera `dist\ASOgui` (onedir)

### 7.2 Auto-bump de versao
O `build_aso_zip.ps1` incrementa a versao automaticamente (patch) quando `-Version` nao e informado. Parametros:
- `-Bump patch` (padrao)
- `-Bump minor`
- `-Bump major`
- `-Version 1.2.3` (manual)

A versao e persistida em `version.txt` e `VERSION.txt`.

### 7.3 Atalhos `.bat`
- `build_zip.bat` (duplo clique) gera ZIP com bump automatico
- `build_zip.bat minor` ou `major` para bump especifico
- `build_zip.bat 1.2.3` para versao fixa
- `build_windows.bat` para build onedir

---

## 8. Empacotamento e dependencias externas
### 8.1 Layout de vendors
Esperado no repositorio (antes do build):
```
vendor\tesseract\tesseract.exe
vendor\tesseract\tessdata\...
vendor\tesseract\libtesseract-5.dll
vendor\poppler\bin\pdftoppm.exe
```

Tambem e aceito:
```
vendor\tesseract\Tesseract-OCR\tesseract.exe
```

O build copia a pasta correta para `dist\ASOgui\tools\tesseract`.

### 8.2 Estrutura final (onedir)
```
dist\ASOgui\
  ASOgui.exe
  _internal\
  VERSION.txt
  .env
  tools\
    tesseract\...
    poppler\bin\...
  playwright-browsers\...
```

---

## 9. Operacao
### 9.1 Execucao manual (desenvolvimento)
- Terminal: `python main.py`
- Clique: `run_main.bat`

### 9.2 Execucao com Runner (producao)
- `ASOguiRunner.exe` ou `python runner.py`
- O runner faz update (rede/GitHub), instala e executa.

### 9.3 Agendamento (Task Scheduler)
1. Abrir Agendador de Tarefas
2. Criar tarefa
3. Acao: executar `C:\ASOgui\ASOguiRunner.exe`
4. Definir gatilhos (diario/horario)

---

## 10. Logs, relatorios e evidencias
### 10.1 Logs
- Pasta `logs/`
- Arquivo de diagnostico: `logs\diagnostico_ultima_execucao.txt`
- Runner: `aso_last_run.log`

### 10.2 Relatorios
- JSON: `relatorios\relatorio_YYYYMMDD_HHMMSS.json`
- Markdown: `relatorios\resumo_execucao_YYYYMMDD_HHMMSS.md`

### 10.3 Mascaramento de PII
- CPF e mascarado em relatorios e textos com `utils_masking.py`
- O mascaramento preserva os ultimos digitos para rastreabilidade

---

## 11. Testes
- Rodar todos os testes:
  ```bash
  python -m pytest
  ```
- Atalho: `run_tests.bat`

A suite cobre:
- OCR fallback
- Parser e validacoes
- Runner (semver, SHA256, locking)
- Relatorios e mascaramento
- Idempotencia

---

## 12. Troubleshooting (principais erros)
### 12.1 `libtesseract-5.dll` nao encontrado
**Causa**: pacote sem DLLs do Tesseract.
**Solucao**: garantir que `vendor\tesseract` contenha a instalacao completa (DLLs + tessdata) antes do build.

### 12.2 OCR retorna vazio
**Causa**: Tesseract nao localizado ou PDF de baixa qualidade.
**Solucao**: revisar `TESSERACT_PATH`, validar `tools\tesseract` no pacote e ajustar qualidade do PDF.

### 12.3 Erro de Outlook COM
**Causa**: Outlook nao aberto, bloqueio de seguranca, perfil incorreto.
**Solucao**: abrir Outlook, confirmar perfil/conta e reiniciar o processo.

### 12.4 Poppler nao encontrado
**Causa**: `POPPLER_PATH` incorreto ou ausente no pacote.
**Solucao**: revisar `.env` e o conteudo de `vendor\poppler\bin`.

### 12.5 Runner nao atualiza
**Causa**: `latest.json` incorreto ou falta do pacote na rede.
**Solucao**: validar `config.json`, SHA256 e permissao de acesso ao compartilhamento.

---

## 13. Seguranca e conformidade
- Nao versionar `.env` com senhas
- Proteger pastas `logs/` e `relatorios/` com controle de acesso
- Restringir acesso ao `config.json` no ambiente de producao
- Registrar alteracoes com controle de versao e changelog interno

---

## 14. Checklist de deploy
1. Atualizar codigo e testes
2. Executar `run_tests.bat`
3. Gerar ZIP com `build_zip.bat`
4. Validar `latest.json` e `*.sha256`
5. Publicar ZIP e JSON no canal de rede
6. Testar instalacao via Runner
7. Monitorar logs na primeira execucao

---

## 15. Responsaveis
**Equipe**: Automacao / Desenvolvimento
**Suporte**: TI Operacional

---
