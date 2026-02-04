# ASOgui Runner/Updater

## Objetivo
O Runner instala, atualiza e executa o ASOgui em modo onedir (pasta completa). Ele suporta canal de rede (preferencial) e fallback via GitHub.

---

## Build (ASOgui onedir + tools + browsers + ZIP)
Script recomendado (gera ZIP de release, SHA256 e latest.json):
```bash
powershell -ExecutionPolicy Bypass -File scripts\build_aso_zip.ps1
```

### Auto-bump de versao
Se `-Version` nao for informado, o script incrementa automaticamente (patch):
- `-Bump patch` (padrao)
- `-Bump minor`
- `-Bump major`
- `-Version 1.2.3` (manual)

Atalho sem terminal:
```bat
build_zip.bat
build_zip.bat minor
build_zip.bat 1.2.3
```

### Notas de build
- O script aceita `PYTHON_EXE` para apontar um `python.exe` especifico.
- Se `PYTHON_EXE` nao estiver definido, o script tenta Python 3.12+ ou `py -3`.
- Os browsers do Playwright sao instalados em `%TEMP%` e copiados para o pacote.

---

## Build (legacy onedir sem ZIP)
```bash
powershell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

Notas:
- Cria `.venv-build` para o build.
- Instala browsers do Playwright no projeto (`playwright-browsers` ou `.playwright`).

---

## Runner (PyInstaller)
```bash
pyinstaller --onefile --noconsole --name ASOguiRunner runner.py
```

---

## Configuracao do Runner (`config.json`)
Edite:
- `network_release_dir` / `network_latest_json`
- `github_repo`
- `install_dir`
- `prefer_network`
- `allow_prerelease`
- `run_args`
- `log_level`
- `ui`

---

## Localizacao do config.json
- Se passar `--config "C:\caminho\config.json"`, usa esse arquivo.
- Caso contrario, procura `config.json` ao lado do exe/script.
- Se nao achar, procura no diretorio atual.

---

## Canal de rede (ZIP)
O Runner procura o `latest.json` na pasta de rede e baixa o ZIP completo.
Exemplo de `latest.json`:
```json
{
  "version": "1.0.0",
  "package_filename": "ASOgui_1.0.0.zip",
  "sha256_filename": "ASOgui_1.0.0.sha256"
}
```

---

## Estrutura final instalada
```
C:\ASOgui\
  app\
    current\
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

## Pasta de downloads (Tesseract/Poppler)
Alguns usuarios nao tem Tesseract/Poppler instalados. Para facilitar, coloque os ZIPs aqui:
```
tools_downloads\
  tesseract.zip
  poppler.zip
```

> Observacao: esses ZIPs sao apenas para distribuicao manual. O build oficial usa `vendor\` e embute as ferramentas no pacote.

---

## Update tools (vendor)
Antes do build, garantir:
```
vendor\tesseract\tesseract.exe
vendor\tesseract\tessdata\...
vendor\poppler\bin\...
```

Ou:
```
vendor\tesseract\Tesseract-OCR\tesseract.exe
```

---

## Execucao
- Runner instalado: `ASOguiRunner.exe`
- Python local: `python runner.py`

---

## Agendamento (Task Scheduler)
1) Abrir Task Scheduler
2) Create Task
3) Action: Start a program
4) Program/script: `C:\ASOgui\ASOguiRunner.exe`
5) Definir gatilhos (diario/horario)
