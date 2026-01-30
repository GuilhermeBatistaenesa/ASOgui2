# ASOgui Runner/Updater

## Build (ASOgui onedir + tools + browsers ZIP)
Use the script:
```
powershell -ExecutionPolicy Bypass -File scripts\build_aso_zip.ps1
```

## Build (legacy onedir without zip)
Use the script:
```
powershell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

PyInstaller (runner):
```
pyinstaller --onefile --noconsole --name ASOguiRunner runner.py
```

## Configure
Edit `config.json` and set:
- `network_release_dir` / `network_latest_json`
- `github_repo`
- `install_dir`
- `prefer_network`
- `allow_prerelease`
- `run_args`

## Offline run
The ASOgui package is self-contained (tools + browsers inside).

## Run manually
```
ASOguiRunner.exe
```
or (with Python):
```
python runner.py
```

## Final package structure
```
dist\ASOgui\
  ASOgui.exe
  _internal\
  VERSION.txt
  .env
  tools\
    tesseract\tesseract.exe
    poppler\bin\pdftoppm.exe
  playwright-browsers\
```

## ZIP release package
The network channel can point to a ZIP containing the full onedir package.
Example `latest.json`:
```
{
  "version": "1.0.0",
  "package_filename": "ASOgui_1.0.0.zip",
  "sha256_filename": "ASOgui_1.0.0.sha256"
}
```

Generate SHA256:
```
Get-FileHash dist\ASOgui_1.0.0.zip -Algorithm SHA256 | ForEach-Object { $_.Hash + "  ASOgui_1.0.0.zip" } > dist\ASOgui_1.0.0.sha256
```

Note: ASOgui is installed as an ONEDIR package (entire folder).

Install layout:
```
C:\ASOgui\
  app\
    current\   <-- full onedir package here
```

## Update tools
Place vendors before build:
```
vendor\tesseract\tesseract.exe
vendor\poppler\bin\...
```
Then rerun `scripts\build_aso_zip.ps1` (or `scripts\build_windows.ps1`).

## Schedule (Task Scheduler)
1) Open Task Scheduler
2) Create Task
3) Action: Start a program
4) Program/script: `C:\ASOgui\ASOguiRunner.exe`
5) Set triggers as desired (daily/hourly)
