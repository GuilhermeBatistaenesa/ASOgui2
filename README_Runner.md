# ASOgui Runner/Updater

## Build (PyInstaller)
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

## Run manually
```
ASOguiRunner.exe
```
or (with Python):
```
python runner.py
```

## Schedule (Task Scheduler)
1) Open Task Scheduler
2) Create Task
3) Action: Start a program
4) Program/script: `C:\ASOgui\ASOguiRunner.exe`
5) Set triggers as desired (daily/hourly)
