import json
import os
import re
import shutil
import hashlib
import subprocess
import threading
import sys
from datetime import datetime, timedelta
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

# UI notes:
# - The update/run flow runs on a background thread to keep the UI responsive.
# - UI updates are scheduled via root.after(...) to avoid cross-thread Tkinter access.

DEFAULT_CONFIG = {
    "network_release_dir": r"P:\\ProcessoASO\\releases",
    "network_latest_json": None,
    "github_repo": "https://github.com/GuilhermeBatistaenesa/ASOgui2",
    "install_dir": r"C:\\ASOgui",
    "prefer_network": True,
    "allow_prerelease": False,
    "run_args": [],
    "log_level": "INFO",
    "ui": True,
}

_UI_LOG_HOOK = None


def app_base_dir():
    # When frozen (PyInstaller), sys.executable points to the exe path.
    # When running as script, use __file__.
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resolve_config_path(cli_path=None):
    # Priority:
    # 1) --config path (cli_path)
    # 2) config.json next to exe/script
    # 3) config.json in current working directory
    if cli_path:
        return cli_path

    p1 = os.path.join(app_base_dir(), "config.json")
    if os.path.exists(p1):
        return p1

    p2 = os.path.join(os.getcwd(), "config.json")
    if os.path.exists(p2):
        return p2

    # Default to p1 (nicer error message)
    return p1


def parse_args():
    # Minimal argv parsing (stdlib only)
    # Supports: --config "C:\path\config.json"
    cfg = None
    argv = sys.argv[1:]
    for i, a in enumerate(argv):
        if a == "--config" and i + 1 < len(argv):
            cfg = argv[i + 1].strip('"')
    return cfg


def load_config(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Config not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    cfg = DEFAULT_CONFIG.copy()
    cfg.update({k: v for k, v in data.items() if v is not None})
    if not cfg.get("network_latest_json"):
        cfg["network_latest_json"] = os.path.join(cfg["network_release_dir"], "latest.json")
    return cfg


def ensure_dirs(cfg):
    base = cfg["install_dir"]
    app_dir = os.path.join(base, "app", "current")
    cache_dir = os.path.join(base, "app", "cache", "downloads")
    runner_dir = os.path.join(base, "runner")
    os.makedirs(app_dir, exist_ok=True)
    os.makedirs(cache_dir, exist_ok=True)
    os.makedirs(runner_dir, exist_ok=True)
    return app_dir, cache_dir, runner_dir


def log_line(log_path, level, msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} [{level}] {msg}"
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")
    if _UI_LOG_HOOK:
        _UI_LOG_HOOK(line)


def parse_semver(v):
    if not v:
        return (0, 0, 0)
    m = re.match(r"^v?(\d+)\.(\d+)\.(\d+)", str(v).strip())
    if not m:
        return (0, 0, 0)
    return tuple(int(x) for x in m.groups())


def compare_semver(a, b):
    ta = parse_semver(a)
    tb = parse_semver(b)
    if ta < tb:
        return -1
    if ta > tb:
        return 1
    return 0


def read_current_version(app_dir):
    version_file = os.path.join(app_dir, "version.txt")
    if not os.path.exists(version_file):
        return "0.0.0"
    try:
        with open(version_file, "r", encoding="utf-8") as f:
            return f.read().strip() or "0.0.0"
    except Exception:
        return "0.0.0"


def write_current_version(app_dir, version):
    version_file = os.path.join(app_dir, "version.txt")
    with open(version_file, "w", encoding="utf-8") as f:
        f.write(str(version))


def acquire_lock(lock_path, max_age_minutes=30):
    if os.path.exists(lock_path):
        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(lock_path))
            if datetime.now() - mtime < timedelta(minutes=max_age_minutes):
                return False, "already running"
            os.remove(lock_path)
        except Exception:
            return False, "lock error"
    try:
        with open(lock_path, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))
        return True, "locked"
    except Exception as e:
        return False, f"lock failed: {e}"


def release_lock(lock_path):
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
    except Exception:
        pass


def safe_read_text(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read().strip()


def fetch_network_latest(cfg):
    latest_path = cfg["network_latest_json"]
    base_dir = cfg["network_release_dir"]
    try:
        with open(latest_path, "r", encoding="utf-8") as f:
            meta = json.load(f)
        version = meta.get("version")
        exe_name = meta.get("exe_filename")
        sha_name = meta.get("sha256_filename")
        if not version or not exe_name:
            return None, "network metadata incomplete"
        exe_path = os.path.join(base_dir, exe_name)
        if not os.path.exists(exe_path):
            return None, "network exe missing"
        sha_value = None
        if sha_name:
            sha_path = os.path.join(base_dir, sha_name)
            if os.path.exists(sha_path):
                sha_value = safe_read_text(sha_path).split()[0]
        return {
            "channel": "network",
            "version": version,
            "exe_path": exe_path,
            "exe_name": exe_name,
            "sha256": sha_value,
        }, None
    except Exception as e:
        return None, f"network error: {e}"


def github_request(url, token=None, timeout=(3, 7)):
    headers = {
        "User-Agent": "ASOguiRunner/1.0",
        "Accept": "application/vnd.github+json",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = Request(url, headers=headers)
    return urlopen(req, timeout=timeout[0] + timeout[1])


def fetch_github_latest(cfg):
    repo = cfg["github_repo"]
    token = os.getenv("GITHUB_TOKEN")
    allow_prerelease = bool(cfg.get("allow_prerelease"))
    try:
        if allow_prerelease:
            url = f"https://api.github.com/repos/{repo}/releases"
            resp = github_request(url, token=token)
            data = json.loads(resp.read().decode("utf-8"))
            releases = [r for r in data if r.get("tag_name")]
            if not releases:
                return None, "no releases"
            rel = releases[0]
        else:
            url = f"https://api.github.com/repos/{repo}/releases/latest"
            resp = github_request(url, token=token)
            rel = json.loads(resp.read().decode("utf-8"))

        tag = rel.get("tag_name", "")
        version = tag.lstrip("v")
        assets = rel.get("assets", [])
        exe_asset = None
        sha_asset = None
        for a in assets:
            name = a.get("name", "")
            if name.endswith(".exe") and name.startswith("ASOgui_"):
                exe_asset = a
            if name.endswith(".sha256") and name.startswith("ASOgui_"):
                sha_asset = a
        if not exe_asset:
            return None, "github exe asset missing"

        sha_value = None
        if sha_asset:
            sha_url = sha_asset.get("browser_download_url")
            try:
                sresp = github_request(sha_url, token=token)
                sha_value = sresp.read().decode("utf-8").strip().split()[0]
            except Exception:
                sha_value = None

        return {
            "channel": "github",
            "version": version,
            "exe_url": exe_asset.get("browser_download_url"),
            "exe_name": exe_asset.get("name"),
            "sha256": sha_value,
        }, None
    except (HTTPError, URLError) as e:
        return None, f"github error: {e}"
    except Exception as e:
        return None, f"github error: {e}"


def choose_best_release(a, b, prefer_network=True):
    if not a and not b:
        return None
    if a and not b:
        return a
    if b and not a:
        return b
    cmp = compare_semver(a["version"], b["version"])
    if cmp == 0:
        return a if prefer_network else b
    return a if cmp > 0 else b


def download_file(url, dest_path, token=None, timeout=(3, 7)):
    headers = {
        "User-Agent": "ASOguiRunner/1.0",
        "Accept": "application/octet-stream",
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = Request(url, headers=headers)
    with urlopen(req, timeout=timeout[0] + timeout[1]) as resp:
        with open(dest_path, "wb") as f:
            while True:
                chunk = resp.read(8192)
                if not chunk:
                    break
                f.write(chunk)


def verify_sha256(file_path, expected):
    if not expected:
        return True
    h = hashlib.sha256()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower() == expected.lower()


def atomic_replace(src, dst):
    os.replace(src, dst)


def run_installed_app(app_dir, args):
    exe_path = os.path.join(app_dir, "ASOgui.exe")
    if not os.path.exists(exe_path):
        return 1
    cmd = [exe_path] + (args or [])
    proc = subprocess.run(cmd, check=False)
    return proc.returncode


def run_flow(cfg):
    app_dir, cache_dir, runner_dir = ensure_dirs(cfg)
    log_path = os.path.join(runner_dir, "runner.log")
    lock_path = os.path.join(runner_dir, "runner.lock")

    ok, msg = acquire_lock(lock_path)
    if not ok:
        log_line(log_path, "WARN", f"Runner already running or lock error: {msg}")
        return 2

    try:
        current_version = read_current_version(app_dir)
        log_line(log_path, "INFO", f"Current version: {current_version}")

        net_rel, net_err = fetch_network_latest(cfg)
        if net_err:
            log_line(log_path, "WARN", f"Network channel failed: {net_err}")

        gh_rel, gh_err = fetch_github_latest(cfg)
        if gh_err:
            log_line(log_path, "WARN", f"GitHub channel failed: {gh_err}")

        chosen = choose_best_release(net_rel, gh_rel, prefer_network=bool(cfg.get("prefer_network")))
        if not chosen:
            log_line(log_path, "WARN", "No release available, running installed version")
            return run_installed_app(app_dir, cfg.get("run_args"))

        if compare_semver(chosen["version"], current_version) <= 0:
            log_line(log_path, "INFO", f"No update needed. Latest: {chosen['version']}")
            return run_installed_app(app_dir, cfg.get("run_args"))

        log_line(log_path, "INFO", f"Chosen release: {chosen['version']} ({chosen['channel']})")

        tmp_path = os.path.join(cache_dir, chosen["exe_name"] + ".tmp")
        if chosen["channel"] == "network":
            shutil.copy2(chosen["exe_path"], tmp_path)
        else:
            download_file(chosen["exe_url"], tmp_path, token=os.getenv("GITHUB_TOKEN"))

        if not verify_sha256(tmp_path, chosen.get("sha256")):
            log_line(log_path, "ERROR", "SHA256 mismatch; not installing")
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            return run_installed_app(app_dir, cfg.get("run_args"))

        dst = os.path.join(app_dir, "ASOgui.exe")
        atomic_replace(tmp_path, dst)
        write_current_version(app_dir, chosen["version"])
        log_line(log_path, "INFO", f"Installed version: {chosen['version']}")

        code = run_installed_app(app_dir, cfg.get("run_args"))
        log_line(log_path, "INFO", f"Exit code: {code}")
        return code
    finally:
        release_lock(lock_path)


def run_headless(cfg):
    return run_flow(cfg)


def run_ui(cfg):
    from tkinter import Tk, Text, Button, messagebox
    from tkinter import ttk

    app_dir, cache_dir, runner_dir = ensure_dirs(cfg)
    log_path = os.path.join(runner_dir, "runner.log")

    root = Tk()
    root.title("ASO - Atualizador e Execução")
    root.geometry("700x420")

    status_label = ttk.Label(root, text="Checando atualizações...", font=("Segoe UI", 14))
    status_label.pack(pady=10)

    pb = ttk.Progressbar(root, mode="indeterminate")
    pb.pack(fill="x", padx=20, pady=5)
    pb.start(10)

    log_box = Text(root, height=14, wrap="word")
    log_box.pack(fill="both", expand=True, padx=20, pady=10)
    log_box.config(state="disabled")

    running = {"value": True}
    result_code = {"value": 1}

    def set_status(text):
        status_label.config(text=text)

    def ui_log(line):
        def _append():
            log_box.config(state="normal")
            log_box.insert("end", line + "\n")
            log_box.see("end")
            log_box.config(state="disabled")
        root.after(0, _append)

    def on_open_log():
        if os.path.exists(log_path):
            subprocess.Popen(["notepad.exe", log_path])
        else:
            messagebox.showwarning("Log", "Arquivo de log não encontrado.")

    def on_exit():
        if running["value"]:
            if messagebox.askyesno("Sair", "Execução em andamento. Deseja sair?"):
                root.quit()
        else:
            root.quit()

    btn_frame = ttk.Frame(root)
    btn_frame.pack(pady=5)

    Button(btn_frame, text="Abrir Log", command=on_open_log).pack(side="left", padx=5)
    Button(btn_frame, text="Sair", command=on_exit).pack(side="left", padx=5)

    def bg_task():
        global _UI_LOG_HOOK
        try:
            set_status("Checando atualizações...")
            _UI_LOG_HOOK = ui_log
            code = run_flow(cfg)
            result_code["value"] = code
            set_status("Finalizado com sucesso" if code == 0 else f"Finalizado com erro (code={code})")
        except Exception as e:
            log_line(log_path, "ERROR", f"Falha no runner: {e}")
            set_status("Falha no runner")
            result_code["value"] = 1
        finally:
            running["value"] = False
            pb.stop()

    threading.Thread(target=bg_task, daemon=True).start()
    root.mainloop()
    return result_code["value"]


def main():
    cfg_cli = parse_args()
    cfg_path = resolve_config_path(cfg_cli)
    cfg = load_config(cfg_path)

    if cfg.get("ui", True):
        return run_ui(cfg)
    return run_headless(cfg)


if __name__ == "__main__":
    raise SystemExit(main())
