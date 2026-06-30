from __future__ import annotations

import os
import shutil
import signal
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FRONTEND = ROOT / "frontend"
BACKEND_URL = "http://127.0.0.1:8765/api/health"
FRONTEND_URL = "http://127.0.0.1:5173"
def stop_old_processes() -> None:
    output = subprocess.run(["netstat", "-ano", "-p", "tcp"], capture_output=True, text=True, check=True).stdout
    connections = (line.split() for line in output.splitlines())
    pids = {parts[-1] for parts in connections if len(parts) >= 5 and parts[1].rsplit(":", 1)[-1] in {"5173", "8765"}}
    for pid in pids:
        subprocess.run(["taskkill", "/PID", pid, "/T", "/F"], capture_output=True)
    if pids:
        print(f"Processos antigos encerrados: {', '.join(sorted(pids))}")


def wait_for(url: str, process: subprocess.Popen, timeout: int = 30) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if process.poll() is not None:
            raise RuntimeError(f"Processo finalizado durante a inicialização: {process.args}")
        try:
            with urllib.request.urlopen(url, timeout=1) as response:
                if response.status < 500:
                    return
        except OSError:
            time.sleep(0.3)
    raise TimeoutError(f"O serviço não respondeu em {url}")


def main() -> int:
    os.chdir(ROOT)
    stop_old_processes()
    npm = shutil.which("npm.cmd") or shutil.which("npm")
    if not npm:
        raise RuntimeError("Node.js/npm não encontrado no PATH.")

    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "-r", "requirements.txt"], check=True)
    if not (FRONTEND / "node_modules").exists():
        subprocess.run([npm, "install"], cwd=FRONTEND, check=True)

    flags = subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
    backend = subprocess.Popen([sys.executable, "-m", "uvicorn", "app:app", "--host", "127.0.0.1", "--port", "8765"], creationflags=flags)
    frontend = subprocess.Popen([npm, "run", "dev", "--", "--host", "127.0.0.1", "--port", "5173", "--strictPort"], cwd=FRONTEND, creationflags=flags)
    processes = [backend, frontend]

    try:
        wait_for(BACKEND_URL, backend)
        wait_for(FRONTEND_URL, frontend)
        print(f"Verbetograma disponível em {FRONTEND_URL}")
        if os.environ.get("VERBETOGRAMA_NO_BROWSER") != "1":
            webbrowser.open(FRONTEND_URL)
        while all(process.poll() is None for process in processes):
            time.sleep(0.5)
        return next((process.returncode for process in processes if process.returncode), 0)
    except KeyboardInterrupt:
        return 0
    finally:
        for process in processes:
            if process.poll() is None:
                process.send_signal(signal.CTRL_BREAK_EVENT if os.name == "nt" else signal.SIGTERM)
        for process in processes:
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        raise SystemExit(1)
