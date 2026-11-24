"""Painel local do Excel Sync: janela pequena com Sincronizar, Manter sincronia e Parar.

Só biblioteca padrão do Python (http.server) + uma janela do Edge/Chrome em modo app.
Uso: python sync_panel.py
"""
import json
import logging
import os
import secrets
import subprocess
import threading
import time
import webbrowser
from collections import deque
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from config import POLL_INTERVAL_SECONDS
from logger import setup_logger
from monitor import start_monitor
from processor import synchronize

HOST = "127.0.0.1"          # somente esta máquina
WINDOW_SIZE = "400,480"
CLOSE_GRACE_SECONDS = 4.0   # tolera F5/recarregar antes de encerrar o programa
MONITOR_JOIN_TIMEOUT_SECONDS = 60
LOG_BUFFER_LINES = 300


class PanelLogHandler(logging.Handler):
    """Guarda as últimas linhas do logger 'excel_sync' para a janela exibir."""

    def __init__(self):
        super().__init__(logging.INFO)
        self._lines = deque(maxlen=LOG_BUFFER_LINES)
        self._next_id = 1
        self._lock = threading.Lock()
        self.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S"))

    def emit(self, record):
        text = self.format(record)
        with self._lock:
            self._lines.append((self._next_id, record.levelname, text))
            self._next_id += 1

    def since(self, last_id):
        with self._lock:
            return [{"id": i, "level": lv, "text": tx} for i, lv, tx in self._lines if i > last_id]


class SyncService:
    """Controla sincronização manual e monitoramento contínuo sem execuções simultâneas."""

    def __init__(self, logger, log_handler):
        self.logger = logger
        self.log_handler = log_handler
        self._sync_lock = threading.Lock()
        self._stop_event = threading.Event()
        self._monitor_thread = None
        self.last_sync = None  # {"ok": bool, "at": "HH:MM:SS"}

    @property
    def is_monitoring(self):
        return self._monitor_thread is not None and self._monitor_thread.is_alive()

    def _run_sync(self, logger, blocking):
        # Um único lock evita duas gravações simultâneas na aba do Google Sheets
        # (clique manual durante uma sincronização do monitor, por exemplo).
        if not self._sync_lock.acquire(blocking=blocking):
            logger.info("Sincronização já em andamento; pedido ignorado.")
            return False
        try:
            synchronize(logger)
            self.last_sync = {"ok": True, "at": datetime.now().strftime("%H:%M:%S")}
            return True
        except Exception:
            self.last_sync = {"ok": False, "at": datetime.now().strftime("%H:%M:%S")}
            raise
        finally:
            self._sync_lock.release()

    def sync_once(self):
        threading.Thread(target=self._manual_sync, daemon=True).start()

    def _manual_sync(self):
        self.logger.info("Sincronização manual solicitada.")
        try:
            self._run_sync(self.logger, blocking=False)
        except Exception:
            self.logger.exception("Falha na sincronização manual.")

    def start_monitoring(self):
        if self.is_monitoring:
            return False
        stop_event = threading.Event()
        self._stop_event = stop_event
        self._monitor_thread = threading.Thread(target=self._monitor_main, args=(stop_event,), daemon=True)
        self._monitor_thread.start()
        return True

    def _monitor_main(self, stop_event):
        try:
            start_monitor(self.logger, stop_event, lambda lg: self._run_sync(lg, blocking=True))
        except Exception:
            self.logger.exception("Monitoramento encerrado por erro.")

    def stop_monitoring(self):
        # Não interrompe uma sincronização em curso: ela termina e o loop encerra em seguida.
        if self.is_monitoring:
            self.logger.info("Parando: a sincronização em curso, se houver, termina antes.")
        self._stop_event.set()

    def shutdown(self):
        self.stop_monitoring()
        if self._monitor_thread is not None:
            self._monitor_thread.join(timeout=MONITOR_JOIN_TIMEOUT_SECONDS)

    def status(self, since):
        monitoring = self.is_monitoring
        return {
            "monitoring": monitoring,
            "stopping": monitoring and self._stop_event.is_set(),
            "syncing": self._sync_lock.locked(),
            "poll_interval": POLL_INTERVAL_SECONDS,
            "last_sync": self.last_sync,
            "logs": self.log_handler.since(since),
        }


class PanelServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, service, token):
        super().__init__((HOST, 0), PanelRequestHandler)  # porta 0 = livre, escolhida pelo SO
        self.service = service
        self.token = token
        self.close_requested_at = None


class PanelRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):  # silencia o log de acesso HTTP
        pass

    def _send(self, code, body, content_type="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _authorized(self, query):
        return secrets.compare_digest(query.get("t", [""])[0], self.server.token)

    def do_GET(self):
        url = urlparse(self.path)
        query = parse_qs(url.query)
        if url.path == "/":
            if not self._authorized(query):
                self._send(403, "Acesso negado.", "text/plain; charset=utf-8")
                return
            self._send(200, PAGE_HTML.replace("__TOKEN__", self.server.token), "text/html; charset=utf-8")
        elif url.path == "/api/status" and self._authorized(query):
            self.server.close_requested_at = None  # página viva (ex.: recarregada): cancela o encerramento
            try:
                since = int(query.get("since", ["0"])[0])
            except ValueError:
                since = 0
            self._send(200, json.dumps(self.server.service.status(since)))
        else:
            self._send(404, "{}")

    def do_POST(self):
        url = urlparse(self.path)
        query = parse_qs(url.query)
        if not self._authorized(query):
            self._send(403, "{}")
            return
        service = self.server.service
        if url.path == "/api/sync":
            service.sync_once()
        elif url.path == "/api/start":
            if service.start_monitoring():
                service.logger.info("Manter sincronia ativado.")
        elif url.path == "/api/stop":
            service.stop_monitoring()
        elif url.path == "/api/closing":
            self.server.close_requested_at = time.monotonic()
        else:
            self._send(404, "{}")
            return
        self._send(200, "{}")


def open_panel_window(url):
    """Abre o painel em janela pequena (Edge/Chrome em modo app); senão, no navegador padrão."""
    candidates = [
        r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
        r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe",
        r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
        r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
        r"%LocalAppData%\Google\Chrome\Application\chrome.exe",
    ]
    for candidate in candidates:
        exe = os.path.expandvars(candidate)
        if os.path.isfile(exe):
            subprocess.Popen([exe, f"--app={url}", f"--window-size={WINDOW_SIZE}"])
            return "janela app"
    webbrowser.open(url)
    return "navegador padrão"


def _close_watcher(server):
    while True:
        time.sleep(0.5)
        requested = server.close_requested_at
        if requested is not None and time.monotonic() - requested > CLOSE_GRACE_SECONDS:
            server.service.logger.info("Janela fechada; encerrando o programa.")
            server.service.shutdown()
            server.shutdown()
            return


PAGE_HTML = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel Sync</title>
<style>
:root{--bg:#f6f7f9;--fg:#1c2128;--card:#fff;--line:#d8dce2;--accent:#0b6bcb;--ok:#1a7f37;--warn:#9a6700;--err:#cf222e;--mute:#6a737d}
@media (prefers-color-scheme:dark){:root{--bg:#14171c;--fg:#e6e9ee;--card:#1d222a;--line:#323a45;--accent:#4c9aff;--ok:#3fb950;--warn:#d29922;--err:#ff7b72;--mute:#8b949e}}
*{box-sizing:border-box}
body{margin:0;padding:12px;background:var(--bg);color:var(--fg);font:14px system-ui,Segoe UI,sans-serif;display:flex;flex-direction:column;height:100vh;gap:10px}
.row{display:flex;gap:6px}
button{flex:1;padding:9px 6px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--fg);font:inherit;cursor:pointer}
button.primary{background:var(--accent);border-color:var(--accent);color:#fff}
button:disabled{opacity:.45;cursor:default}
#status{display:flex;align-items:center;gap:8px;font-weight:600}
#dot{width:10px;height:10px;border-radius:50%;background:var(--mute)}
#last{color:var(--mute);font-size:12px}
#log{flex:1;overflow:auto;background:var(--card);border:1px solid var(--line);border-radius:6px;padding:8px;font:11.5px ui-monospace,Consolas,monospace}
#log div{white-space:pre-wrap;word-break:break-word;margin-bottom:3px}
.WARNING{color:var(--warn)}.ERROR,.CRITICAL{color:var(--err)}
</style></head><body>
<div class="row">
  <button id="bSync" class="primary">Sincronizar</button>
  <button id="bKeep">Manter sincronia</button>
  <button id="bStop">Parar</button>
</div>
<div><div id="status"><span id="dot"></span><span id="state">Conectando...</span></div><div id="last"></div></div>
<div id="log"></div>
<script>
const TOKEN="__TOKEN__";
const $=id=>document.getElementById(id);
let lastId=0;
const api=(path,method)=>fetch(path+(path.includes("?")?"&":"?")+"t="+TOKEN,{method:method||"GET"});
const post=path=>api(path,"POST").then(poll);
$("bSync").onclick=()=>post("/api/sync");
$("bKeep").onclick=()=>post("/api/start");
$("bStop").onclick=()=>post("/api/stop");
window.addEventListener("pagehide",()=>navigator.sendBeacon("/api/closing?t="+TOKEN));
function render(s){
  let text="Parado",color="var(--mute)";
  if(s.stopping){text="Parando...";color="var(--warn)"}
  else if(s.monitoring){text="Monitorando (a cada "+s.poll_interval+"s)";color="var(--ok)"}
  if(s.syncing){text+=" · sincronizando...";color="var(--accent)"}
  $("state").textContent=text;$("dot").style.background=color;
  $("bSync").disabled=s.syncing;
  $("bKeep").disabled=s.monitoring;
  $("bStop").disabled=!s.monitoring||s.stopping;
  $("last").textContent=s.last_sync?("Última sincronização: "+(s.last_sync.ok?"OK":"COM ERRO")+" às "+s.last_sync.at):"Nenhuma sincronização nesta sessão.";
  const log=$("log"),atEnd=log.scrollTop+log.clientHeight>=log.scrollHeight-20;
  for(const l of s.logs){const d=document.createElement("div");d.className=l.level;d.textContent=l.text;log.appendChild(d);lastId=l.id}
  while(log.children.length>300)log.removeChild(log.firstChild);
  if(atEnd)log.scrollTop=log.scrollHeight;
}
function poll(){
  return api("/api/status?since="+lastId).then(r=>r.json()).then(render).catch(()=>{$("state").textContent="Sem conexão com o programa"});
}
poll();setInterval(poll,1000);
</script></body></html>
"""


def main():
    logger = setup_logger()
    log_handler = PanelLogHandler()
    logger.addHandler(log_handler)
    service = SyncService(logger, log_handler)
    token = secrets.token_urlsafe(16)
    server = PanelServer(service, token)
    url = f"http://{HOST}:{server.server_address[1]}/?t={token}"
    logger.info("Painel Excel Sync em %s (%s).", url.split("?")[0], open_panel_window(url))
    # O token não vai para o arquivo de log; só para o console, caso a janela não abra.
    print(f"Se a janela não abrir, copie este endereço no navegador:\n{url}", flush=True)
    threading.Thread(target=_close_watcher, args=(server,), daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("Interrompido pelo teclado; encerrando.")
        service.shutdown()
    finally:
        server.server_close()
