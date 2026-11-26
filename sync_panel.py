"""Painel local do Excel Sync.

Menu inicial com dois fluxos:
- Planilha de Manutenção: exporta o Excel local para o Google Sheets.
- Planilha Medidores: importa 3 Google Sheets para uma planilha Excel local.

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

from config import (
    MEDIDORES_IMPORT_INTERVAL_SECONDS,
    POLL_INTERVAL_SECONDS,
    validate_config,
    validate_medidores_config,
)
from logger import setup_logger
from medidores import importar
from monitor import start_monitor
from processor import synchronize

HOST = "127.0.0.1"          # somente esta máquina
WINDOW_SIZE = "420,600"
CLOSE_GRACE_SECONDS = 4.0   # tolera F5/recarregar antes de encerrar o programa
MONITOR_JOIN_TIMEOUT_SECONDS = 60
LOG_BUFFER_LINES = 300


class PanelLogHandler(logging.Handler):
    """Guarda as últimas linhas do logger de um fluxo para a janela exibir."""

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


class BaseFlowService:
    """Controla execução manual e monitoramento contínuo de um fluxo."""

    title = ""
    last_label = "Última execução"
    running_label = "executando..."
    manual_log = "Execução manual solicitada."
    monitor_started_log = "Monitoramento ativado."

    def __init__(self, logger, log_handler, poll_interval):
        self.logger = logger
        self.log_handler = log_handler
        self.poll_interval = poll_interval
        self._sync_lock = threading.Lock()
        self._stop_event = threading.Event()
        self._monitor_thread = None
        self.last_sync = None  # {"ok": bool, "at": "HH:MM:SS"}

    @property
    def is_monitoring(self):
        return self._monitor_thread is not None and self._monitor_thread.is_alive()

    def _execute(self, logger):
        raise NotImplementedError

    def _run_sync(self, logger, blocking):
        # Um único lock evita duas gravações simultâneas no destino
        # (clique manual durante uma execução do monitor, por exemplo).
        if not self._sync_lock.acquire(blocking=blocking):
            logger.info("Execução já em andamento; pedido ignorado.")
            return False
        try:
            self._execute(logger)
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
        self.logger.info(self.manual_log)
        try:
            self._run_sync(self.logger, blocking=False)
        except Exception:
            self.logger.exception("Falha na execução manual.")

    def start_monitoring(self):
        if self.is_monitoring:
            return False
        stop_event = threading.Event()
        self._stop_event = stop_event
        self._monitor_thread = threading.Thread(target=self._monitor_main, args=(stop_event,), daemon=True)
        self._monitor_thread.start()
        return True

    def _monitor_main(self, stop_event):
        raise NotImplementedError

    def stop_monitoring(self):
        # Não interrompe uma execução em curso: ela termina e o loop encerra em seguida.
        if self.is_monitoring:
            self.logger.info("Parando: a execução em curso, se houver, termina antes.")
        self._stop_event.set()

    def shutdown(self):
        self.stop_monitoring()
        if self._monitor_thread is not None:
            self._monitor_thread.join(timeout=MONITOR_JOIN_TIMEOUT_SECONDS)

    def configuration_error(self):
        return None

    def status(self, since):
        monitoring = self.is_monitoring
        try:
            config_error = self.configuration_error()
        except Exception as exc:
            config_error = str(exc)
        return {
            "title": self.title,
            "monitoring": monitoring,
            "stopping": monitoring and self._stop_event.is_set(),
            "syncing": self._sync_lock.locked(),
            "poll_interval": self.poll_interval,
            "last_sync": self.last_sync,
            "last_label": self.last_label,
            "running_label": self.running_label,
            "config_error": config_error,
            "logs": self.log_handler.since(since),
        }


class ManutencaoService(BaseFlowService):
    title = "Planilha de Manutenção"
    last_label = "Última sincronização"
    running_label = "sincronizando..."
    manual_log = "Sincronização manual solicitada."
    monitor_started_log = "Manter sincronia ativado."

    def __init__(self, logger, log_handler):
        super().__init__(logger, log_handler, POLL_INTERVAL_SECONDS)

    def _execute(self, logger):
        synchronize(logger)

    def _monitor_main(self, stop_event):
        try:
            start_monitor(self.logger, stop_event, lambda lg: self._run_sync(lg, blocking=True))
        except Exception:
            self.logger.exception("Monitoramento encerrado por erro.")

    def configuration_error(self):
        try:
            validate_config()
        except Exception as exc:
            return str(exc)
        return None


class MedidoresService(BaseFlowService):
    title = "Planilha Medidores"
    last_label = "Última importação"
    running_label = "importando..."
    manual_log = "Importação manual solicitada."
    monitor_started_log = "Importação automática ativada."

    def __init__(self, logger, log_handler):
        super().__init__(logger, log_handler, MEDIDORES_IMPORT_INTERVAL_SECONDS)
        self._last_fingerprint = None

    def _execute(self, logger):
        self._last_fingerprint, _ = importar(logger, self._last_fingerprint)

    def _monitor_main(self, stop_event):
        self.logger.info(
            "Monitorando as 3 origens a cada %s segundos.",
            MEDIDORES_IMPORT_INTERVAL_SECONDS,
        )
        while not stop_event.is_set():
            try:
                self._run_sync(self.logger, blocking=True)
            except Exception:
                self.logger.exception("Falha na importação. Nova tentativa no próximo ciclo.")
            if stop_event.wait(MEDIDORES_IMPORT_INTERVAL_SECONDS):
                break
        self.logger.info("Monitoramento de medidores parado.")

    def configuration_error(self):
        try:
            validate_medidores_config()
        except Exception as exc:
            return str(exc)
        return None


class PanelServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, services, token):
        super().__init__((HOST, 0), PanelRequestHandler)  # porta 0 = livre, escolhida pelo SO
        self.services = services
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

    def _flow_and_action(self):
        parts = [part for part in urlparse(self.path).path.split("/") if part]
        if len(parts) == 3 and parts[0] == "api":
            return parts[1], parts[2]
        return None, None

    def do_GET(self):
        url = urlparse(self.path)
        query = parse_qs(url.query)
        if url.path == "/":
            if not self._authorized(query):
                self._send(403, "Acesso negado.", "text/plain; charset=utf-8")
                return
            self._send(200, PAGE_HTML.replace("__TOKEN__", self.server.token), "text/html; charset=utf-8")
            return
        if not self._authorized(query):
            self._send(403, "{}")
            return
        flow, action = self._flow_and_action()
        if flow is None or action != "status":
            self._send(404, "{}")
            return
        service = self.server.services.get(flow)
        if service is None:
            self._send(404, "{}")
            return
        self.server.close_requested_at = None  # página viva (ex.: recarregada): cancela o encerramento
        try:
            since = int(query.get("since", ["0"])[0])
        except ValueError:
            since = 0
        payload = service.status(since)
        payload["flow"] = flow
        self._send(200, json.dumps(payload))

    def do_POST(self):
        url = urlparse(self.path)
        query = parse_qs(url.query)
        if not self._authorized(query):
            self._send(403, "{}")
            return
        if url.path == "/api/closing":
            self.server.close_requested_at = time.monotonic()
            self._send(200, "{}")
            return
        flow, action = self._flow_and_action()
        if flow is None:
            self._send(404, "{}")
            return
        service = self.server.services.get(flow)
        if service is None:
            self._send(404, "{}")
            return
        if action == "sync":
            service.sync_once()
        elif action == "start":
            if service.start_monitoring():
                service.logger.info(service.monitor_started_log)
        elif action == "stop":
            service.stop_monitoring()
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
            logging.getLogger("excel_sync").info("Janela fechada; encerrando o programa.")
            # Encerra os fluxos em paralelo para não somar os tempos de espera.
            threads = [
                threading.Thread(target=service.shutdown, daemon=True)
                for service in server.services.values()
            ]
            for thread in threads:
                thread.start()
            for thread in threads:
                thread.join()
            server.shutdown()
            return


def build_services(logger):
    """Cria um serviço e um buffer de log independente para cada fluxo."""
    services = {}
    for key, service_class in (("manutencao", ManutencaoService), ("medidores", MedidoresService)):
        log_handler = PanelLogHandler()
        flow_logger = logger.getChild(key)
        for existing in list(flow_logger.handlers):
            if isinstance(existing, PanelLogHandler):
                flow_logger.removeHandler(existing)
        flow_logger.addHandler(log_handler)
        services[key] = service_class(flow_logger, log_handler)
    return services


PAGE_HTML = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel Sync v11</title>
<style>
:root{--bg:#f6f7f9;--fg:#1c2128;--card:#fff;--line:#d8dce2;--accent:#0b6bcb;--ok:#1a7f37;--warn:#9a6700;--err:#cf222e;--mute:#6a737d;--soft:#eef2f7}
@media (prefers-color-scheme:dark){:root{--bg:#14171c;--fg:#e6e9ee;--card:#1d222a;--line:#323a45;--accent:#4c9aff;--ok:#3fb950;--warn:#d29922;--err:#ff7b72;--mute:#8b949e;--soft:#222932}}
*{box-sizing:border-box}
[hidden]{display:none!important}
body{margin:0;padding:12px;background:var(--bg);color:var(--fg);font:14px system-ui,Segoe UI,sans-serif;display:flex;flex-direction:column;height:100vh;gap:10px}
header{display:flex;align-items:center;gap:8px}
header h1{font-size:16px;margin:0}
header p{margin:2px 0 0;color:var(--mute);font-size:12px}
#back{width:34px;height:34px;flex:0 0 34px;padding:0;font-size:18px;line-height:1}
button{flex:1;padding:9px 6px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--fg);font:inherit;cursor:pointer}
button.primary{background:var(--accent);border-color:var(--accent);color:#fff}
button:disabled{opacity:.45;cursor:default}
#menu{display:flex;flex-direction:column;gap:10px;flex:1;overflow:auto}
.card{flex:0 0 auto;text-align:left;padding:14px;border:1px solid var(--line);border-radius:10px;background:var(--card)}
.card:hover{border-color:var(--accent)}
.card .head{display:flex;align-items:center;justify-content:space-between;gap:8px}
.card strong{font-size:15px}
.card .desc{display:block;margin:6px 0 0;color:var(--mute);font-size:12.5px;line-height:1.35}
.badge{font-size:11px;font-weight:600;padding:2px 8px;border-radius:999px;background:var(--soft);color:var(--mute);white-space:nowrap}
.badge.ok{background:rgba(26,127,55,.14);color:var(--ok)}
.badge.bad{background:rgba(207,34,46,.14);color:var(--err)}
.badge.warn{background:rgba(154,103,0,.14);color:var(--warn)}
#panel{display:flex;flex-direction:column;gap:10px;flex:1;min-height:0}
.row{display:flex;gap:6px}
#configError{white-space:pre-wrap;background:rgba(207,34,46,.1);border:1px solid rgba(207,34,46,.35);color:var(--err);border-radius:6px;padding:8px;font-size:12px}
#statusBox{display:flex;flex-direction:column;gap:2px}
#status{display:flex;align-items:center;gap:8px;font-weight:600}
#dot{width:10px;height:10px;border-radius:50%;background:var(--mute);flex:0 0 10px}
#last{color:var(--mute);font-size:12px}
#log{flex:1;min-height:0;overflow:auto;background:var(--card);border:1px solid var(--line);border-radius:6px;padding:8px;font:11.5px ui-monospace,Consolas,monospace}
#log div{white-space:pre-wrap;word-break:break-word;margin-bottom:3px}
.WARNING{color:var(--warn)}.ERROR,.CRITICAL{color:var(--err)}
</style></head><body>
<header>
  <button id="back" title="Voltar ao menu" hidden>&#8592;</button>
  <div><h1 id="title">Excel Sync v11</h1><p id="subtitle">Escolha a planilha</p></div>
</header>
<main id="menu">
  <button class="card" data-flow="manutencao">
    <span class="head"><strong>Planilha de Manutenção</strong><span class="badge" id="badge-manutencao">Parado</span></span>
    <span class="desc">Exporta a planilha Excel local para o Google Sheets.</span>
  </button>
  <button class="card" data-flow="medidores">
    <span class="head"><strong>Planilha Medidores</strong><span class="badge" id="badge-medidores">Parado</span></span>
    <span class="desc">Importa 3 Google Sheets diferentes para uma única planilha Excel local.</span>
  </button>
</main>
<main id="panel" hidden>
  <div class="row">
    <button id="bRun" class="primary">Sincronizar</button>
    <button id="bKeep">Manter sincronia</button>
    <button id="bStop">Parar</button>
  </div>
  <div id="configError" hidden></div>
  <div id="statusBox"><div id="status"><span id="dot"></span><span id="state">Conectando...</span></div><div id="last"></div></div>
  <div id="log"></div>
</main>
<script>
const TOKEN="__TOKEN__";
const $=id=>document.getElementById(id);
const FLOWS={
  manutencao:{title:"Planilha de Manutenção",subtitle:"Excel local → Google Sheets",run:"Sincronizar",keep:"Manter sincronia"},
  medidores:{title:"Planilha Medidores",subtitle:"3 Google Sheets → planilha local",run:"Importar agora",keep:"Importar automaticamente"}
};
let flow=null,lastId=0,panelTimer=null,menuTimer=null;

const api=(path,method)=>fetch(path+(path.includes("?")?"&":"?")+"t="+TOKEN,{method:method||"GET"});
const post=(name,action)=>api("/api/"+name+"/"+action,"POST").then(poll);

function setBadge(name,s){
  const badge=$("badge-"+name);
  if(!badge)return;
  if(s.config_error){badge.textContent="Configurar";badge.className="badge bad";}
  else if(s.stopping){badge.textContent="Parando...";badge.className="badge warn";}
  else if(s.monitoring){badge.textContent="Monitorando";badge.className="badge ok";}
  else if(s.syncing){badge.textContent="Executando...";badge.className="badge warn";}
  else{badge.textContent="Parado";badge.className="badge";}
}

function openFlow(name){
  const cfg=FLOWS[name];
  if(!cfg)return;
  flow=name;lastId=0;
  $("log").innerHTML="";
  $("title").textContent=cfg.title;
  $("subtitle").textContent=cfg.subtitle;
  $("bRun").textContent=cfg.run;
  $("bKeep").textContent=cfg.keep;
  $("menu").hidden=true;$("panel").hidden=false;$("back").hidden=false;
  $("state").textContent="Conectando...";$("dot").style.background="var(--mute)";
  $("last").textContent="";$("configError").hidden=true;
  clearInterval(menuTimer);menuTimer=null;
  poll();
  clearInterval(panelTimer);panelTimer=setInterval(poll,1000);
}

function openMenu(){
  flow=null;lastId=0;
  $("title").textContent="Excel Sync v11";
  $("subtitle").textContent="Escolha a planilha";
  $("panel").hidden=true;$("menu").hidden=false;$("back").hidden=true;
  clearInterval(panelTimer);panelTimer=null;
  pollMenu();
  clearInterval(menuTimer);menuTimer=setInterval(pollMenu,2000);
}

function route(){
  const name=location.hash.replace(/^#/,"");
  if(FLOWS[name])openFlow(name);else openMenu();
}

function pollMenu(){
  for(const name of Object.keys(FLOWS)){
    api("/api/"+name+"/status?since=999999999").then(r=>r.json()).then(s=>{
      if(!flow)setBadge(name,s);
    }).catch(()=>{});
  }
}

function poll(){
  if(!flow)return Promise.resolve();
  const name=flow;
  return api("/api/"+name+"/status?since="+lastId).then(r=>r.json()).then(s=>{
    if(flow!==name)return;
    render(s);
  }).catch(()=>{
    if(flow===name){$("state").textContent="Sem conexão com o programa";$("dot").style.background="var(--err)";}
  });
}

function render(s){
  let text="Parado",color="var(--mute)";
  if(s.config_error){text="Configuração pendente";color="var(--err)";}
  else if(s.stopping){text="Parando...";color="var(--warn)";}
  else if(s.monitoring){text="Monitorando (a cada "+s.poll_interval+"s)";color="var(--ok)";}
  if(s.syncing){text+=" · "+s.running_label;color="var(--accent)";}
  $("state").textContent=text;$("dot").style.background=color;
  $("bRun").disabled=s.syncing||!!s.config_error;
  $("bKeep").disabled=s.monitoring||!!s.config_error;
  $("bStop").disabled=!s.monitoring||s.stopping;
  $("last").textContent=s.last_sync
    ?(s.last_label+": "+(s.last_sync.ok?"OK":"COM ERRO")+" às "+s.last_sync.at)
    :"Nenhuma execução nesta sessão.";
  if(s.config_error){
    $("configError").hidden=false;
    $("configError").textContent="Configuração pendente:\\n"+s.config_error;
  }else{
    $("configError").hidden=true;
  }
  const log=$("log"),atEnd=log.scrollTop+log.clientHeight>=log.scrollHeight-20;
  for(const line of s.logs){
    if(line.id<=lastId)continue;
    const div=document.createElement("div");
    div.className=line.level;
    div.textContent=line.text;
    log.appendChild(div);
    lastId=line.id;
  }
  while(log.children.length>300)log.removeChild(log.firstChild);
  if(atEnd)log.scrollTop=log.scrollHeight;
}

$("bRun").onclick=()=>flow&&post(flow,"sync");
$("bKeep").onclick=()=>flow&&post(flow,"start");
$("bStop").onclick=()=>flow&&post(flow,"stop");
$("back").onclick=()=>{location.hash="";};
document.querySelectorAll(".card").forEach(card=>{
  card.onclick=()=>{location.hash=card.dataset.flow;};
});
window.addEventListener("hashchange",route);
window.addEventListener("pagehide",()=>navigator.sendBeacon("/api/closing?t="+TOKEN));
route();
</script></body></html>
"""


def main():
    logger = setup_logger()
    services = build_services(logger)
    token = secrets.token_urlsafe(16)
    server = PanelServer(services, token)
    url = f"http://{HOST}:{server.server_address[1]}/?t={token}"
    logger.info("Painel Excel Sync em %s (%s).", url.split("?")[0], open_panel_window(url))
    # O token não vai para o arquivo de log; só para o console, caso a janela não abra.
    print(f"Se a janela não abrir, copie este endereço no navegador:\n{url}", flush=True)
    threading.Thread(target=_close_watcher, args=(server,), daemon=True).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("Interrompido pelo teclado; encerrando.")
    finally:
        for service in services.values():
            service.shutdown()
        server.server_close()
