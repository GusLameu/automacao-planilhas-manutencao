import threading
import time
from pathlib import Path
from config import EXCEL_FILE, POLL_INTERVAL_SECONDS, DEBOUNCE_SECONDS, STABILITY_TIMEOUT_SECONDS, STABILITY_INTERVAL_SECONDS
from processor import synchronize

class MonitorStopped(Exception):
    """Levantada quando o monitoramento é interrompido por stop_event."""

def signature(path):
    st = Path(path).stat()
    return st.st_size, st.st_mtime_ns

def wait_stable(path, stop_event):
    deadline = time.monotonic() + STABILITY_TIMEOUT_SECONDS
    previous = None
    stable = 0
    while time.monotonic() < deadline:
        if stop_event.is_set():
            raise MonitorStopped()
        try:
            current = signature(path)
        except OSError:
            stop_event.wait(STABILITY_INTERVAL_SECONDS); continue
        if current == previous and current[0] > 0:
            stable += 1
            if stable >= 2:
                return current
        else:
            stable = 0
        previous = current
        stop_event.wait(STABILITY_INTERVAL_SECONDS)
    raise TimeoutError("O arquivo Excel não estabilizou no tempo limite.")

def start_monitor(logger, stop_event=None, sync_func=synchronize):
    # stop_event=None mantém o comportamento original (loop infinito, usado por main.py).
    if stop_event is None:
        stop_event = threading.Event()
    if EXCEL_FILE is None or not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"Arquivo Excel indisponível: {EXCEL_FILE}. Verifique acesso à rede.")
    logger.info("Monitorando arquivo de rede: %s", EXCEL_FILE)
    logger.info("Consulta periódica a cada %s segundos.", POLL_INTERVAL_SECONDS)
    try:
        wait_stable(EXCEL_FILE, stop_event)
        sync_func(logger)
    except MonitorStopped:
        logger.info("Monitoramento parado.")
        return
    except Exception:
        logger.exception("Falha na sincronização inicial.")
    try:
        last = signature(EXCEL_FILE)
    except OSError:
        last = None
    pending_since = None
    while not stop_event.is_set():
        if stop_event.wait(POLL_INTERVAL_SECONDS):
            break
        try:
            current = signature(EXCEL_FILE)
        except OSError:
            logger.warning("Arquivo de rede indisponível; aguardando reconexão.")
            continue
        if last is None:
            last = current
            continue
        if current != last:
            last = current
            pending_since = time.monotonic()
            logger.info("Alteração detectada no Excel.")
        if pending_since is None or time.monotonic() - pending_since < DEBOUNCE_SECONDS:
            continue
        try:
            stable = wait_stable(EXCEL_FILE, stop_event)
            if stable != last:
                last = stable
                pending_since = time.monotonic()
                continue
            sync_func(logger)
            last = signature(EXCEL_FILE)
            pending_since = None
        except MonitorStopped:
            break
        except Exception:
            logger.exception("Falha na sincronização após alteração. Nova tentativa será feita após novo evento.")
            pending_since = None
    logger.info("Monitoramento parado.")