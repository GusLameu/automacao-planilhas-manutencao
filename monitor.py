import time
from pathlib import Path
from config import EXCEL_FILE, POLL_INTERVAL_SECONDS, DEBOUNCE_SECONDS, STABILITY_TIMEOUT_SECONDS, STABILITY_INTERVAL_SECONDS
from processor import synchronize

def signature(path):
    st = Path(path).stat()
    return st.st_size, st.st_mtime_ns

def wait_stable(path):
    deadline = time.monotonic() + STABILITY_TIMEOUT_SECONDS
    previous = None
    stable = 0
    while time.monotonic() < deadline:
        try:
            current = signature(path)
        except OSError:
            time.sleep(STABILITY_INTERVAL_SECONDS); continue
        if current == previous and current[0] > 0:
            stable += 1
            if stable >= 2:
                return current
        else:
            stable = 0
        previous = current
        time.sleep(STABILITY_INTERVAL_SECONDS)
    raise TimeoutError("O arquivo Excel não estabilizou no tempo limite.")

def start_monitor(logger):
    if EXCEL_FILE is None or not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"Arquivo Excel indisponível: {EXCEL_FILE}. Verifique acesso à rede.")
    logger.info("Monitorando arquivo de rede: %s", EXCEL_FILE)
    logger.info("Consulta periódica a cada %s segundos.", POLL_INTERVAL_SECONDS)
    try:
        wait_stable(EXCEL_FILE)
        synchronize(logger)
    except Exception:
        logger.exception("Falha na sincronização inicial.")
    try:
        last = signature(EXCEL_FILE)
    except OSError:
        last = None
    pending_since = None
    while True:
        time.sleep(POLL_INTERVAL_SECONDS)
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
            stable = wait_stable(EXCEL_FILE)
            if stable != last:
                last = stable
                pending_since = time.monotonic()
                continue
            synchronize(logger)
            last = signature(EXCEL_FILE)
            pending_since = None
        except Exception:
            logger.exception("Falha na sincronização após alteração. Nova tentativa será feita após novo evento.")
            pending_since = None
