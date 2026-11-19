from logger import setup_logger
from monitor import start_monitor

if __name__ == "__main__":
    logger = setup_logger()
    logger.info("Iniciando Excel Sync.")
    start_monitor(logger)
