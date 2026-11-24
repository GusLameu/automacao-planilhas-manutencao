import sys

if __name__ == "__main__":
    if "--console" in sys.argv:
        from logger import setup_logger
        from monitor import start_monitor

        logger = setup_logger()
        logger.info("Iniciando Excel Sync.")
        start_monitor(logger)
    else:
        # Padrão: abre a janela com Sincronizar / Manter sincronia / Parar.
        from sync_panel import main as run_panel

        run_panel()
