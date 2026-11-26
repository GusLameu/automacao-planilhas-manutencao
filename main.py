import sys

if __name__ == "__main__":
    if "--console" in sys.argv:
        from logger import setup_logger

        logger = setup_logger()
        if "--medidores" in sys.argv:
            from medidores import run_console

            logger.info("Iniciando Excel Sync (Planilha Medidores) em modo console.")
            run_console(logger)
        else:
            from monitor import start_monitor

            logger.info("Iniciando Excel Sync (Planilha de Manutenção) em modo console.")
            start_monitor(logger)
    else:
        # Padrão: abre o menu com Planilha de Manutenção / Planilha Medidores.
        from sync_panel import main as run_panel

        run_panel()
