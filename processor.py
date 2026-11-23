from config import EXCEL_FILE, SOURCE_SHEETS, GOOGLE_SPREADSHEET_ID, GOOGLE_WORKSHEET_NAME, validate_config
from excel_reader import read_and_combine_excel
from google_sheets import replace_worksheet_values

def synchronize(logger):
    credential_path = validate_config()
    logger.info("Lendo arquivo: %s", EXCEL_FILE)
    values, used, skipped = read_and_combine_excel(EXCEL_FILE, SOURCE_SHEETS, logger)
    logger.info("Abas configuradas: %s", SOURCE_SHEETS)
    logger.info("Abas efetivamente utilizadas: %s", used)
    logger.info("Abas ignoradas: %s", skipped)
    # Só limpa o destino depois de todas as abas terem sido lidas e validadas.
    replace_worksheet_values(credential_path, GOOGLE_SPREADSHEET_ID, GOOGLE_WORKSHEET_NAME, values, logger)
    logger.info("Concluído: %d registros enviados; abas utilizadas=%s; ignoradas=%s", len(values)-1, used, skipped)
