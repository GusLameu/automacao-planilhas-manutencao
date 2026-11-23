import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def replace_worksheet_values(credentials_path, spreadsheet_id, worksheet_name, values, logger):
    creds = Credentials.from_service_account_file(str(credentials_path), scopes=SCOPES)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        raise ValueError(f"A aba de destino '{worksheet_name}' não existe. Crie-a antes da execução.")
    rows = len(values)
    cols = max(len(row) for row in values)
    worksheet.clear()
    worksheet.update(range_name="A1", values=values, value_input_option="USER_ENTERED")
    worksheet.format(["B:B", "J:J"], {"numberFormat": {"type": "DATE", "pattern": "dd/MM/yyyy"}})
    logger.info("Destino '%s' atualizado com %d linhas e %d colunas.", worksheet_name, rows, cols)
