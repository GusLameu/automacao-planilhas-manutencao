import os
from pathlib import Path
from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parent
load_dotenv(PROJECT_DIR / ".env")

_excel = os.getenv("EXCEL_FILE", "").strip().strip('"')
EXCEL_FILE = Path(_excel) if _excel else None
SOURCE_SHEETS = [x.strip() for x in os.getenv("SOURCE_SHEETS", "").split(",") if x.strip()]
GOOGLE_SPREADSHEET_ID = os.getenv("GOOGLE_SPREADSHEET_ID", "").strip()
GOOGLE_WORKSHEET_NAME = os.getenv("GOOGLE_WORKSHEET_NAME", "Todas as demandas").strip()
GOOGLE_CREDENTIALS = Path(os.getenv("GOOGLE_CREDENTIALS", "credentials/service_account.json")).expanduser()
POLL_INTERVAL_SECONDS = float(os.getenv("POLL_INTERVAL_SECONDS", "5"))
DEBOUNCE_SECONDS = float(os.getenv("DEBOUNCE_SECONDS", "3"))
STABILITY_TIMEOUT_SECONDS = int(os.getenv("STABILITY_TIMEOUT_SECONDS", "120"))
STABILITY_INTERVAL_SECONDS = float(os.getenv("STABILITY_INTERVAL_SECONDS", "1"))

def validate_config():
    errors = []
    if EXCEL_FILE is None or not EXCEL_FILE.is_file():
        errors.append(f"Arquivo Excel não encontrado ou sem acesso: {EXCEL_FILE}. Verifique a conexão e as permissões da pasta de rede.")
    if not SOURCE_SHEETS:
        errors.append("SOURCE_SHEETS não contém abas configuradas.")
    if not GOOGLE_SPREADSHEET_ID:
        errors.append("GOOGLE_SPREADSHEET_ID não foi configurado.")
    if not GOOGLE_WORKSHEET_NAME:
        errors.append("GOOGLE_WORKSHEET_NAME não foi configurado.")
    credential_path = GOOGLE_CREDENTIALS if GOOGLE_CREDENTIALS.is_absolute() else PROJECT_DIR / GOOGLE_CREDENTIALS
    if not credential_path.is_file():
        errors.append(f"Credencial Google não encontrada: {credential_path}")
    if errors:
        raise ValueError("\n".join(errors))
    return credential_path
