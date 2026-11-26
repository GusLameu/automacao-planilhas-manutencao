import os
import re
from pathlib import Path

from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parent
load_dotenv(PROJECT_DIR / ".env")

# ---------------------------------------------------------------------------
# Fluxo 1: Planilha de Manutenção (Excel local -> Google Sheets)
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Fluxo 2: Planilha Medidores (3 Google Sheets -> Excel local)
# ---------------------------------------------------------------------------
def extract_spreadsheet_id(value):
    """Aceita o ID puro ou a URL completa de uma planilha do Google Sheets."""
    value = (value or "").strip().strip('"').strip("'")
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", value)
    return match.group(1) if match else value


def _project_path(value, default):
    value = (value or "").strip().strip('"')
    path = Path(value).expanduser() if value else default
    return path if path.is_absolute() else PROJECT_DIR / path


MEDIDORES_EXCEL_FILE = _project_path(
    os.getenv("MEDIDORES_EXCEL_FILE", ""),
    Path(r"Z:\00. Equipe GEMED\06. Gustavo Lameu\DashboardAPP\medidados-cogerh\Medidores - 2026.xlsx"),
)
MEDIDORES_WORKSHEET_NAME = os.getenv("MEDIDORES_WORKSHEET_NAME", "Planilha1").strip() or "Planilha1"
MEDIDORES_IMPORT_INTERVAL_SECONDS = float(os.getenv("MEDIDORES_IMPORT_INTERVAL_SECONDS", "60"))

MEDIDORES_HEADERS = [
    "TOMBO",
    "Nº SÉRIE ELETRÔNICA",
    "LOCAL INSTALAÇÃO",
    "BACIA HIDROGRÁFICA",
    "MODELO ELETRÔNICA",
    "TIPO DE MEDIDOR",
]

_DEFAULT_MEDIDORES_SOURCES = [
    {
        "label": "ELETROMAGNÉTICO",
        "spreadsheet_id": "1SYjDRZ-5v0ZKfnPPBEPW6XYWYX6xQhQm",
        "start_row": 3,
        "columns": ["D", "I", "O", "P", "H"],
        "tipo_medidor": "ELETROMAGNÉTICO",
    },
    {
        "label": "ULTRASSÔNICO",
        "spreadsheet_id": "17YaHYvcenWuTLMS-X-EcfjqpsQjNd6GL",
        "start_row": 3,
        "columns": ["D", "I", "Q", "R", "H"],
        "tipo_medidor": "ULTRASSÔNICO",
    },
    {
        "label": "CONDUTOS LIVRES",
        "spreadsheet_id": "1m2OKrCvN19HGqps8qTa1ERsndp3jDQMK",
        "start_row": 2,
        "columns": ["C", "F", "N", "R", "E"],
        "tipo_medidor": "CONDUTOS LIVRES",
    },
]

MEDIDORES_SOURCES = []
for _index, _default in enumerate(_DEFAULT_MEDIDORES_SOURCES, start=1):
    _env_id = extract_spreadsheet_id(os.getenv(f"MEDIDORES_SOURCE_{_index}_ID", ""))
    MEDIDORES_SOURCES.append(
        {
            "label": os.getenv(f"MEDIDORES_SOURCE_{_index}_LABEL", "").strip() or _default["label"],
            "spreadsheet_id": _env_id or _default["spreadsheet_id"],
            "worksheet_name": os.getenv(f"MEDIDORES_SOURCE_{_index}_SHEET", "").strip(),
            "start_row": int(os.getenv(f"MEDIDORES_SOURCE_{_index}_START_ROW", _default["start_row"])),
            "columns": _default["columns"],
            "tipo_medidor": _default["tipo_medidor"],
        }
    )


def _credentials_path():
    return GOOGLE_CREDENTIALS if GOOGLE_CREDENTIALS.is_absolute() else PROJECT_DIR / GOOGLE_CREDENTIALS


def validate_config():
    """Valida o fluxo Planilha de Manutenção e devolve o caminho da credencial."""
    errors = []
    if EXCEL_FILE is None or not EXCEL_FILE.is_file():
        errors.append(f"Arquivo Excel não encontrado ou sem acesso: {EXCEL_FILE}. Verifique a conexão e as permissões da pasta de rede.")
    if not SOURCE_SHEETS:
        errors.append("SOURCE_SHEETS não contém abas configuradas.")
    if not GOOGLE_SPREADSHEET_ID:
        errors.append("GOOGLE_SPREADSHEET_ID não foi configurado.")
    if not GOOGLE_WORKSHEET_NAME:
        errors.append("GOOGLE_WORKSHEET_NAME não foi configurado.")
    credential_path = _credentials_path()
    if not credential_path.is_file():
        errors.append(f"Credencial Google não encontrada: {credential_path}")
    if errors:
        raise ValueError("\n".join(errors))
    return credential_path


def validate_medidores_config():
    """Valida o fluxo Planilha Medidores e devolve o caminho da credencial."""
    errors = []
    missing = [
        f"MEDIDORES_SOURCE_{index}_ID"
        for index, source in enumerate(MEDIDORES_SOURCES, start=1)
        if not source["spreadsheet_id"]
    ]
    if missing:
        errors.append(
            "Origem(ns) do Google Sheets não configurada(s): " + ", ".join(missing) + ". "
            "Preencha o ID ou a URL de cada planilha no .env e reinicie o programa."
        )
    if MEDIDORES_EXCEL_FILE.suffix.lower() not in {".xlsx", ".xlsm"}:
        errors.append(f"MEDIDORES_EXCEL_FILE deve ser .xlsx ou .xlsm: {MEDIDORES_EXCEL_FILE}")
    if MEDIDORES_IMPORT_INTERVAL_SECONDS <= 0:
        errors.append("MEDIDORES_IMPORT_INTERVAL_SECONDS deve ser maior que zero.")
    if (
        len(MEDIDORES_WORKSHEET_NAME) > 31
        or any(char in "[]:*?/\\" for char in MEDIDORES_WORKSHEET_NAME)
    ):
        errors.append(
            f"MEDIDORES_WORKSHEET_NAME inválido para o Excel: {MEDIDORES_WORKSHEET_NAME!r}. "
            "Use até 31 caracteres e evite os caracteres []:*?/\\."
        )
    credential_path = _credentials_path()
    if not credential_path.is_file():
        errors.append(f"Credencial Google não encontrada: {credential_path}")
    if errors:
        raise ValueError("\n".join(errors))
    return credential_path
