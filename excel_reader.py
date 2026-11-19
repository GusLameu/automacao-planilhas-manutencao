from pathlib import Path
from datetime import date, datetime, time as datetime_time
from openpyxl import load_workbook

FIRST_COLUMN = 1
LAST_COLUMN = 14  # A:N
COLUMN_COUNT = LAST_COLUMN - FIRST_COLUMN + 1

def normalize(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, (date, datetime_time)):
        return value.isoformat()
    return value

def read_and_combine_excel(file_path, source_sheet_names, logger):
    file_path = Path(file_path)
    if file_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("O arquivo de origem deve ser .xlsx ou .xlsm.")
    wb = load_workbook(file_path, read_only=True, data_only=True, keep_vba=file_path.suffix.lower()==".xlsm")
    try:
        missing = [s for s in source_sheet_names if s not in wb.sheetnames]
        if missing:
            raise ValueError(f"Abas não encontradas no Excel: {missing}")
        all_rows, expected_headers, used, skipped = [], None, [], []
        for name in source_sheet_names:
            ws = wb[name]
            it = ws.iter_rows(min_col=FIRST_COLUMN, max_col=LAST_COLUMN, values_only=True)
            first = next(it, None)
            if first is None or all(v is None or v == "" for v in first):
                skipped.append(name); logger.info("Aba '%s' vazia: ignorada.", name); continue
            headers = [str(v).strip() if v is not None else "" for v in first]
            if len(headers) != COLUMN_COUNT or any(not h for h in headers):
                raise ValueError(f"A aba '{name}' precisa ter cabeçalhos preenchidos em todas as colunas A:N.")
            if len(set(headers)) != len(headers):
                raise ValueError(f"A aba '{name}' possui cabeçalhos duplicados em A:N.")
            if expected_headers is None:
                expected_headers = headers
            elif headers != expected_headers:
                raise ValueError(f"Os cabeçalhos da aba '{name}' diferem dos cabeçalhos da primeira aba. Verifique nomes e ordem em A:N.")
            count = 0
            for row in it:
                vals = list(row[:COLUMN_COUNT])
                if all(v is None or v == "" for v in vals):
                    continue
                all_rows.append([normalize(v) for v in vals])
                count += 1
            if count:
                used.append(name); logger.info("Aba '%s': %d registro(s) lido(s) em A:N.", name, count)
            else:
                skipped.append(name); logger.info("Aba '%s' sem registros: ignorada.", name)
        if not all_rows:
            raise ValueError("Nenhum registro encontrado nas abas configuradas. O destino NÃO será apagado.")
        return [expected_headers] + all_rows, used, skipped
    finally:
        wb.close()
