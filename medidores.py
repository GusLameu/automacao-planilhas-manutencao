"""Importa 3 planilhas do Google Sheets para uma única planilha Excel local.

O fluxo lê todas as origens configuradas, valida se os cabeçalhos são iguais,
empilha os registros e substitui o conteúdo da aba de destino no Excel.
"""
import hashlib
import io
import json
import os
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import gspread
from google.auth.transport.requests import AuthorizedSession
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import column_index_from_string, get_column_letter

from config import (
    MEDIDORES_EXCEL_FILE,
    MEDIDORES_HEADERS,
    MEDIDORES_IMPORT_INTERVAL_SECONDS,
    MEDIDORES_SOURCES,
    MEDIDORES_WORKSHEET_NAME,
    validate_medidores_config,
)
from google_sheets import open_client

SHEETS_EPOCH = datetime(1899, 12, 30)
DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S")
INVALID_SHEET_CHARS = set('[]:*?/\\')


def _is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")


def _parse_date_text(text):
    text = str(text).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _serial_to_value(serial, has_time):
    moment = SHEETS_EPOCH + timedelta(days=float(serial))
    return moment if has_time else moment.date()


def _convert_cell(formatted, raw):
    """Converte o valor bruto do Sheets preservando números e datas.

    O Sheets devolve datas como número de série quando usamos
    UNFORMATTED_VALUE; o valor formatado serve para identificar a data.
    """
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        parsed = _parse_date_text(formatted) if isinstance(formatted, str) else None
        if parsed is not None:
            has_time = ":" in formatted or float(raw) != int(raw)
            return _serial_to_value(raw, has_time)
        return raw
    if raw is None:
        return ""
    text = str(formatted if formatted not in (None, "") else raw).strip()
    return text


def _pad_matrix(matrix, rows, cols):
    return [
        [
            matrix[row][col] if row < len(matrix) and col < len(matrix[row]) else ""
            for col in range(cols)
        ]
        for row in range(rows)
    ]


def _read_worksheet(worksheet):
    """Lê a aba em dois formatos e devolve as linhas já convertidas."""
    formatted = worksheet.get_all_values(value_render_option="FORMATTED_VALUE", pad_values=True)
    raw = worksheet.get_all_values(
        value_render_option="UNFORMATTED_VALUE",
        date_time_render_option="SERIAL_NUMBER",
        pad_values=True,
    )
    rows = max(len(formatted), len(raw))
    cols = max([len(row) for row in formatted] + [len(row) for row in raw] + [0])
    if rows == 0 or cols == 0:
        return []
    formatted = _pad_matrix(formatted, rows, cols)
    raw = _pad_matrix(raw, rows, cols)
    return [
        [_convert_cell(formatted_value, raw_value) for formatted_value, raw_value in zip(formatted_row, raw_row)]
        for formatted_row, raw_row in zip(formatted, raw)
    ]


def _read_drive_xlsx(client, source, logger):
    """Baixa um arquivo .xlsx do Google Drive via Drive API e lê as linhas."""
    label = source["label"]
    file_id = source["spreadsheet_id"]
    session = AuthorizedSession(client.http_client.auth)
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
    response = session.get(url, timeout=60)
    if response.status_code != 200:
        raise ValueError(
            f"Não foi possível baixar o arquivo Excel de '{label}' via Drive API "
            f"(HTTP {response.status_code}): {response.text}"
        )

    workbook = load_workbook(io.BytesIO(response.content), data_only=True, read_only=True)
    try:
        if source["worksheet_name"]:
            if source["worksheet_name"] not in workbook.sheetnames:
                raise ValueError(
                    f"A aba '{source['worksheet_name']}' não existe no arquivo de '{label}'."
                )
            worksheet = workbook[source["worksheet_name"]]
        else:
            worksheet = workbook.active

        rows = [
            ["" if cell is None else cell for cell in row]
            for row in worksheet.iter_rows(values_only=True)
        ]
        return rows, label, worksheet.title
    finally:
        workbook.close()


def _read_source(client, source, logger):
    label = source["label"]
    try:
        spreadsheet = client.open_by_key(source["spreadsheet_id"])
    except gspread.SpreadsheetNotFound as exc:
        raise ValueError(
            f"Planilha do Google não encontrada para '{label}'. "
            "Verifique o ID/URL e compartilhe a planilha com a conta de serviço."
        ) from exc
    except gspread.exceptions.APIError as exc:
        detail = str(exc)
        if "Office file" in detail or "not supported for this document" in detail:
            logger.info("Origem '%s' é um arquivo Excel no Drive; baixando via Drive API.", label)
            rows, title, ws_title = _read_drive_xlsx(client, source, logger)
        else:
            raise ValueError(
                f"Não foi possível abrir a planilha de '{label}'. "
                "Verifique o ID/URL e o compartilhamento com a conta de serviço. "
                f"Detalhe: {exc}"
            ) from exc
    else:
        title = getattr(spreadsheet, "title", label)
        if source["worksheet_name"]:
            try:
                worksheet = spreadsheet.worksheet(source["worksheet_name"])
            except gspread.WorksheetNotFound as exc:
                raise ValueError(
                    f"A aba '{source['worksheet_name']}' não existe em '{title}' ({label})."
                ) from exc
        else:
            worksheet = spreadsheet.sheet1
        ws_title = worksheet.title
        rows = _read_worksheet(worksheet)

    if not rows:
        return [], [], title, ws_title

    if source.get("columns"):
        start_row = max(1, int(source.get("start_row", 1)))
        col_indices = [column_index_from_string(col) - 1 for col in source["columns"]]
        tipo_medidor = source.get("tipo_medidor", label)
        body = []
        for row in rows[start_row - 1 :]:
            values = [row[idx] if idx < len(row) else "" for idx in col_indices]
            if all(_is_blank(value) for value in values):
                continue
            values.append(tipo_medidor)
            body.append(values)
        return list(MEDIDORES_HEADERS), body, title, ws_title

    first_row = rows[0]
    if all(_is_blank(value) for value in first_row):
        raise ValueError(
            f"A primeira linha da origem '{label}' ({title} / {worksheet.title}) está vazia. "
            "Preencha os cabeçalhos antes da importação."
        )

    headers = [str(value).strip() for value in first_row]
    empty_columns = [get_column_letter(index) for index, header in enumerate(headers, start=1) if not header]
    if empty_columns:
        raise ValueError(
            f"A origem '{label}' ({title} / {worksheet.title}) tem cabeçalho vazio na(s) coluna(s) "
            f"{', '.join(empty_columns)}."
        )
    duplicates = sorted({header for header in headers if headers.count(header) > 1})
    if duplicates:
        raise ValueError(
            f"A origem '{label}' ({title} / {worksheet.title}) tem cabeçalhos duplicados: "
            f"{', '.join(duplicates)}."
        )

    body = []
    for row in rows[1:]:
        values = list(row[: len(headers)])
        values += [""] * (len(headers) - len(values))
        if all(_is_blank(value) for value in values):
            continue
        body.append(values)
    return headers, body, title, worksheet.title


def read_all_sources(credential_path, logger, sources=None):
    """Lê e empilha as origens. Devolve (valores, usadas, ignoradas)."""
    sources = sources or MEDIDORES_SOURCES
    client = open_client(credential_path)
    expected_headers = None
    combined = []
    used = []
    skipped = []

    for source in sources:
        headers, body, title, worksheet_title = _read_source(client, source, logger)
        if not headers:
            skipped.append(source["label"])
            logger.info("Origem '%s' (%s) está vazia: ignorada.", source["label"], title)
            continue
        if expected_headers is None:
            expected_headers = headers
        elif headers != expected_headers:
            raise ValueError(
                f"Os cabeçalhos da origem '{source['label']}' ({title} / {worksheet_title}) "
                "diferem dos da primeira origem.\n"
                f"Esperado: {expected_headers}\n"
                f"Encontrado: {headers}"
            )
        combined.extend(body)
        used.append(source["label"])
        logger.info(
            "Origem '%s' (%s / %s): %d registro(s) em %d coluna(s).",
            source["label"],
            title,
            worksheet_title,
            len(body),
            len(headers),
        )

    if expected_headers is None:
        raise ValueError("Nenhuma das 3 origens possui dados. A planilha local NÃO foi alterada.")
    if not combined:
        raise ValueError("As origens possuem apenas cabeçalhos, sem registros. A planilha local NÃO foi alterada.")
    return [expected_headers] + combined, used, skipped


def fingerprint(values):
    payload = json.dumps(values, ensure_ascii=False, default=str, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _validate_sheet_name(name):
    if (
        not name
        or len(name) > 31
        or any(char in INVALID_SHEET_CHARS for char in name)
    ):
        raise ValueError(
            f"Nome de aba inválido para o Excel: {name!r}. "
            "Use até 31 caracteres e evite os caracteres []:*?/\\."
        )


def write_values(values, logger, path=None, worksheet_name=None):
    """Substitui a aba de destino no Excel local, preservando as outras abas."""
    path = Path(path or MEDIDORES_EXCEL_FILE)
    worksheet_name = worksheet_name or MEDIDORES_WORKSHEET_NAME
    _validate_sheet_name(worksheet_name)

    if not values or not values[0]:
        raise ValueError("Nenhum dado para gravar na planilha local.")
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError(f"O arquivo de destino deve ser .xlsx ou .xlsm: {path}")

    path.parent.mkdir(parents=True, exist_ok=True)
    created = not path.is_file()
    if created:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = worksheet_name
    else:
        try:
            workbook = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
        except PermissionError as exc:
            raise ValueError(
                f"Não foi possível abrir '{path}'. Feche o arquivo no Excel e tente novamente."
            ) from exc
        if worksheet_name in workbook.sheetnames:
            worksheet = workbook[worksheet_name]
        else:
            worksheet = workbook.create_sheet(worksheet_name)
            logger.info("Aba '%s' criada em '%s'.", worksheet_name, path)

    # Preserva o cabeçalho da linha 1 quando a aba já existe; sobrescreve a partir da linha 2.
    if created:
        rows_to_write = values
        start_row = 1
    else:
        if worksheet.max_row and worksheet.max_row >= 2:
            worksheet.delete_rows(2, worksheet.max_row - 1)
        rows_to_write = values[1:]
        start_row = 2

    for row_index, row in enumerate(rows_to_write, start=start_row):
        for column_index, value in enumerate(row, start=1):
            cell = worksheet.cell(row=row_index, column=column_index)
            cell.value = value
            # Evita que textos iniciados por "=" virem fórmulas no Excel.
            if isinstance(value, str) and value.startswith("="):
                cell.data_type = "s"
            if isinstance(value, datetime):
                cell.number_format = "DD/MM/YYYY HH:MM:SS"
            elif isinstance(value, date):
                cell.number_format = "DD/MM/YYYY"

    if created:
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        worksheet.freeze_panes = "A2"

    # Grava primeiro em arquivo temporário para não corromper o destino em caso de erro.
    temporary = path.with_name(f".{path.name}.tmp")
    try:
        workbook.save(temporary)
        os.replace(temporary, path)
    except OSError as exc:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
        if isinstance(exc, PermissionError) or getattr(exc, "winerror", None) == 32:
            raise ValueError(
                f"Não foi possível gravar '{path}'. Feche o arquivo no Excel e tente novamente."
            ) from exc
        raise

    logger.info(
        "Planilha local '%s' atualizada: %d linha(s) e %d coluna(s)%s.",
        path,
        len(values),
        len(values[0]),
        " (arquivo criado)" if created else "",
    )


def importar(logger, last_fingerprint=None):
    """Executa a importação. Devolve (fingerprint, houve_gravacao)."""
    credential_path = validate_medidores_config()
    values, used, skipped = read_all_sources(credential_path, logger)
    current = fingerprint(values)

    if current == last_fingerprint and MEDIDORES_EXCEL_FILE.is_file():
        logger.debug("Sem alterações nas origens; '%s' não foi reescrita.", MEDIDORES_EXCEL_FILE)
        return current, False

    write_values(values, logger)
    logger.info(
        "Concluído: %d registro(s) importado(s) de %d origem(ns); origens utilizadas=%s; ignoradas=%s.",
        len(values) - 1,
        len(used),
        used,
        skipped,
    )
    return current, True


def run_console(logger):
    """Importação periódica para uso via `python main.py --console --medidores`."""
    state = {"fingerprint": None}
    logger.info(
        "Importação de medidores a cada %s segundos. Ctrl+C para parar.",
        MEDIDORES_IMPORT_INTERVAL_SECONDS,
    )
    try:
        while True:
            try:
                state["fingerprint"], _ = importar(logger, state["fingerprint"])
            except Exception:
                logger.exception("Falha na importação de medidores.")
            time.sleep(MEDIDORES_IMPORT_INTERVAL_SECONDS)
    except KeyboardInterrupt:
        logger.info("Interrompido pelo teclado; encerrando.")
