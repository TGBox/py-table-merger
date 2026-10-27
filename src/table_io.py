"""Gemeinsame Funktionen zum Einlesen von Datenbank-Exporten und zum formatierten Excel-Export."""
import os

import pandas as pd

# Nur diese Werte gelten als "leer". Bewusst NICHT die Pandas-Standardliste,
# sonst würden z. B. Einträge wie "NA" oder "null" ebenfalls verschwinden.
NA_VALUES = ["NULL", ""]
# Reihenfolge der Encoding-Versuche (latin-1 kann jede Bytefolge lesen = letzter Fallback).
CSV_ENCODINGS = ["utf-8-sig", "cp1252", "latin-1"]
# Maximale Spaltenbreite im Excel-Export (Zeichen); längere Inhalte werden umgebrochen.
MAX_EXCEL_COL_WIDTH = 80
# Excel erlaubt maximal 32.767 Zeichen pro Zelle.
MAX_EXCEL_CELL_LEN = 32767
TRUNCATION_NOTE = " […gekürzt]"


def is_text_column(series: pd.Series) -> bool:
    """True für Text-Spalten, egal ob object- oder (ab pandas 3) str-Datentyp."""
    return pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series)


def detect_encoding(filepath: str) -> str:
    with open(filepath, "rb") as f:
        raw = f.read()
    for encoding in CSV_ENCODINGS:
        try:
            raw.decode(encoding)
            return encoding
        except UnicodeDecodeError:
            continue
    return CSV_ENCODINGS[-1]


def read_table(filepath: str) -> pd.DataFrame:
    """Liest CSV/Excel ein. Alle Werte bleiben Text (dtype=str), damit führende
    Nullen (PLZ, Telefon, IDs) erhalten bleiben und Schlüssel immer vergleichbar sind."""
    ext = os.path.splitext(filepath)[1].lower()
    read_opts = dict(dtype=str, keep_default_na=False, na_values=NA_VALUES)
    if ext == ".csv":
        encoding = detect_encoding(filepath)
        for sep in [";", ",", "\t"]:
            try:
                df = pd.read_csv(filepath, sep=sep, encoding=encoding, **read_opts)
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
        return pd.read_csv(filepath, encoding=encoding, **read_opts)
    elif ext in [".xlsx", ".xls"]:
        return pd.read_excel(filepath, **read_opts)
    raise ValueError("Ungültiges Dateiformat.")


def _truncate_long_text(df: pd.DataFrame) -> pd.DataFrame:
    """Kürzt Texte, die länger sind, als eine Excel-Zelle fassen kann."""
    limit = MAX_EXCEL_CELL_LEN - len(TRUNCATION_NOTE)
    df = df.copy()
    for col in df.columns:
        series = df[col]
        if not is_text_column(series):
            continue
        too_long = series.map(lambda v: isinstance(v, str) and len(v) > MAX_EXCEL_CELL_LEN)
        if too_long.any():
            df.loc[too_long, col] = series[too_long].str[:limit] + TRUNCATION_NOTE
    return df


def _display_len(value) -> int:
    if isinstance(value, str):
        # Bei mehrzeiligen Texten zählt die längste Zeile
        return max((len(line) for line in value.splitlines()), default=0)
    if hasattr(value, "strftime"):
        return 16 if getattr(value, "hour", 0) or getattr(value, "minute", 0) else 10
    return len(str(value))


def write_excel_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str,
                      autofilter: bool = True) -> None:
    """Schreibt ein DataFrame als formatiertes Blatt (xlsxwriter):
    fette Überschriften, passende Spaltenbreiten, Umbruch bei langen Texten,
    keine "Zahl als Text"-Warnungen, optional Autofilter + fixierte Kopfzeile."""
    df = _truncate_long_text(df)
    df.to_excel(writer, index=False, sheet_name=sheet_name)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_format = workbook.add_format({'bold': True, 'text_wrap': False, 'valign': 'top'})
    wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})

    # Werte werden als Text eingelesen (führende Nullen!) -> grüne Excel-Dreiecke unterdrücken
    worksheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})

    for col_num, col_name in enumerate(df.columns.values):
        worksheet.write(0, col_num, col_name, header_format)

        max_len = len(str(col_name))
        column_data = df.iloc[:, col_num].dropna()
        if not column_data.empty:
            max_len = max(max_len, int(column_data.map(_display_len).max()))

        col_width = max(max_len + 3, 10)
        if col_width > MAX_EXCEL_COL_WIDTH:
            worksheet.set_column(col_num, col_num, MAX_EXCEL_COL_WIDTH, wrap_format)
        else:
            worksheet.set_column(col_num, col_num, col_width)

    if autofilter and len(df.columns):
        worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
        worksheet.freeze_panes(1, 0)
