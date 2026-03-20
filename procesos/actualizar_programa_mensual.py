from __future__ import annotations

from pathlib import Path
from typing import Dict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.comments import Comment

from utils.excel_utils import (
    read_excel_file,
    rename_columns_by_alias,
    validate_required_columns,
)
from utils.texto_utils import normalize_column_name, normalize_key, safe_str

TURNO_ALIASES = {
    "PUNTO": ["PUNTO", "PUNTO_MONITOREO", "ESTACION", "PTO"],
    "TIPO": ["TIPO", "MATRIZ", "TIPO_MATRIZ"],
    "DIA": ["DIA", "DIA_MES"],
    "FECHA": ["FECHA", "FECHA_MONITOREO", "FECHA_MUESTREO"],
    "ESTADO": ["ESTADO", "RESULTADO", "STATUS"],
    "COMENTARIO": ["COMENTARIO", "OBSERVACION", "OBSERVACIONES", "DETALLE"],
}

MENSUAL_PUNTO_HEADERS = {"PUNTO", "PUNTO_MONITOREO", "ESTACION", "PTO"}
MENSUAL_TIPO_HEADERS = {"TIPO", "MATRIZ", "TIPO_MATRIZ"}


def _safe_sheet_limits(ws) -> tuple[int, int]:
    max_row = ws.max_row if isinstance(ws.max_row, int) and ws.max_row >= 1 else 1
    max_col = ws.max_column if isinstance(ws.max_column, int) and ws.max_column >= 1 else 1
    return max_row, max_col


def _extract_day(value: object) -> int | None:
    if isinstance(value, int):
        return value if 1 <= value <= 31 else None
    if isinstance(value, float) and value.is_integer():
        day = int(value)
        return day if 1 <= day <= 31 else None
    if hasattr(value, "day"):
        try:
            day = int(value.day)
            return day if 1 <= day <= 31 else None
        except Exception:
            return None
    value_str = safe_str(value)
    if value_str.isdigit():
        day = int(value_str)
        return day if 1 <= day <= 31 else None
    return None


def _cargar_turno(path_programa_turno: Path | str) -> pd.DataFrame:
    df = read_excel_file(path_programa_turno)
    df = rename_columns_by_alias(df, TURNO_ALIASES)
    validate_required_columns(df, ["PUNTO", "TIPO", "ESTADO"], "programa turno")

    if "DIA" not in df.columns:
        validate_required_columns(
            df, ["FECHA"], "programa turno (falta DIA y se requiere FECHA)"
        )
        df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce", dayfirst=True)
        df["DIA"] = df["FECHA"].dt.day

    if "COMENTARIO" not in df.columns:
        df["COMENTARIO"] = ""

    df["DIA"] = pd.to_numeric(df["DIA"], errors="coerce").astype("Int64")
    df = df[df["DIA"].notna()].copy()

    df["PUNTO"] = df["PUNTO"].map(safe_str)
    df["TIPO"] = df["TIPO"].map(safe_str)
    df["ESTADO"] = df["ESTADO"].map(safe_str)
    df["COMENTARIO"] = df["COMENTARIO"].map(safe_str)
    return df


def _find_header_columns(ws) -> tuple[int, int, Dict[int, int], int, int, int]:
    max_row, max_col = _safe_sheet_limits(ws)
    punto_headers = {normalize_column_name(v) for v in MENSUAL_PUNTO_HEADERS}
    tipo_headers = {normalize_column_name(v) for v in MENSUAL_TIPO_HEADERS}

    best = None
    best_score = -1
    header_scan_rows = min(20, max_row)

    for header_row in range(1, header_scan_rows + 1):
        point_col = None
        tipo_col = None
        day_to_col: Dict[int, int] = {}
        comentario_actual_col = None

        for col in range(1, max_col + 1):
            value = ws.cell(row=header_row, column=col).value
            normalized = normalize_column_name(value)

            if normalized in punto_headers:
                point_col = col
            elif normalized in tipo_headers:
                tipo_col = col
            elif normalized == "COMENTARIO_ACTUAL":
                comentario_actual_col = col
            else:
                day = _extract_day(value)
                if day is not None:
                    day_to_col[day] = col

        score = int(point_col is not None) + int(tipo_col is not None) + len(day_to_col)
        if score > best_score:
            best_score = score
            best = (header_row, point_col, tipo_col, day_to_col, comentario_actual_col)

    if best is None:
        raise ValueError("No fue posible detectar encabezados del programa mensual.")

    header_row, point_col, tipo_col, day_to_col, comentario_actual_col = best

    if point_col is None or tipo_col is None:
        raise ValueError(
            "No se encontraron columnas de PUNTO y/o MATRIZ/TIPO en el programa mensual."
        )
    if not day_to_col:
        raise ValueError("No se encontraron columnas de dia (1..31) en el programa mensual.")

    if comentario_actual_col is None:
        comentario_actual_col = max_col + 1
        ws.cell(row=header_row, column=comentario_actual_col, value="COMENTARIO_ACTUAL")

    return point_col, tipo_col, day_to_col, comentario_actual_col, header_row, max_row


def actualizar_programa_mensual(
    path_programa_turno: Path | str,
    path_programa_mensual: Path | str,
    output_path_programa_mensual: Path | str,
) -> pd.DataFrame:
    print("[3/5] Actualizando programa mensual...")
    turno_df = _cargar_turno(path_programa_turno)

    wb = load_workbook(Path(path_programa_mensual))
    ws = wb.active

    (
        point_col,
        tipo_col,
        day_to_col,
        comentario_actual_col,
        header_row,
        max_row,
    ) = _find_header_columns(ws)

    row_lookup: Dict[tuple[str, str], int] = {}
    for row in range(header_row + 1, max_row + 1):
        point = ws.cell(row=row, column=point_col).value
        tipo = ws.cell(row=row, column=tipo_col).value
        key = normalize_key(point, tipo)
        row_lookup[key] = row

    aplicados = []
    total = len(turno_df)
    updated = 0
    skipped = 0

    for _, record in turno_df.iterrows():
        day = int(record["DIA"])
        key = normalize_key(record["PUNTO"], record["TIPO"])
        row = row_lookup.get(key)
        col = day_to_col.get(day)
        if row is None or col is None:
            skipped += 1
            continue

        estado = safe_str(record["ESTADO"])
        comentario = safe_str(record["COMENTARIO"])

        cell = ws.cell(row=row, column=col)
        cell.value = estado

        if comentario:
            cell.comment = Comment(comentario, "automatizador_cierre_ot")
            ws.cell(row=row, column=comentario_actual_col, value=comentario)

        aplicados.append(
            {
                "PUNTO": safe_str(record["PUNTO"]),
                "TIPO": safe_str(record["TIPO"]),
                "DIA": day,
                "ESTADO": estado,
                "COMENTARIO": comentario,
            }
        )
        updated += 1

    output_path = Path(output_path_programa_mensual)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    print(
        f"[3/5] Registros turno: {total} | actualizados en mensual: {updated} | omitidos: {skipped}"
    )
    return pd.DataFrame(
        aplicados, columns=["PUNTO", "TIPO", "DIA", "ESTADO", "COMENTARIO"]
    )
