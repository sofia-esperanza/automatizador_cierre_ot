from __future__ import annotations

import datetime as dt
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
from openpyxl import load_workbook

from procesos.clasificar_cierre_ot import cargar_diccionarios
from utils.excel_utils import save_dataframe_to_excel
from utils.texto_utils import (
    first_non_empty,
    normalize_column_name,
    normalize_key,
    normalize_text,
    safe_str,
)

SEMANAL_ESTADO_COL_HINTS = {
    "R",
    "NR",
    "RR",
    "S",
    "REALIZADO",
    "NO_REALIZADO",
}

MENSUAL_CODIGO_HEADERS_PRIORITY = [
    "CODIGO",
    "CODIGO_MST",
    "CODIGO_PUNTO",
]
MENSUAL_MATRIZ_TERRENO_HEADER = "MATRIZ_TERRENO"
ESTADOS_PROTEGIDOS = {"R", "NR", "RR", "S"}

MESES_ES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}

MATRIZ_ALIAS_CANONICAL: Dict[str, set[str]] = {
    "AP": {"AP", "AGUA POTABLE"},
    "AR": {"AR", "AGUA RESIDUAL"},
    "ASUB": {
        "ASUB",
        "AGUA SUBTERRANEA",
        "AGUAS SUBTERRANEAS",
        "AGUA SUBTERRANEA S",
    },
    "NF": {
        "NF",
        "NIVEL FREATICO",
        "NIVELES FREATICOS",
    },
    "ASUP": {
        "ASUP",
        "AGUA SUPERFICIAL",
        "AGUAS SUPERFICIALES",
    },
    "CAUDAL": {
        "CAUDAL",
        "CAUDALES",
    },
    "FOTOMETRO": {
        "FOTOMETRO",
        "ESTACION 14",
        "ESTACION14",
    },
    "HK": {"HK", "HOUSEKEEPING"},
}

MATRIZ_ALIAS_CANONICAL_NORMALIZED: Dict[str, set[str]] = {
    canonical: {normalize_text(v) for v in aliases}
    for canonical, aliases in MATRIZ_ALIAS_CANONICAL.items()
}


def _canonical_matriz(value: object) -> str:
    normalized = normalize_text(value)
    if not normalized:
        return ""
    for canonical, normalized_aliases in MATRIZ_ALIAS_CANONICAL_NORMALIZED.items():
        if normalized in normalized_aliases:
            return canonical
    return normalized


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


def _extract_date(value: object) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return None


def _is_estado_header(normalized_header: str) -> bool:
    if not normalized_header:
        return False
    if normalized_header in SEMANAL_ESTADO_COL_HINTS:
        return True
    if "ESTADO" in normalized_header or "STATUS" in normalized_header:
        return True
    return "REALIZADO" in normalized_header


def _select_sheet_by_name(wb, preferred_name: str):
    preferred_norm = normalize_text(preferred_name)
    for ws in wb.worksheets:
        if normalize_text(ws.title) == preferred_norm:
            return ws
    return None


def _detectar_layout_semanal(ws) -> tuple[int, int, int, List[int]]:
    max_row, max_col = _safe_sheet_limits(ws)
    best_date_row = None
    best_day_cols: List[int] = []

    for row in range(1, min(25, max_row) + 1):
        day_cols: List[int] = []
        for col in range(1, max_col + 1):
            if _extract_day(ws.cell(row=row, column=col).value) is not None:
                day_cols.append(col)
        if len(day_cols) > len(best_day_cols):
            best_day_cols = day_cols
            best_date_row = row

    if best_date_row is None or len(best_day_cols) < 2:
        raise ValueError(
            f"No se detecto una fila de fechas/dias valida en hoja semanal '{ws.title}'."
        )

    header_row = max(1, best_date_row - 1)
    data_start_row = best_date_row + 1
    return header_row, best_date_row, data_start_row, best_day_cols


def _detectar_bloques_dia(
    ws, header_row: int, date_row: int, day_cols: List[int]
) -> List[Tuple[int, dt.date | None, int, int, int | None, int]]:
    """
    Retorna bloques diarios en formato:
    (dia, fecha_ref, col_codigo, col_actividad, col_tarea, col_estado)
    """
    _, max_col = _safe_sheet_limits(ws)
    blocks: List[Tuple[int, dt.date | None, int, int, int | None, int]] = []

    for idx, day_col in enumerate(day_cols):
        day_value = ws.cell(row=date_row, column=day_col).value
        day = _extract_day(day_value)
        if day is None:
            continue
        fecha_ref = _extract_date(day_value)

        next_day_col = day_cols[idx + 1] if idx + 1 < len(day_cols) else (max_col + 1)
        search_end = min(next_day_col - 1, day_col + 8, max_col)

        actividad_col = None
        tarea_col = None
        estado_col = None

        for col in range(day_col, search_end + 1):
            header = normalize_column_name(ws.cell(row=header_row, column=col).value)
            if not header:
                continue
            if actividad_col is None and "ACTIVIDAD" in header:
                actividad_col = col
                if "TAREA" in header:
                    tarea_col = col
                continue
            if tarea_col is None and "TAREA" in header:
                tarea_col = col
                continue
            if estado_col is None and _is_estado_header(header):
                estado_col = col

        if actividad_col is None or estado_col is None:
            continue

        blocks.append((day, fecha_ref, day_col, actividad_col, tarea_col, estado_col))

    return blocks


def _leer_semanal_horizontal(
    path_programa_turno: Path | str,
    hoja_preferida: str = "Control SGS",
) -> tuple[pd.DataFrame, str]:
    print("[3/5] Leyendo programa semanal (horizontal)...")
    wb = load_workbook(Path(path_programa_turno), data_only=True, read_only=True)
    records: List[Dict[str, object]] = []
    ws = None
    try:
        ws = _select_sheet_by_name(wb, hoja_preferida)
        if ws is None:
            disponibles = ", ".join(wb.sheetnames)
            raise ValueError(
                f"No se encontro la hoja semanal requerida '{hoja_preferida}'. "
                f"Hojas disponibles: {disponibles}"
            )

        header_row, date_row, data_start_row, day_cols = _detectar_layout_semanal(ws)
        blocks = _detectar_bloques_dia(ws, header_row, date_row, day_cols)
        if not blocks:
            raise ValueError(
                f"No se detectaron bloques diarios validos en hoja semanal '{ws.title}'."
            )

        max_row, _ = _safe_sheet_limits(ws)
        for row in range(data_start_row, max_row + 1):
            for day, fecha_ref, col_codigo, col_actividad, col_tarea, col_estado in blocks:
                codigo = safe_str(ws.cell(row=row, column=col_codigo).value)
                actividad = safe_str(ws.cell(row=row, column=col_actividad).value)
                tarea = safe_str(ws.cell(row=row, column=col_tarea).value) if col_tarea else ""
                estado = safe_str(ws.cell(row=row, column=col_estado).value)

                if not any([actividad, tarea, estado]):
                    continue
                if normalize_text(codigo) in {"TOTAL", "SUBTOTAL"}:
                    continue

                records.append(
                    {
                        "CODIGO": codigo,
                        "ACTIVIDAD": actividad,
                        "TAREA": tarea,
                        "ACTIVIDAD_DICCIONARIO": first_non_empty([actividad, tarea]),
                        "DIA": day,
                        "ESTADO": estado,
                        "FECHA_REFERENCIA": fecha_ref,
                        "ORIGEN_HOJA": ws.title,
                        "ORIGEN_FILA": row,
                    }
                )
    finally:
        wb.close()

    if not records:
        hoja_nombre = ws.title if ws is not None else hoja_preferida
        raise ValueError(
            f"No se detectaron registros semanales en formato horizontal en '{hoja_nombre}'."
        )

    df = pd.DataFrame(records)
    df["DIA"] = pd.to_numeric(df["DIA"], errors="coerce").astype("Int64")
    df = df[df["DIA"].notna()].copy()
    return df, ws.title if ws is not None else hoja_preferida


def _load_diccionario_matriz(path_diccionario: Path | str, cache_dir: Path | str | None) -> Dict[str, str]:
    diccionarios = cargar_diccionarios(path_diccionario, cache_dir=cache_dir)
    raw = diccionarios.get("MATRIZ", {})
    result: Dict[str, str] = {}
    for k, v in raw.items():
        nk = normalize_text(k)
        vv = safe_str(v)
        if nk and vv:
            result[nk] = vv
    return result


def _clasificar_matriz(
    semanal_df: pd.DataFrame, dic_matriz: Dict[str, str]
) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = semanal_df.copy()

    def _resolver_matriz(row: pd.Series) -> tuple[str, str]:
        actividad = safe_str(row.get("ACTIVIDAD", ""))
        tarea = safe_str(row.get("TAREA", ""))
        base = safe_str(row.get("ACTIVIDAD_DICCIONARIO", ""))
        for candidate in [actividad, tarea, base]:
            key = normalize_text(candidate)
            if key and key in dic_matriz:
                return safe_str(dic_matriz[key]), candidate
        return "", first_non_empty([actividad, tarea, base])

    resolved = df.apply(_resolver_matriz, axis=1, result_type="expand")
    resolved.columns = ["MATRIZ", "ACTIVIDAD_DICCIONARIO"]
    df["MATRIZ"] = resolved["MATRIZ"]
    df["ACTIVIDAD_DICCIONARIO"] = resolved["ACTIVIDAD_DICCIONARIO"]

    no_clasificadas = df[df["MATRIZ"].map(safe_str).eq("")].copy()
    if not no_clasificadas.empty:
        no_clasificadas = no_clasificadas[
            ["CODIGO", "ACTIVIDAD", "TAREA", "ACTIVIDAD_DICCIONARIO", "DIA", "ESTADO", "ORIGEN_HOJA", "ORIGEN_FILA"]
        ].drop_duplicates()
    else:
        no_clasificadas = pd.DataFrame(
            columns=[
                "CODIGO",
                "ACTIVIDAD",
                "TAREA",
                "ACTIVIDAD_DICCIONARIO",
                "DIA",
                "ESTADO",
                "ORIGEN_HOJA",
                "ORIGEN_FILA",
            ]
        )
    return df, no_clasificadas


def _find_header_columns(ws) -> tuple[int, int, Dict[int, int], int, int]:
    max_row, max_col = _safe_sheet_limits(ws)

    codigo_col = None
    matriz_col = None
    header_row = None
    best_header_score = -1

    codigo_priority = [normalize_column_name(v) for v in MENSUAL_CODIGO_HEADERS_PRIORITY]
    matriz_required = normalize_column_name(MENSUAL_MATRIZ_TERRENO_HEADER)

    for row in range(1, min(40, max_row) + 1):
        normalized_by_col: Dict[int, str] = {}
        for col in range(1, max_col + 1):
            normalized = normalize_column_name(ws.cell(row=row, column=col).value)
            if normalized:
                normalized_by_col[col] = normalized

        c_col = None
        for header in codigo_priority:
            for col, normalized in normalized_by_col.items():
                if normalized == header:
                    c_col = col
                    break
            if c_col is not None:
                break

        m_col = None
        for col, normalized in normalized_by_col.items():
            if normalized == matriz_required:
                m_col = col
                break

        score = int(c_col is not None) + int(m_col is not None)
        if score > best_header_score:
            best_header_score = score
            header_row = row
            codigo_col = c_col
            matriz_col = m_col

    if codigo_col is None or matriz_col is None or header_row is None:
        raise ValueError(
            "No se detectaron columnas CODIGO y MATRIZ TERRENO en programa mensual."
        )

    best_day_row = None
    best_day_count = -1
    best_day_map: Dict[int, int] = {}
    for row in range(1, min(40, max_row) + 1):
        day_to_col: Dict[int, int] = {}
        for col in range(1, max_col + 1):
            day = _extract_day(ws.cell(row=row, column=col).value)
            if day is not None:
                day_to_col[day] = col
        if len(day_to_col) > best_day_count:
            best_day_count = len(day_to_col)
            best_day_row = row
            best_day_map = day_to_col

    if best_day_row is None or not best_day_map:
        raise ValueError("No se detectaron columnas de dia (1..31) en programa mensual.")

    data_start_row = max(header_row, best_day_row) + 1
    return codigo_col, matriz_col, best_day_map, data_start_row, max_row


def _inferir_hoja_mensual_desde_semanal(semanal_df: pd.DataFrame) -> str | None:
    if "FECHA_REFERENCIA" not in semanal_df.columns:
        return None
    fechas = pd.to_datetime(semanal_df["FECHA_REFERENCIA"], errors="coerce").dropna()
    if fechas.empty:
        return None
    periodo = fechas.dt.to_period("M").mode()
    if periodo.empty:
        return None
    p = periodo.iloc[0]
    mes = MESES_ES.get(int(p.month))
    if mes is None:
        return None
    return f"{mes} {int(p.year)}"


def _seleccionar_hoja_mensual(wb, hoja_preferida: str | None = None):
    if hoja_preferida:
        ws_preferida = _select_sheet_by_name(wb, hoja_preferida)
        if ws_preferida is not None:
            try:
                parsed = _find_header_columns(ws_preferida)
                return ws_preferida, parsed
            except Exception as exc:
                raise ValueError(
                    f"La hoja mensual '{ws_preferida.title}' no tiene estructura valida."
                ) from exc

    best = None
    best_days = -1
    for ws in wb.worksheets:
        try:
            parsed = _find_header_columns(ws)
        except Exception:
            continue
        day_count = len(parsed[2])
        if day_count > best_days:
            best = (ws, parsed)
            best_days = day_count

    if best is None:
        raise ValueError(
            "No se detecto una hoja mensual valida con columnas CODIGO/MATRIZ y dias 1..31."
        )
    return best


def _es_estado_protegido(value: object) -> bool:
    return normalize_text(value) in ESTADOS_PROTEGIDOS


def _is_housekeeping_codigo(value: object) -> bool:
    normalized = normalize_text(value).replace(" ", "")
    return normalized in {"HOUSEKEEPING", "HOUSKEEPING"}


def _pick_housekeeping_row(candidates: List[Tuple[int, str]]) -> int | None:
    if not candidates:
        return None
    for row, matriz in candidates:
        if _canonical_matriz(matriz) == "HK":
            return row
    return candidates[0][0]


def _build_diagnostico_no_cruzados(
    no_cruzados_df: pd.DataFrame,
    matrices_por_codigo: Dict[str, set[str]],
    row_lookup: Dict[tuple[str, str], int],
    day_to_col: Dict[int, int],
) -> pd.DataFrame:
    columns = [
        "CODIGO",
        "MATRIZ_SEMANAL",
        "DIA",
        "ESTADO",
        "MOTIVO_ORIGINAL",
        "CODIGO_EXISTE_EN_MENSUAL",
        "MATRICES_EN_MENSUAL_PARA_CODIGO",
        "COINCIDE_CODIGO_MATRIZ",
        "COLUMNA_DIA_EXISTE",
    ]
    if no_cruzados_df.empty:
        return pd.DataFrame(columns=columns)

    rows: List[Dict[str, object]] = []
    for _, rec in no_cruzados_df.iterrows():
        codigo = safe_str(rec.get("CODIGO", ""))
        matriz = safe_str(rec.get("MATRIZ", ""))
        estado = safe_str(rec.get("ESTADO", ""))
        motivo = safe_str(rec.get("MOTIVO", ""))

        dia_val = rec.get("DIA")
        dia = int(dia_val) if pd.notna(dia_val) else None

        key = normalize_key(codigo, _canonical_matriz(matriz))
        matrices_mensual = sorted(matrices_por_codigo.get(codigo, set()))

        rows.append(
            {
                "CODIGO": codigo,
                "MATRIZ_SEMANAL": matriz,
                "DIA": dia,
                "ESTADO": estado,
                "MOTIVO_ORIGINAL": motivo,
                "CODIGO_EXISTE_EN_MENSUAL": codigo in matrices_por_codigo,
                "MATRICES_EN_MENSUAL_PARA_CODIGO": " | ".join(matrices_mensual),
                "COINCIDE_CODIGO_MATRIZ": key in row_lookup,
                "COLUMNA_DIA_EXISTE": (dia in day_to_col) if dia is not None else False,
            }
        )

    return pd.DataFrame(rows, columns=columns)


def actualizar_programa_mensual(
    path_programa_turno: Path | str,
    path_programa_mensual: Path | str,
    output_path_programa_mensual: Path | str,
    path_diccionario: Path | str,
) -> pd.DataFrame:
    print("[3/5] Actualizando programa mensual...")
    output_path = Path(output_path_programa_mensual)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    semanal_df, hoja_semanal_usada = _leer_semanal_horizontal(
        path_programa_turno, hoja_preferida="Control SGS"
    )
    dic_matriz = _load_diccionario_matriz(path_diccionario, cache_dir=output_path.parent)
    semanal_df, actividades_no_clasificadas_df = _clasificar_matriz(semanal_df, dic_matriz)

    work_df = semanal_df.copy()
    work_df = work_df[
        work_df["CODIGO"].map(safe_str).ne("")
        & work_df["MATRIZ"].map(safe_str).ne("")
        & work_df["ESTADO"].map(safe_str).ne("")
    ].copy()

    duplicados_df = work_df[
        work_df.duplicated(subset=["CODIGO", "MATRIZ", "DIA"], keep=False)
    ].copy()
    work_df = work_df.drop_duplicates(subset=["CODIGO", "MATRIZ", "DIA"], keep="last")

    hoja_mensual_objetivo = _inferir_hoja_mensual_desde_semanal(semanal_df)
    wb = load_workbook(Path(path_programa_mensual))
    ws, parsed = _seleccionar_hoja_mensual(wb, hoja_preferida=hoja_mensual_objetivo)
    codigo_col, matriz_col, day_to_col, data_start_row, max_row = parsed

    row_lookup: Dict[tuple[str, str], int] = {}
    matrices_por_codigo: Dict[str, set[str]] = {}
    rows_por_codigo: Dict[str, List[Tuple[int, str]]] = {}
    housekeeping_rows: List[Tuple[int, str]] = []
    blank_streak = 0
    for row in range(data_start_row, max_row + 1):
        codigo = safe_str(ws.cell(row=row, column=codigo_col).value)
        matriz = safe_str(ws.cell(row=row, column=matriz_col).value)
        if codigo == "" and matriz == "":
            blank_streak += 1
            if blank_streak >= 200 and row > data_start_row + 200:
                break
            continue
        blank_streak = 0
        matrices_por_codigo.setdefault(codigo, set())
        rows_por_codigo.setdefault(normalize_text(codigo), [])
        rows_por_codigo[normalize_text(codigo)].append((row, matriz))
        if matriz:
            matrices_por_codigo[codigo].add(matriz)
        if _is_housekeeping_codigo(codigo):
            housekeeping_rows.append((row, matriz))

        key = normalize_key(codigo, _canonical_matriz(matriz))
        if key == normalize_key("", ""):
            continue
        # Conserva la primera fila para evitar desfases cuando hay filas duplicadas por codigo/matriz.
        if key not in row_lookup:
            row_lookup[key] = row

    aplicados: List[Dict[str, object]] = []
    no_cruzados: List[Dict[str, object]] = []
    protegidos = 0
    no_actualizados_otro = 0

    for _, record in work_df.iterrows():
        codigo = safe_str(record["CODIGO"])
        matriz = safe_str(record["MATRIZ"])
        dia = int(record["DIA"])
        estado = safe_str(record["ESTADO"])

        key = normalize_key(codigo, _canonical_matriz(matriz))
        row = row_lookup.get(key)
        col = day_to_col.get(dia)

        if row is None and _is_housekeeping_codigo(codigo):
            direct_candidates = rows_por_codigo.get(normalize_text(codigo), [])
            row = _pick_housekeeping_row(direct_candidates)
            if row is None:
                row = _pick_housekeeping_row(housekeeping_rows)

        if row is None or col is None:
            no_cruzados.append(
                {
                    "CODIGO": codigo,
                    "MATRIZ": matriz,
                    "DIA": dia,
                    "ESTADO": estado,
                    "ACTIVIDAD": safe_str(record["ACTIVIDAD_DICCIONARIO"]),
                    "MOTIVO": "SIN_COINCIDENCIA" if row is None else "DIA_NO_ENCONTRADO",
                }
            )
            continue

        cell = ws.cell(row=row, column=col)
        valor_actual = safe_str(cell.value)
        valor_actual_norm = normalize_text(valor_actual)

        if valor_actual_norm in {"", "1"}:
            cell.value = estado
            aplicados.append(
                {
                    "CODIGO": codigo,
                    "MATRIZ": matriz,
                    "DIA": dia,
                    "ESTADO": estado,
                    "ACTIVIDAD": safe_str(record["ACTIVIDAD_DICCIONARIO"]),
                    "ORIGEN_HOJA": safe_str(record["ORIGEN_HOJA"]),
                    "ORIGEN_FILA": int(record["ORIGEN_FILA"]),
                }
            )
        elif _es_estado_protegido(valor_actual):
            protegidos += 1
        else:
            no_actualizados_otro += 1

    wb.save(output_path)

    no_cruzados_df = pd.DataFrame(
        no_cruzados,
        columns=["CODIGO", "MATRIZ", "DIA", "ESTADO", "ACTIVIDAD", "MOTIVO"],
    )
    if no_cruzados_df.empty:
        no_cruzados_df = pd.DataFrame(
            columns=["CODIGO", "MATRIZ", "DIA", "ESTADO", "ACTIVIDAD", "MOTIVO"]
        )

    diagnostico_no_cruzados_df = _build_diagnostico_no_cruzados(
        no_cruzados_df=no_cruzados_df,
        matrices_por_codigo=matrices_por_codigo,
        row_lookup=row_lookup,
        day_to_col=day_to_col,
    )

    if duplicados_df.empty:
        duplicados_df = pd.DataFrame(
            columns=[
                "CODIGO",
                "ACTIVIDAD",
                "TAREA",
                "ACTIVIDAD_DICCIONARIO",
                "DIA",
                "ESTADO",
                "MATRIZ",
                "ORIGEN_HOJA",
                "ORIGEN_FILA",
            ]
        )

    save_dataframe_to_excel(no_cruzados_df, output_path.parent / "no_cruzados.xlsx")
    save_dataframe_to_excel(
        actividades_no_clasificadas_df,
        output_path.parent / "actividades_no_clasificadas.xlsx",
    )
    save_dataframe_to_excel(duplicados_df, output_path.parent / "duplicados.xlsx")
    save_dataframe_to_excel(
        diagnostico_no_cruzados_df,
        output_path.parent / "diagnostico_no_cruzados_rapido.xlsx",
    )

    aplicados_df = pd.DataFrame(
        aplicados,
        columns=["CODIGO", "MATRIZ", "DIA", "ESTADO", "ACTIVIDAD", "ORIGEN_HOJA", "ORIGEN_FILA"],
    )
    if aplicados_df.empty:
        aplicados_df = pd.DataFrame(
            columns=["CODIGO", "MATRIZ", "DIA", "ESTADO", "ACTIVIDAD", "ORIGEN_HOJA", "ORIGEN_FILA"]
        )

    aplicados_df["PUNTO"] = aplicados_df["CODIGO"]
    aplicados_df["TIPO"] = aplicados_df["MATRIZ"]
    aplicados_df["COMENTARIO"] = aplicados_df["ACTIVIDAD"]

    print(
        "[3/5] Hojas usadas: "
        f"semanal='{hoja_semanal_usada}' | mensual='{ws.title}' (objetivo='{hoja_mensual_objetivo or 'auto'}')"
    )
    print(
        "[3/5] Semanal total: "
        f"{len(semanal_df)} | clasificados matriz: {(semanal_df['MATRIZ'].map(safe_str) != '').sum()} "
        f"| aplicados mensual: {len(aplicados_df)} | no cruzados: {len(no_cruzados_df)} "
        f"| actividades no clasificadas: {len(actividades_no_clasificadas_df)} "
        f"| duplicados: {len(duplicados_df)} | protegidos sin cambio: {protegidos} "
        f"| no actualizados (otro valor): {no_actualizados_otro}"
    )

    return aplicados_df[
        [
            "PUNTO",
            "TIPO",
            "DIA",
            "ESTADO",
            "COMENTARIO",
            "CODIGO",
            "MATRIZ",
            "ACTIVIDAD",
            "ORIGEN_HOJA",
            "ORIGEN_FILA",
        ]
    ]
