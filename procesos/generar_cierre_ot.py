from __future__ import annotations

import pandas as pd

from utils.excel_utils import validate_required_columns
from utils.texto_utils import normalize_column_name, normalize_key, safe_str


def generar_cierre_ot_base(df_limpio: pd.DataFrame) -> pd.DataFrame:
    print("[2/5] Generando Cierre OT base...")
    validate_required_columns(
        df_limpio, ["OT", "PUNTO", "TIPO", "FECHA", "DIA"], "MSEWJO limpio"
    )

    base = df_limpio[["OT", "PUNTO", "TIPO", "FECHA", "DIA"]].copy()
    base["CODIGO_CIERRE"] = ""
    base["COMENTARIO"] = ""

    before = len(base)
    base = base.drop_duplicates(subset=["OT", "PUNTO", "TIPO", "DIA"]).copy()
    print(f"[2/5] Filas base: {len(base)} (duplicados removidos: {before - len(base)})")

    return base[
        ["OT", "PUNTO", "TIPO", "FECHA", "DIA", "CODIGO_CIERRE", "COMENTARIO"]
    ]


def actualizar_cierre_ot(cierre_base_df: pd.DataFrame, clasificados_df: pd.DataFrame) -> pd.DataFrame:
    print("[5/5] Actualizando Cierre OT final...")

    cierre = cierre_base_df.copy()
    clasificados = clasificados_df.copy()

    cierre_norm_map = {normalize_column_name(col): col for col in cierre.columns}
    clasif_norm_map = {normalize_column_name(col): col for col in clasificados.columns}

    cierre_desc_col = cierre_norm_map.get(normalize_column_name("Descripcion de tarea programada 2"))
    cierre_matriz_col = cierre_norm_map.get(normalize_column_name("Matriz Terreno"))
    if cierre_desc_col is None or cierre_matriz_col is None:
        print(
            "[5/5] Advertencia: cierre base no contiene columnas "
            "'Descripcion de tarea programada 2' + 'Matriz Terreno'. "
            "Se conserva cierre base sin merge tecnico."
        )
        return cierre

    connector_col = clasif_norm_map.get(normalize_column_name("CONECTOR"))
    desc_col = clasif_norm_map.get(normalize_column_name("DESC_TAREA_2"))
    tipo_col = clasif_norm_map.get(normalize_column_name("TIPO"))
    codigo_col = clasif_norm_map.get(normalize_column_name("CODIGO_CIERRE"))
    comentario_col = clasif_norm_map.get(normalize_column_name("COMENTARIO"))

    if connector_col is None:
        if desc_col is None or tipo_col is None:
            print(
                "[5/5] Advertencia: clasificados no contiene 'CONECTOR' "
                "ni columnas alternativas para construir llave. "
                "Se conserva cierre base sin merge tecnico."
            )
            return cierre
        clasificados["_KEY"] = (
            clasificados[desc_col].astype(str) + "/" + clasificados[tipo_col].astype(str)
        ).apply(normalize_key)
    else:
        clasificados["_KEY"] = clasificados[connector_col].astype(str).apply(normalize_key)

    if codigo_col is None:
        clasificados["CODIGO_CIERRE"] = ""
    else:
        clasificados["CODIGO_CIERRE"] = clasificados[codigo_col].map(safe_str)

    if comentario_col is None:
        clasificados["COMENTARIO"] = ""
    else:
        clasificados["COMENTARIO"] = clasificados[comentario_col].map(safe_str)

    # Llave cierre OT
    cierre["_KEY"] = (
        cierre[cierre_desc_col].astype(str) + "/" + cierre[cierre_matriz_col].astype(str)
    ).apply(normalize_key)

    clasificados = clasificados.drop_duplicates(subset=["_KEY"], keep="last")

    cierre = cierre.merge(
        clasificados[["_KEY", "CODIGO_CIERRE", "COMENTARIO"]],
        on="_KEY",
        how="left",
        suffixes=("", "_NEW"),
    )

    cierre["CODIGO_CIERRE"] = cierre["CODIGO_CIERRE_NEW"].fillna(cierre["CODIGO_CIERRE"])
    cierre["COMENTARIO"] = cierre["COMENTARIO_NEW"].fillna(cierre["COMENTARIO"])

    cierre = cierre.drop(columns=["_KEY", "CODIGO_CIERRE_NEW", "COMENTARIO_NEW"])

    return cierre
