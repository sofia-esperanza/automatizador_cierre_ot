from __future__ import annotations

import pandas as pd

from utils.excel_utils import validate_required_columns
from utils.texto_utils import normalize_key, safe_str


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


def actualizar_cierre_ot(
    cierre_base_df: pd.DataFrame, clasificados_df: pd.DataFrame
) -> pd.DataFrame:
    print("[5/5] Actualizando Cierre OT final...")
    validate_required_columns(
        clasificados_df,
        ["PUNTO", "TIPO", "DIA", "CODIGO_CIERRE", "COMENTARIO"],
        "registros clasificados",
    )

    cierre = cierre_base_df.copy()
    clasificados = clasificados_df.copy()

    cierre["_KEY"] = cierre.apply(
        lambda r: normalize_key(r["PUNTO"], r["TIPO"], r["DIA"]), axis=1
    )
    clasificados["_KEY"] = clasificados.apply(
        lambda r: normalize_key(r["PUNTO"], r["TIPO"], r["DIA"]), axis=1
    )

    clasificados = clasificados.drop_duplicates(subset=["_KEY"], keep="last")
    clasificados["CODIGO_CIERRE"] = clasificados["CODIGO_CIERRE"].map(safe_str)
    clasificados["COMENTARIO"] = clasificados["COMENTARIO"].map(safe_str)

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
