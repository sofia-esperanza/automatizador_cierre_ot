from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd

from utils.excel_utils import (
    read_excel_file,
    rename_columns_by_alias,
    validate_required_columns,
)
from utils.texto_utils import contains_keyword, normalize_text, safe_str

CLASIFICACION_ALIASES = {
    "ESTADO": ["ESTADO"],
    "PALABRA_CLAVE_COMENTARIO": [
        "PALABRA_CLAVE_COMENTARIO",
        "PALABRA_CLAVE",
        "KEYWORD",
    ],
    "CODIGO_CIERRE": ["CODIGO_CIERRE", "CODIGO_DE_CIERRE", "CODIGO"],
}


def _cargar_reglas(path_matriz_clasificacion: Path | str) -> pd.DataFrame:
    df = read_excel_file(path_matriz_clasificacion)
    df = rename_columns_by_alias(df, CLASIFICACION_ALIASES)
    validate_required_columns(
        df,
        ["ESTADO", "PALABRA_CLAVE_COMENTARIO", "CODIGO_CIERRE"],
        "matriz de clasificacion",
    )

    df["ESTADO"] = df["ESTADO"].map(safe_str)
    df["PALABRA_CLAVE_COMENTARIO"] = df["PALABRA_CLAVE_COMENTARIO"].map(safe_str)
    df["CODIGO_CIERRE"] = df["CODIGO_CIERRE"].map(safe_str)
    return df


def clasificar_registros(
    registros_df: pd.DataFrame, path_matriz_clasificacion: Path | str
) -> pd.DataFrame:
    print("[4/5] Clasificando registros...")
    validate_required_columns(
        registros_df, ["PUNTO", "TIPO", "DIA", "ESTADO", "COMENTARIO"], "registros turno"
    )

    reglas_df = _cargar_reglas(path_matriz_clasificacion)

    reglas_por_estado: Dict[str, List[Tuple[str, str]]] = defaultdict(list)
    for _, rule in reglas_df.iterrows():
        estado_norm = normalize_text(rule["ESTADO"])
        keyword = safe_str(rule["PALABRA_CLAVE_COMENTARIO"])
        codigo = safe_str(rule["CODIGO_CIERRE"])
        if not estado_norm or not codigo:
            continue
        reglas_por_estado[estado_norm].append((keyword, codigo))

    classified = registros_df.copy()
    codigos = []

    for _, row in classified.iterrows():
        estado_norm = normalize_text(row["ESTADO"])
        comentario = safe_str(row["COMENTARIO"])
        reglas_estado = reglas_por_estado.get(estado_norm, [])

        codigo = ""
        fallback = ""
        for keyword, candidate_code in reglas_estado:
            if keyword:
                if contains_keyword(comentario, keyword):
                    codigo = candidate_code
                    break
            elif not fallback:
                fallback = candidate_code

        if not codigo:
            codigo = fallback

        codigos.append(codigo)

    classified["CODIGO_CIERRE"] = codigos
    print(
        f"[4/5] Registros clasificados con codigo: "
        f"{(classified['CODIGO_CIERRE'].astype(str).str.strip() != '').sum()} "
        f"de {len(classified)}"
    )
    return classified
