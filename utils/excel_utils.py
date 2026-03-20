from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List

import pandas as pd

from utils.texto_utils import normalize_column_name


class ColumnValidationError(ValueError):
    """Error de validacion cuando faltan columnas requeridas."""


def read_excel_file(path: Path | str, sheet_name: int | str = 0) -> pd.DataFrame:
    excel_path = Path(path)
    if not excel_path.exists():
        raise FileNotFoundError(f"No existe el archivo: {excel_path}")
    return pd.read_excel(excel_path, sheet_name=sheet_name)


def normalize_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_column_name(c) for c in df.columns]
    return df


def rename_columns_by_alias(
    df: pd.DataFrame, aliases: Dict[str, Iterable[str]]
) -> pd.DataFrame:
    """
    Renombra columnas usando aliases.
    aliases ejemplo:
    {
        "PUNTO": ["PUNTO", "PUNTO_MONITOREO", "ESTACION"]
    }
    """
    df = df.copy()
    alias_map: Dict[str, str] = {}
    for canonical, options in aliases.items():
        canonical_norm = normalize_column_name(canonical)
        for alias in options:
            alias_map[normalize_column_name(alias)] = canonical_norm

    rename_dict: Dict[str, str] = {}
    for column in df.columns:
        column_norm = normalize_column_name(column)
        if column_norm in alias_map:
            rename_dict[column] = alias_map[column_norm]

    return df.rename(columns=rename_dict)


def validate_required_columns(
    df: pd.DataFrame, required_columns: Iterable[str], context: str
) -> None:
    required = [normalize_column_name(c) for c in required_columns]
    missing: List[str] = [c for c in required if c not in df.columns]
    if missing:
        missing_str = ", ".join(missing)
        raise ColumnValidationError(
            f"Faltan columnas requeridas en {context}: {missing_str}"
        )


def save_dataframe_to_excel(df: pd.DataFrame, path: Path | str) -> Path:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    return output_path
