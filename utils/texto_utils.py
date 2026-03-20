from __future__ import annotations

import re
import unicodedata
from typing import Iterable, Tuple


def safe_str(value: object) -> str:
    """Convierte cualquier valor a string seguro."""
    if value is None:
        return ""
    if isinstance(value, float) and str(value) == "nan":
        return ""
    return str(value).strip()


def normalize_text(value: object) -> str:
    """
    Normaliza texto para comparaciones:
    - quita acentos
    - convierte a mayuscula
    - colapsa espacios
    """
    text = safe_str(value)
    normalized = unicodedata.normalize("NFKD", text)
    no_accents = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    compact = re.sub(r"\s+", " ", no_accents)
    return compact.strip().upper()


def normalize_column_name(value: object) -> str:
    """Normaliza encabezados para mapear columnas con nombres variantes."""
    text = normalize_text(value)
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def contains_keyword(text: object, keyword: object) -> bool:
    """Busca keyword (normalizada) dentro de text (normalizado)."""
    normalized_keyword = normalize_text(keyword)
    if not normalized_keyword:
        return False
    return normalized_keyword in normalize_text(text)


def normalize_key(*values: object) -> Tuple[str, ...]:
    """Crea una llave de comparacion robusta para merges por texto."""
    return tuple(normalize_text(v) for v in values)


def first_non_empty(values: Iterable[object]) -> str:
    """Retorna el primer valor no vacio de un iterable."""
    for value in values:
        text = safe_str(value)
        if text:
            return text
    return ""
