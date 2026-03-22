from __future__ import annotations

import json
import os
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List

import pandas as pd

from utils.excel_utils import (
    save_dataframe_to_excel,
    validate_required_columns,
)
from utils.texto_utils import normalize_column_name, normalize_text, safe_str

CACHE_FILENAME = "diccionario_cache.json"
META_FILENAME = "diccionario_meta.json"

MATRIZ_COL_ALIASES = {
    "ACTIVIDAD": [
        "ACTIVIDAD",
        "ACTIVIDAD_TAREA",
        "ACTIVIDAD O TAREA",
        "TAREA",
        "DESCRIPCION",
        "DESCRIPCION_TAREA",
    ],
    "MATRIZ_TERRENO": [
        "MATRIZ_TERRENO",
        "MATRIZ TERRENO",
        "MATRIZ",
        "TIPO_MATRIZ",
    ],
}

CIERRE_COL_ALIASES = {
    "ESTADO": ["ESTADO", "STATUS", "RESULTADO"],
    "PALABRA_CLAVE_COMENTARIO": [
        "PALABRA_CLAVE_COMENTARIO",
        "PALABRA CLAVE COMENTARIO",
        "PALABRA_CLAVE",
        "KEYWORD",
    ],
    "CODIGO_CIERRE": [
        "CODIGO_CIERRE",
        "CODIGO_DE_CIERRE",
        "CODIGO",
        "CÓDIGO_DE_CIERRE",
    ],
}

ACTIVIDAD_INPUT_ALIASES = [
    "ACTIVIDAD",
    "ACTIVIDAD_TAREA",
    "DESCRIPCION_TAREA_2",
    "DESCRIPCION",
    "TAREA",
    "TIPO",
]


def limpiar_texto(value: object) -> str:
    return normalize_text(value)


def _cache_paths(path_diccionario_excel: Path | str, cache_dir: Path | str | None = None) -> tuple[Path, Path]:
    base = Path(cache_dir) if cache_dir is not None else Path(path_diccionario_excel).parent
    base.mkdir(parents=True, exist_ok=True)
    return base / CACHE_FILENAME, base / META_FILENAME


def _read_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError(f"JSON invalido en {path}")
    return data


def _write_json(path: Path, payload: Dict[str, Any]) -> None:
    with path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def _find_column(df: pd.DataFrame, aliases: Iterable[str]) -> str | None:
    norm_map = {normalize_column_name(col): col for col in df.columns}
    for alias in aliases:
        candidate = norm_map.get(normalize_column_name(alias))
        if candidate is not None:
            return candidate
    return None


def _sheet_matches(df: pd.DataFrame, aliases_map: Dict[str, Iterable[str]]) -> bool:
    for alias_group in aliases_map.values():
        if _find_column(df, alias_group) is None:
            return False
    return True


def _load_sheets(path_diccionario_excel: Path | str) -> Dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(Path(path_diccionario_excel))
    return {sheet_name: excel.parse(sheet_name=sheet_name) for sheet_name in excel.sheet_names}


def _select_sheet(
    sheets: Dict[str, pd.DataFrame],
    aliases_map: Dict[str, Iterable[str]],
    preferred_names: Iterable[str],
) -> tuple[str | None, pd.DataFrame | None]:
    preferred_norm = [normalize_column_name(name) for name in preferred_names]
    for sheet_name, df in sheets.items():
        if normalize_column_name(sheet_name) in preferred_norm and _sheet_matches(df, aliases_map):
            return sheet_name, df

    for sheet_name, df in sheets.items():
        if _sheet_matches(df, aliases_map):
            return sheet_name, df
    return None, None


def _build_diccionarios_from_excel(path_diccionario_excel: Path | str) -> Dict[str, Any]:
    sheets = _load_sheets(path_diccionario_excel)

    _, matriz_df = _select_sheet(
        sheets=sheets,
        aliases_map=MATRIZ_COL_ALIASES,
        preferred_names=["MATRIZ_TERRENO", "MATRIZ TERRENO"],
    )
    _, cierre_df = _select_sheet(
        sheets=sheets,
        aliases_map=CIERRE_COL_ALIASES,
        preferred_names=["COD_CIERRE", "COD CIERRE", "CLASIFICACION", "MATRIZ_CLASIFICACION"],
    )

    matriz_dict: Dict[str, str] = {}
    if matriz_df is not None:
        actividad_col = _find_column(matriz_df, MATRIZ_COL_ALIASES["ACTIVIDAD"])
        matriz_col = _find_column(matriz_df, MATRIZ_COL_ALIASES["MATRIZ_TERRENO"])
        if actividad_col and matriz_col:
            for _, row in matriz_df.iterrows():
                actividad_norm = limpiar_texto(row.get(actividad_col))
                matriz_val = safe_str(row.get(matriz_col))
                if actividad_norm and matriz_val:
                    matriz_dict[actividad_norm] = matriz_val

    cierre_dict: Dict[str, List[Dict[str, str]]] = defaultdict(list)
    if cierre_df is not None:
        estado_col = _find_column(cierre_df, CIERRE_COL_ALIASES["ESTADO"])
        keyword_col = _find_column(cierre_df, CIERRE_COL_ALIASES["PALABRA_CLAVE_COMENTARIO"])
        codigo_col = _find_column(cierre_df, CIERRE_COL_ALIASES["CODIGO_CIERRE"])
        if estado_col and keyword_col and codigo_col:
            for _, row in cierre_df.iterrows():
                estado_norm = limpiar_texto(row.get(estado_col))
                keyword_norm = limpiar_texto(row.get(keyword_col))
                codigo = safe_str(row.get(codigo_col))
                if estado_norm and codigo:
                    cierre_dict[estado_norm].append({"keyword": keyword_norm, "codigo": codigo})

    return {"MATRIZ": matriz_dict, "CIERRE": dict(cierre_dict)}


def cargar_diccionarios(
    path_diccionario_excel: Path | str,
    cache_dir: Path | str | None = None,
) -> Dict[str, Any]:
    """
    Lee diccionarios desde Excel y usa cache JSON cuando no hay cambios.
    Regenera cache automaticamente si cambia el Excel fuente.
    """
    source_path = Path(path_diccionario_excel)
    if not source_path.exists():
        raise FileNotFoundError(f"No existe el archivo de diccionarios: {source_path}")

    cache_warning = ""
    try:
        cache_path, meta_path = _cache_paths(source_path, cache_dir=cache_dir)
    except Exception as exc:
        fallback = Path.cwd() / "_diccionario_cache"
        cache_path, meta_path = _cache_paths(source_path, cache_dir=fallback)
        cache_warning = (
            "[ADVERTENCIA] No se pudo usar carpeta de cache junto al diccionario; "
            f"se usara fallback en {fallback}. Motivo: {exc}"
        )

    if cache_warning:
        print(cache_warning)
    source_mtime = os.path.getmtime(source_path)

    if cache_path.exists() and meta_path.exists():
        try:
            meta = _read_json(meta_path)
            cache = _read_json(cache_path)
            cached_mtime = float(meta.get("source_mtime", -1))
            if cached_mtime == float(source_mtime):
                if "MATRIZ" in cache and "CIERRE" in cache:
                    print("[4/5] Diccionarios cargados desde cache.")
                    return cache
        except Exception:
            pass

    print("[4/5] Leyendo diccionarios desde Excel y regenerando cache...")
    diccionarios = _build_diccionarios_from_excel(source_path)
    try:
        _write_json(cache_path, diccionarios)
        _write_json(
            meta_path,
            {
                "source_path": str(source_path.resolve()),
                "source_mtime": float(source_mtime),
            },
        )
    except Exception as exc:
        print(
            "[ADVERTENCIA] No se pudo escribir cache de diccionarios. "
            f"Se continuara sin cache. Motivo: {exc}"
        )
    return diccionarios


def clasificar_matriz(
    actividad_tarea: object,
    diccionario_matriz: Dict[str, str],
    nuevas_actividades: List[str],
) -> str:
    actividad_original = safe_str(actividad_tarea)
    actividad_norm = limpiar_texto(actividad_tarea)
    if not actividad_norm:
        return ""

    matriz = safe_str(diccionario_matriz.get(actividad_norm, ""))
    if not matriz:
        nuevas_actividades.append(actividad_original)
    return matriz


def clasificar_cierre(
    estado: object,
    comentario: object,
    diccionario_cierre: Dict[str, List[Dict[str, str]]],
    comentarios_no_clasificados: List[Dict[str, str]],
) -> str:
    estado_norm = limpiar_texto(estado)
    comentario_norm = limpiar_texto(comentario)
    reglas_estado = diccionario_cierre.get(estado_norm, [])

    codigo = ""
    fallback = ""
    for rule in reglas_estado:
        keyword_norm = limpiar_texto(rule.get("keyword", ""))
        candidate_code = safe_str(rule.get("codigo", ""))
        if not candidate_code:
            continue

        if keyword_norm:
            if keyword_norm in comentario_norm:
                codigo = candidate_code
                break
        elif not fallback:
            fallback = candidate_code

    if not codigo:
        codigo = fallback

    if not codigo:
        comentario_original = safe_str(comentario)
        if comentario_original:
            comentarios_no_clasificados.append(
                {
                    "ESTADO": safe_str(estado),
                    "COMENTARIO": comentario_original,
                }
            )
    return codigo


def _dedup_strings(values: Iterable[str]) -> List[str]:
    seen = set()
    result: List[str] = []
    for raw in values:
        value = safe_str(raw)
        if not value:
            continue
        key = limpiar_texto(value)
        if key in seen:
            continue
        seen.add(key)
        result.append(value)
    return result


def _dedup_comment_rows(rows: Iterable[Dict[str, str]]) -> List[Dict[str, str]]:
    seen = set()
    result: List[Dict[str, str]] = []
    for row in rows:
        estado = safe_str(row.get("ESTADO"))
        comentario = safe_str(row.get("COMENTARIO"))
        if not comentario:
            continue
        key = (limpiar_texto(estado), limpiar_texto(comentario))
        if key in seen:
            continue
        seen.add(key)
        result.append({"ESTADO": estado, "COMENTARIO": comentario})
    return result


def detectar_nuevos_valores(
    nuevas_actividades: List[str],
    comentarios_no_clasificados: List[Dict[str, str]],
    modo: str = "automatico",
    export_dir: Path | str | None = None,
    diccionarios: Dict[str, Any] | None = None,
) -> Dict[str, List[Any]]:
    modo_norm = limpiar_texto(modo).lower()
    nuevas_actividades_u = _dedup_strings(nuevas_actividades)
    comentarios_no_clasificados_u = _dedup_comment_rows(comentarios_no_clasificados)

    if not nuevas_actividades_u and not comentarios_no_clasificados_u:
        return {
            "nuevas_actividades": [],
            "comentarios_no_clasificados": [],
        }

    if modo_norm in {"automatico", "auto"}:
        if nuevas_actividades_u:
            print("[ADVERTENCIA] Nuevas actividades sin matriz detectadas:")
            for item in nuevas_actividades_u:
                print(f"  - {item}")
        if comentarios_no_clasificados_u:
            print("[ADVERTENCIA] Comentarios sin clasificar detectados:")
            for item in comentarios_no_clasificados_u:
                print(f"  - ESTADO={item['ESTADO']} | COMENTARIO={item['COMENTARIO']}")

    elif modo_norm in {"interactivo", "interactiva"}:
        if diccionarios is None:
            diccionarios = {"MATRIZ": {}, "CIERRE": {}}
        matriz_dict = diccionarios.setdefault("MATRIZ", {})
        cierre_dict = diccionarios.setdefault("CIERRE", {})

        for actividad in nuevas_actividades_u:
            respuesta = input(
                f"[INTERACTIVO] Matriz para actividad '{actividad}' (enter para omitir): "
            ).strip()
            if respuesta:
                matriz_dict[limpiar_texto(actividad)] = respuesta

        for row in comentarios_no_clasificados_u:
            estado = row["ESTADO"]
            comentario = row["COMENTARIO"]
            respuesta = input(
                f"[INTERACTIVO] Codigo cierre para ESTADO='{estado}' y comentario "
                f"'{comentario}' (enter para omitir): "
            ).strip()
            if respuesta:
                estado_norm = limpiar_texto(estado)
                cierre_dict.setdefault(estado_norm, []).append(
                    {"keyword": limpiar_texto(comentario), "codigo": respuesta}
                )

    elif modo_norm in {"controlado", "exportar"}:
        export_base = Path(export_dir) if export_dir is not None else Path.cwd()
        export_base.mkdir(parents=True, exist_ok=True)

        if nuevas_actividades_u:
            df_new_acts = pd.DataFrame({"ACTIVIDAD": nuevas_actividades_u})
            save_dataframe_to_excel(df_new_acts, export_base / "nuevas_actividades.xlsx")
            print(
                "[ADVERTENCIA] Se exportaron actividades nuevas a "
                f"{export_base / 'nuevas_actividades.xlsx'}"
            )

        if comentarios_no_clasificados_u:
            df_new_comments = pd.DataFrame(comentarios_no_clasificados_u)
            save_dataframe_to_excel(
                df_new_comments, export_base / "comentarios_no_clasificados.xlsx"
            )
            print(
                "[ADVERTENCIA] Se exportaron comentarios no clasificados a "
                f"{export_base / 'comentarios_no_clasificados.xlsx'}"
            )

    else:
        print(
            f"[ADVERTENCIA] Modo de nuevos valores no reconocido: '{modo}'. "
            "Se usara modo automatico."
        )
        return detectar_nuevos_valores(
            nuevas_actividades=nuevas_actividades_u,
            comentarios_no_clasificados=comentarios_no_clasificados_u,
            modo="automatico",
            export_dir=export_dir,
            diccionarios=diccionarios,
        )

    return {
        "nuevas_actividades": nuevas_actividades_u,
        "comentarios_no_clasificados": comentarios_no_clasificados_u,
    }


def _actualizar_diccionario_fuente(
    path_diccionario_excel: Path | str,
    nuevas_actividades: List[str],
    comentarios_no_clasificados: List[Dict[str, str]],
) -> Dict[str, int]:
    """
    Actualiza el Excel de diccionario agregando filas nuevas pendientes:
    - MATRIZ_TERRENO: actividad nueva con MATRIZ TERRENO en blanco.
    - COD_CIERRE: estado/comentario nuevo con CODIGO en blanco.
    """
    source = Path(path_diccionario_excel)
    if not source.exists():
        raise FileNotFoundError(f"No existe el archivo de diccionario para actualizar: {source}")

    sheets = _load_sheets(source)
    matriz_sheet_name, matriz_df = _select_sheet(
        sheets=sheets,
        aliases_map=MATRIZ_COL_ALIASES,
        preferred_names=["MATRIZ_TERRENO", "MATRIZ TERRENO"],
    )
    cierre_sheet_name, cierre_df = _select_sheet(
        sheets=sheets,
        aliases_map=CIERRE_COL_ALIASES,
        preferred_names=["COD_CIERRE", "COD CIERRE", "CLASIFICACION", "MATRIZ_CLASIFICACION"],
    )

    matriz_sheet_name = matriz_sheet_name or "MATRIZ_TERRENO"
    cierre_sheet_name = cierre_sheet_name or "COD_CIERRE"

    if matriz_df is None:
        matriz_df = pd.DataFrame(columns=["ACTIVIDAD", "MATRIZ TERRENO"])
    if cierre_df is None:
        cierre_df = pd.DataFrame(
            columns=["ESTADO", "PALABRA CLAVE COMENTARIO", "CÓDIGO DE CIERRE"]
        )

    actividad_col = _find_column(matriz_df, MATRIZ_COL_ALIASES["ACTIVIDAD"]) or "ACTIVIDAD"
    matriz_col = _find_column(matriz_df, MATRIZ_COL_ALIASES["MATRIZ_TERRENO"]) or "MATRIZ TERRENO"
    estado_col = _find_column(cierre_df, CIERRE_COL_ALIASES["ESTADO"]) or "ESTADO"
    keyword_col = _find_column(cierre_df, CIERRE_COL_ALIASES["PALABRA_CLAVE_COMENTARIO"]) or "PALABRA CLAVE COMENTARIO"
    codigo_col = _find_column(cierre_df, CIERRE_COL_ALIASES["CODIGO_CIERRE"]) or "CÓDIGO DE CIERRE"

    for col in [actividad_col, matriz_col]:
        if col not in matriz_df.columns:
            matriz_df[col] = ""
    for col in [estado_col, keyword_col, codigo_col]:
        if col not in cierre_df.columns:
            cierre_df[col] = ""

    existing_acts = {limpiar_texto(v) for v in matriz_df[actividad_col].tolist() if safe_str(v)}
    existing_pairs = {
        (limpiar_texto(r.get(estado_col)), limpiar_texto(r.get(keyword_col)))
        for _, r in cierre_df.iterrows()
        if safe_str(r.get(estado_col)) or safe_str(r.get(keyword_col))
    }

    add_acts = 0
    add_comments = 0

    for actividad in _dedup_strings(nuevas_actividades):
        key = limpiar_texto(actividad)
        if not key or key in existing_acts:
            continue
        matriz_df = pd.concat(
            [
                matriz_df,
                pd.DataFrame([{actividad_col: actividad, matriz_col: ""}]),
            ],
            ignore_index=True,
        )
        existing_acts.add(key)
        add_acts += 1

    for item in _dedup_comment_rows(comentarios_no_clasificados):
        estado = safe_str(item.get("ESTADO"))
        keyword = safe_str(item.get("COMENTARIO"))
        pair = (limpiar_texto(estado), limpiar_texto(keyword))
        if pair in existing_pairs or not (pair[0] or pair[1]):
            continue
        cierre_df = pd.concat(
            [
                cierre_df,
                pd.DataFrame(
                    [{estado_col: estado, keyword_col: keyword, codigo_col: ""}]
                ),
            ],
            ignore_index=True,
        )
        existing_pairs.add(pair)
        add_comments += 1

    if add_acts == 0 and add_comments == 0:
        return {"actividades_agregadas": 0, "comentarios_agregados": 0}

    sheets[matriz_sheet_name] = matriz_df
    sheets[cierre_sheet_name] = cierre_df

    with pd.ExcelWriter(source, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return {
        "actividades_agregadas": add_acts,
        "comentarios_agregados": add_comments,
    }


def _detectar_columna_actividad(registros_df: pd.DataFrame) -> str | None:
    norm_map = {normalize_column_name(col): col for col in registros_df.columns}
    for alias in ACTIVIDAD_INPUT_ALIASES:
        detected = norm_map.get(normalize_column_name(alias))
        if detected is not None:
            return detected
    return None


def clasificar_registros(
    registros_df: pd.DataFrame,
    path_matriz_clasificacion: Path | str,
    modo_nuevos_valores: str = "automatico",
    export_dir: Path | str | None = None,
    actualizar_diccionario_fuente: bool = False,
) -> pd.DataFrame:
    print("[4/5] Clasificando registros...")
    validate_required_columns(
        registros_df, ["PUNTO", "TIPO", "DIA", "ESTADO", "COMENTARIO"], "registros turno"
    )

    try:
        diccionarios = cargar_diccionarios(
            path_matriz_clasificacion,
            cache_dir=export_dir,
        )
    except Exception as exc:
        raise RuntimeError(
            "No fue posible cargar diccionarios. Verifica el archivo Excel de diccionarios."
        ) from exc

    dic_matriz: Dict[str, str] = diccionarios.get("MATRIZ", {})
    dic_cierre: Dict[str, List[Dict[str, str]]] = diccionarios.get("CIERRE", {})

    activity_col = _detectar_columna_actividad(registros_df)
    if activity_col is None:
        print(
            "[ADVERTENCIA] No se detecto columna de actividad/tarea en registros. "
            "Se intentara con columna TIPO."
        )
        activity_col = "TIPO"

    classified = registros_df.copy()
    codigos: List[str] = []
    matrices: List[str] = []

    nuevas_actividades: List[str] = []
    comentarios_no_clasificados: List[Dict[str, str]] = []

    for _, row in classified.iterrows():
        actividad = row.get(activity_col, "")
        estado = row.get("ESTADO", "")
        comentario = row.get("COMENTARIO", "")

        matriz = clasificar_matriz(
            actividad_tarea=actividad,
            diccionario_matriz=dic_matriz,
            nuevas_actividades=nuevas_actividades,
        )
        codigo = clasificar_cierre(
            estado=estado,
            comentario=comentario,
            diccionario_cierre=dic_cierre,
            comentarios_no_clasificados=comentarios_no_clasificados,
        )

        matrices.append(matriz)
        codigos.append(codigo)

    classified["MATRIZ_TERRENO"] = matrices
    classified["CODIGO_CIERRE"] = codigos

    novedades = detectar_nuevos_valores(
        nuevas_actividades=nuevas_actividades,
        comentarios_no_clasificados=comentarios_no_clasificados,
        modo=modo_nuevos_valores,
        export_dir=export_dir,
        diccionarios=diccionarios,
    )

    if actualizar_diccionario_fuente and (
        novedades["nuevas_actividades"] or novedades["comentarios_no_clasificados"]
    ):
        try:
            resumen_update = _actualizar_diccionario_fuente(
                path_diccionario_excel=path_matriz_clasificacion,
                nuevas_actividades=novedades["nuevas_actividades"],
                comentarios_no_clasificados=novedades["comentarios_no_clasificados"],
            )
            if resumen_update["actividades_agregadas"] or resumen_update["comentarios_agregados"]:
                print(
                    "[4/5] Diccionario fuente actualizado: "
                    f"{resumen_update['actividades_agregadas']} actividades nuevas y "
                    f"{resumen_update['comentarios_agregados']} comentarios clave nuevos."
                )
        except Exception as exc:
            print(
                "[ADVERTENCIA] No se pudo actualizar el diccionario fuente: "
                f"{exc}"
            )

    actividades_contenido = classified[activity_col].map(safe_str).str.strip() != ""
    comentarios_contenido = classified["COMENTARIO"].map(safe_str).str.strip() != ""
    missing_matriz = int((actividades_contenido & (classified["MATRIZ_TERRENO"].map(safe_str) == "")).sum())
    missing_cierre = int((comentarios_contenido & (classified["CODIGO_CIERRE"].map(safe_str) == "")).sum())

    if missing_matriz:
        print(
            "[ADVERTENCIA] Hay actividades sin matriz asignada: "
            f"{missing_matriz}. Revisa nuevas actividades detectadas."
        )
    if missing_cierre:
        print(
            "[ADVERTENCIA] Hay comentarios sin codigo de cierre: "
            f"{missing_cierre}. Revisa comentarios no clasificados."
        )

    print(
        f"[4/5] Registros clasificados con codigo: "
        f"{(classified['CODIGO_CIERRE'].astype(str).str.strip() != '').sum()} "
        f"de {len(classified)}"
    )
    if novedades["nuevas_actividades"] or novedades["comentarios_no_clasificados"]:
        print(
            "[4/5] Novedades detectadas: "
            f"{len(novedades['nuevas_actividades'])} actividades nuevas, "
            f"{len(novedades['comentarios_no_clasificados'])} comentarios no clasificados."
        )
    return classified
