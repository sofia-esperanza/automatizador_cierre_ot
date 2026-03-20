from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Dict

from procesos.actualizar_programa_mensual import actualizar_programa_mensual
from procesos.clasificar_cierre_ot import clasificar_registros
from procesos.generar_cierre_ot import actualizar_cierre_ot, generar_cierre_ot_base
from procesos.generar_cierre_turno_desde_msewjo import (
    generar_cierre_turno_desde_msewjo,
    get_turn_window,
)
from procesos.limpieza_msewjo import limpiar_msewjo
from utils.excel_utils import read_excel_file, save_dataframe_to_excel

TEMP_DIRNAME = "_temp"
ETAPA_1_DIRNAME = "etapa_1_limpieza_base"
ETAPA_2_DIRNAME = "etapa_2_actualizacion_mensual"
ETAPA_3_DIRNAME = "etapa_3_clasificacion"


def _stage_dir(carpeta_salida: str | Path, stage_dirname: str) -> Path:
    stage_dir = Path(carpeta_salida) / TEMP_DIRNAME / stage_dirname
    stage_dir.mkdir(parents=True, exist_ok=True)
    return stage_dir


def ejecutar_etapa_1_limpieza_base(
    ruta_msewjo: str | Path,
    carpeta_salida: str | Path,
    fecha_ancla_turno: dt.date | None = None,
) -> Dict[str, Path]:
    print("[Etapa 1] Cierre OT base desde MSEWJO...")
    stage_dir = _stage_dir(carpeta_salida, ETAPA_1_DIRNAME)

    if fecha_ancla_turno is None:
        fecha_ancla_turno = dt.date.today()
    turn_start, turn_end = get_turn_window(fecha_ancla_turno)
    cierre_turno_filename = (
        f"Cierre de OT Turno {turn_start.strftime('%d.%m')} al {turn_end.strftime('%d.%m')}.xlsx"
    )
    cierre_turno_path = stage_dir / cierre_turno_filename
    generar_cierre_turno_desde_msewjo(
        path_msewjo=ruta_msewjo,
        output_path=cierre_turno_path,
        turn_start=turn_start,
        turn_end=turn_end,
        include_sin_fecha=True,
    )

    # Archivos tecnicos para etapas posteriores (clasificacion/merge).
    df_limpio = limpiar_msewjo(ruta_msewjo)
    msewjo_limpio_path = save_dataframe_to_excel(
        df_limpio, stage_dir / "msewjo_limpio_tecnico.xlsx"
    )

    df_cierre_base = generar_cierre_ot_base(df_limpio)
    cierre_base_path = save_dataframe_to_excel(
        df_cierre_base, stage_dir / "cierre_ot_base_tecnico.xlsx"
    )

    return {
        "cierre_ot_turno_base": cierre_turno_path,
        "msewjo_limpio_tecnico": msewjo_limpio_path,
        "cierre_ot_base_tecnico": cierre_base_path,
    }


def ejecutar_etapa_2_actualizar_mensual(
    ruta_programa_turno: str | Path,
    ruta_programa_mensual: str | Path,
    carpeta_salida: str | Path,
) -> Dict[str, Path]:
    print("[Etapa 2] Actualizacion programa mensual...")
    stage_dir = _stage_dir(carpeta_salida, ETAPA_2_DIRNAME)

    mensual_actualizado_path = stage_dir / "programa_mensual_actualizado.xlsx"
    df_turno_aplicado = actualizar_programa_mensual(
        ruta_programa_turno, ruta_programa_mensual, mensual_actualizado_path
    )
    registros_turno_aplicado_path = save_dataframe_to_excel(
        df_turno_aplicado, stage_dir / "registros_turno_aplicado.xlsx"
    )

    return {
        "programa_mensual_actualizado": mensual_actualizado_path,
        "registros_turno_aplicado": registros_turno_aplicado_path,
    }


def ejecutar_etapa_3_clasificacion(
    ruta_matriz_clasificacion: str | Path,
    carpeta_salida: str | Path,
    ruta_cierre_base: str | Path | None = None,
    ruta_registros_turno_aplicado: str | Path | None = None,
) -> Dict[str, Path]:
    print("[Etapa 3] Clasificacion + cierre final...")
    stage_dir = _stage_dir(carpeta_salida, ETAPA_3_DIRNAME)

    if ruta_cierre_base is None:
        ruta_cierre_base = (
            Path(carpeta_salida) / TEMP_DIRNAME / ETAPA_1_DIRNAME / "cierre_ot_base_tecnico.xlsx"
        )
    if ruta_registros_turno_aplicado is None:
        ruta_registros_turno_aplicado = (
            Path(carpeta_salida)
            / TEMP_DIRNAME
            / ETAPA_2_DIRNAME
            / "registros_turno_aplicado.xlsx"
        )

    ruta_cierre_base = Path(ruta_cierre_base)
    ruta_registros_turno_aplicado = Path(ruta_registros_turno_aplicado)

    if not ruta_cierre_base.exists():
        raise FileNotFoundError(
            f"No existe cierre base para etapa 3: {ruta_cierre_base}. Ejecuta etapa 1 primero."
        )
    if not ruta_registros_turno_aplicado.exists():
        raise FileNotFoundError(
            "No existe registros_turno_aplicado para etapa 3: "
            f"{ruta_registros_turno_aplicado}. Ejecuta etapa 2 primero."
        )

    df_cierre_base = read_excel_file(ruta_cierre_base)
    df_turno_aplicado = read_excel_file(ruta_registros_turno_aplicado)

    df_clasificados = clasificar_registros(df_turno_aplicado, ruta_matriz_clasificacion)
    clasificados_path = save_dataframe_to_excel(
        df_clasificados, stage_dir / "registros_clasificados.xlsx"
    )

    df_cierre_final = actualizar_cierre_ot(df_cierre_base, df_clasificados)
    cierre_final_path = save_dataframe_to_excel(df_cierre_final, stage_dir / "cierre_ot_final.xlsx")

    return {
        "registros_clasificados": clasificados_path,
        "cierre_ot_final": cierre_final_path,
    }


def ejecutar_limpieza_y_base(
    ruta_msewjo: str | Path,
    carpeta_salida: str | Path,
) -> Dict[str, Path]:
    return ejecutar_etapa_1_limpieza_base(ruta_msewjo, carpeta_salida)


def ejecutar_flujo(
    ruta_msewjo: str | Path,
    ruta_programa_turno: str | Path,
    ruta_programa_mensual: str | Path,
    ruta_matriz_clasificacion: str | Path,
    carpeta_salida: str | Path,
) -> Dict[str, Path]:
    print("Iniciando procesamiento completo...")
    etapa_1 = ejecutar_etapa_1_limpieza_base(ruta_msewjo, carpeta_salida)
    etapa_2 = ejecutar_etapa_2_actualizar_mensual(
        ruta_programa_turno, ruta_programa_mensual, carpeta_salida
    )
    etapa_3 = ejecutar_etapa_3_clasificacion(
        ruta_matriz_clasificacion=ruta_matriz_clasificacion,
        carpeta_salida=carpeta_salida,
        ruta_cierre_base=etapa_1["cierre_ot_base_tecnico"],
        ruta_registros_turno_aplicado=etapa_2["registros_turno_aplicado"],
    )

    result = {**etapa_1, **etapa_2, **etapa_3}
    print("Proceso completo finalizado.")
    for key, value in result.items():
        print(f"- {key}: {value}")
    return result


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Automatizador de cierre OT para monitoreo ambiental."
    )
    parser.add_argument("--msewjo", required=True, help="Ruta archivo MSEWJO")
    parser.add_argument("--turno", required=True, help="Ruta programa de monitoreo por turno")
    parser.add_argument("--mensual", required=True, help="Ruta programa de monitoreo mensual")
    parser.add_argument("--matriz", required=True, help="Ruta matriz de clasificacion de cierre")
    parser.add_argument("--output", required=True, help="Carpeta de salida")
    return parser


if __name__ == "__main__":
    args = _build_parser().parse_args()
    ejecutar_flujo(
        ruta_msewjo=args.msewjo,
        ruta_programa_turno=args.turno,
        ruta_programa_mensual=args.mensual,
        ruta_matriz_clasificacion=args.matriz,
        carpeta_salida=args.output,
    )
