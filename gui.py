from __future__ import annotations

import calendar
import datetime as dt
import json
from pathlib import Path
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from main import (
    ETAPA_1_DIRNAME,
    ETAPA_2_DIRNAME,
    TEMP_DIRNAME,
    ejecutar_etapa_1_limpieza_base,
    ejecutar_etapa_2_actualizar_mensual,
    ejecutar_etapa_3_clasificacion,
    ejecutar_flujo,
)

ETAPAS = {
    "Etapa 1 - Cierre OT base (desde MSEWJO)": "etapa_1",
    "Etapa 2 - Actualizar mensual": "etapa_2",
    "Etapa 3 - Clasificacion final": "etapa_3",
    "Flujo completo (1+2+3)": "completo",
}

PALETA = {
    "primario": "#00B0AA",
    "primario_oscuro": "#297a76",
    "primario_claro": "#9ce2dd",
    "superficie": "#ffffff",
    "gris_medio": "#b8b8b8",
    "gris_oscuro": "#666666",
    "texto": "#000000",
}

CAMPO_META = {
    "msewjo": {"label": "Archivo MSEWJO", "selector": "file"},
    "turno": {"label": "Programa Turno", "selector": "file"},
    "mensual": {"label": "Programa Mensual", "selector": "file"},
    "fecha_desde": {"label": "Fecha desde", "selector": "date"},
    "fecha_hasta": {"label": "Fecha hasta", "selector": "date"},
    "output": {"label": "Carpeta Salida", "selector": "dir"},
}

CONFIG_DIR = Path.home() / ".automatizador_cierre_ot"
CONFIG_PATH = CONFIG_DIR / "config_gui.json"


class DatePickerDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Widget,
        initial_date: dt.date,
        on_select: callable,
    ) -> None:
        super().__init__(parent)
        self.title("Seleccionar fecha")
        self.resizable(False, False)
        self.configure(bg=PALETA["superficie"])
        self.transient(parent)
        self.grab_set()

        self._on_select_callback = on_select
        self.current_year = initial_date.year
        self.current_month = initial_date.month

        container = tk.Frame(self, bg=PALETA["superficie"])
        container.pack(padx=10, pady=10)

        header = tk.Frame(container, bg=PALETA["superficie"])
        header.pack(fill="x")

        tk.Button(
            header,
            text="<",
            width=3,
            command=lambda: self._change_month(-1),
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
        ).pack(side="left")

        self.month_label_var = tk.StringVar()
        tk.Label(
            header,
            textvariable=self.month_label_var,
            bg=PALETA["superficie"],
            fg=PALETA["texto"],
            font=("Segoe UI", 10, "bold"),
            width=18,
            anchor="center",
        ).pack(side="left", padx=6)

        tk.Button(
            header,
            text=">",
            width=3,
            command=lambda: self._change_month(1),
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
        ).pack(side="left")

        weekdays = tk.Frame(container, bg=PALETA["superficie"])
        weekdays.pack(pady=(8, 4))
        for idx, day_name in enumerate(["L", "M", "X", "J", "V", "S", "D"]):
            tk.Label(
                weekdays,
                text=day_name,
                width=4,
                bg=PALETA["superficie"],
                fg=PALETA["gris_oscuro"],
                font=("Segoe UI", 9, "bold"),
            ).grid(row=0, column=idx, padx=1, pady=1)

        self.days_frame = tk.Frame(container, bg=PALETA["superficie"])
        self.days_frame.pack()

        actions = tk.Frame(container, bg=PALETA["superficie"])
        actions.pack(fill="x", pady=(8, 0))

        tk.Button(
            actions,
            text="Hoy",
            width=8,
            command=self._select_today,
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 9),
            cursor="hand2",
        ).pack(side="left")

        tk.Button(
            actions,
            text="Cerrar",
            width=8,
            command=self.destroy,
            bg=PALETA["superficie"],
            fg=PALETA["gris_oscuro"],
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 9),
            cursor="hand2",
        ).pack(side="right")

        self._render_days()

    def _month_title(self) -> str:
        month_name = {
            1: "Enero",
            2: "Febrero",
            3: "Marzo",
            4: "Abril",
            5: "Mayo",
            6: "Junio",
            7: "Julio",
            8: "Agosto",
            9: "Septiembre",
            10: "Octubre",
            11: "Noviembre",
            12: "Diciembre",
        }[self.current_month]
        return f"{month_name} {self.current_year}"

    def _change_month(self, delta: int) -> None:
        year = self.current_year
        month = self.current_month + delta
        if month < 1:
            month = 12
            year -= 1
        elif month > 12:
            month = 1
            year += 1
        self.current_year = year
        self.current_month = month
        self._render_days()

    def _render_days(self) -> None:
        self.month_label_var.set(self._month_title())
        for child in self.days_frame.winfo_children():
            child.destroy()

        cal = calendar.Calendar(firstweekday=0)
        weeks = cal.monthdayscalendar(self.current_year, self.current_month)
        for r, week in enumerate(weeks):
            for c, day in enumerate(week):
                if day == 0:
                    tk.Label(
                        self.days_frame,
                        text="",
                        width=4,
                        bg=PALETA["superficie"],
                    ).grid(row=r, column=c, padx=1, pady=1)
                    continue

                tk.Button(
                    self.days_frame,
                    text=str(day),
                    width=4,
                    command=lambda d=day: self._pick_day(d),
                    bg=PALETA["superficie"],
                    fg=PALETA["texto"],
                    activebackground=PALETA["primario_claro"],
                    relief=tk.SOLID,
                    bd=1,
                    font=("Segoe UI", 9),
                    cursor="hand2",
                ).grid(row=r, column=c, padx=1, pady=1)

    def _pick_day(self, day: int) -> None:
        selected = dt.date(self.current_year, self.current_month, day)
        self._on_select_callback(selected)
        self.destroy()

    def _select_today(self) -> None:
        today = dt.date.today()
        self._on_select_callback(today)
        self.destroy()


class AutomatizadorGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Automatizador Cierre OT")
        self.geometry("980x540")
        self.resizable(False, False)
        self.configure(bg=PALETA["superficie"])

        diccionario_guardado = self._load_saved_diccionario_path()
        self.vars = {
            "msewjo": tk.StringVar(),
            "turno": tk.StringVar(),
            "mensual": tk.StringVar(),
            "fecha_desde": tk.StringVar(),
            "fecha_hasta": tk.StringVar(),
            "matriz": tk.StringVar(value=diccionario_guardado),
            "output": tk.StringVar(value=str(Path("output").resolve())),
        }
        first_stage = next(iter(ETAPAS.keys()))
        self.stage_label_var = tk.StringVar(value=first_stage)
        self.requirements_var = tk.StringVar(value="")
        self.diccionario_badge_var = tk.StringVar(value="Ultima actualizacion: --/--/--")
        self.status_var = tk.StringVar(
            value="Selecciona etapa, archivos y presiona PROCESAR."
        )
        self.field_rows: dict[str, tk.Frame] = {}
        self.date_buttons: dict[str, tk.Button] = {}
        self.process_btn: tk.Button

        self._build_ui()
        self.stage_label_var.trace_add("write", self._on_stage_change)
        self._refresh_fields_visibility()

    def _build_ui(self) -> None:
        root_container = tk.Frame(self, bg=PALETA["superficie"])
        root_container.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Frame(root_container, bg=PALETA["primario_oscuro"], height=6).pack(fill="x", pady=(0, 12))

        title_row = tk.Frame(root_container, bg=PALETA["superficie"])
        title_row.pack(fill="x", pady=(0, 8))

        title_container = tk.Frame(title_row, bg=PALETA["superficie"])
        title_container.pack(side="left", fill="x", expand=True)

        title = tk.Label(
            title_container,
            text="Automatizador Cierre OT",
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            font=("Segoe UI", 20, "bold"),
            anchor="w",
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            title_container,
            text="Selecciona una etapa y adjunta solo los archivos requeridos",
            bg=PALETA["superficie"],
            fg=PALETA["gris_oscuro"],
            font=("Segoe UI", 11),
            anchor="w",
        )
        subtitle.pack(anchor="w", pady=(6, 0))

        diccionario_panel = tk.Frame(title_row, bg=PALETA["superficie"])
        diccionario_panel.pack(side="right", padx=(12, 2))

        tk.Button(
            diccionario_panel,
            text="↻",
            command=self._reset_inputs,
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            activebackground=PALETA["superficie"],
            activeforeground=PALETA["primario_oscuro"],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            padx=2,
            pady=0,
            font=("Segoe UI", 16, "bold"),
            cursor="hand2",
        ).pack(anchor="e", pady=(0, 6))

        tk.Button(
            diccionario_panel,
            text="Diccionarios",
            width=16,
            command=self._select_diccionario,
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            activebackground=PALETA["primario"],
            activeforeground=PALETA["superficie"],
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        ).pack(anchor="e")

        tk.Label(
            diccionario_panel,
            textvariable=self.diccionario_badge_var,
            bg=PALETA["superficie"],
            fg=PALETA["gris_oscuro"],
            font=("Segoe UI", 9),
            anchor="e",
            justify="right",
        ).pack(anchor="e", pady=(4, 0))

        card = tk.Frame(
            root_container,
            bg=PALETA["superficie"],
            highlightbackground=PALETA["gris_medio"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(fill="both", expand=True)

        header = tk.Frame(card, bg=PALETA["superficie"])
        header.pack(fill="x", padx=16, pady=(14, 4))

        tk.Label(
            header,
            text="Etapa",
            width=16,
            anchor="w",
            bg=PALETA["superficie"],
            fg=PALETA["gris_oscuro"],
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=(8, 4), pady=8)

        stage_menu = tk.OptionMenu(header, self.stage_label_var, *ETAPAS.keys())
        stage_menu.config(
            width=42,
            bg=PALETA["superficie"],
            fg=PALETA["texto"],
            activebackground=PALETA["superficie"],
            activeforeground=PALETA["texto"],
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightbackground=PALETA["gris_medio"],
            font=("Segoe UI", 10),
        )
        stage_menu["menu"].config(
            bg=PALETA["superficie"],
            fg=PALETA["texto"],
            activebackground=PALETA["primario_claro"],
            activeforeground=PALETA["texto"],
            font=("Segoe UI", 10),
        )
        stage_menu.grid(row=0, column=1, sticky="w", padx=(4, 8), pady=8)

        tk.Label(
            header,
            textvariable=self.requirements_var,
            anchor="w",
            justify="left",
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            font=("Segoe UI", 10, "bold"),
            wraplength=860,
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=(8, 8), pady=(0, 8))

        tk.Frame(card, bg=PALETA["gris_medio"], height=1).pack(fill="x", padx=16, pady=(0, 8))

        fields_container = tk.Frame(card, bg=PALETA["superficie"])
        fields_container.pack(fill="both", expand=True, padx=18, pady=(2, 8))

        field_order = ["msewjo", "turno", "mensual", "fecha_desde", "fecha_hasta", "output"]
        for idx, key in enumerate(field_order):
            if key == "fecha_hasta":
                continue

            if key == "fecha_desde":
                row_frame = tk.Frame(fields_container, bg=PALETA["superficie"])
                row_frame.grid(row=idx, column=0, sticky="we", pady=6)
                self.field_rows["fecha_desde"] = row_frame
                self.field_rows["fecha_hasta"] = row_frame

                tk.Label(
                    row_frame,
                    text=CAMPO_META["fecha_desde"]["label"],
                    width=18,
                    anchor="w",
                    bg=PALETA["superficie"],
                    fg=PALETA["gris_oscuro"],
                    font=("Segoe UI", 10, "bold"),
                ).grid(row=0, column=0, sticky="w")

                btn_desde = tk.Button(
                    row_frame,
                    text=self._calendar_button_text("fecha_desde"),
                    width=12,
                    command=lambda: self._open_calendar_for("fecha_desde"),
                    bg=PALETA["superficie"],
                    fg=PALETA["primario_oscuro"],
                    activebackground=PALETA["primario"],
                    activeforeground=PALETA["superficie"],
                    relief=tk.SOLID,
                    bd=1,
                    font=("Segoe UI", 10),
                    cursor="hand2",
                )
                btn_desde.grid(row=0, column=1, padx=(8, 24), sticky="w")
                self.date_buttons["fecha_desde"] = btn_desde

                tk.Label(
                    row_frame,
                    text=CAMPO_META["fecha_hasta"]["label"],
                    width=12,
                    anchor="w",
                    bg=PALETA["superficie"],
                    fg=PALETA["gris_oscuro"],
                    font=("Segoe UI", 10, "bold"),
                ).grid(row=0, column=2, sticky="w")

                btn_hasta = tk.Button(
                    row_frame,
                    text=self._calendar_button_text("fecha_hasta"),
                    width=12,
                    command=lambda: self._open_calendar_for("fecha_hasta"),
                    bg=PALETA["superficie"],
                    fg=PALETA["primario_oscuro"],
                    activebackground=PALETA["primario"],
                    activeforeground=PALETA["superficie"],
                    relief=tk.SOLID,
                    bd=1,
                    font=("Segoe UI", 10),
                    cursor="hand2",
                )
                btn_hasta.grid(row=0, column=3, padx=(8, 0), sticky="w")
                self.date_buttons["fecha_hasta"] = btn_hasta
                continue

            meta = CAMPO_META[key]
            row_frame = tk.Frame(fields_container, bg=PALETA["superficie"])
            row_frame.grid(row=idx, column=0, sticky="we", pady=6)
            self.field_rows[key] = row_frame

            tk.Label(
                row_frame,
                text=meta["label"],
                width=18,
                anchor="w",
                bg=PALETA["superficie"],
                fg=PALETA["gris_oscuro"],
                font=("Segoe UI", 10, "bold"),
            ).grid(row=0, column=0, sticky="w")

            selector = meta.get("selector", "none")
            if selector in {"file", "dir", "none"}:
                entry = tk.Entry(
                    row_frame,
                    textvariable=self.vars[key],
                    width=66,
                    relief=tk.SOLID,
                    bd=1,
                    highlightthickness=1,
                    highlightbackground=PALETA["gris_medio"],
                    highlightcolor=PALETA["primario_oscuro"],
                    font=("Segoe UI", 10),
                    bg=PALETA["superficie"],
                    fg=PALETA["texto"],
                    insertbackground=PALETA["texto"],
                )
                entry.grid(row=0, column=1, padx=8, sticky="we")

            if selector in {"file", "dir"}:
                handler = self._select_file if selector == "file" else self._select_dir
                tk.Button(
                    row_frame,
                    text="Buscar",
                    width=12,
                    command=lambda k=key, h=handler: h(k),
                    bg=PALETA["superficie"],
                    fg=PALETA["primario_oscuro"],
                    activebackground=PALETA["primario"],
                    activeforeground=PALETA["superficie"],
                    relief=tk.SOLID,
                    bd=1,
                    font=("Segoe UI", 10),
                    cursor="hand2",
                ).grid(row=0, column=2, padx=(2, 0))
            elif selector == "date":
                button = tk.Button(
                    row_frame,
                    text=self._calendar_button_text(key),
                    width=12,
                    command=lambda k=key: self._open_calendar_for(k),
                    bg=PALETA["superficie"],
                    fg=PALETA["primario_oscuro"],
                    activebackground=PALETA["primario"],
                    activeforeground=PALETA["superficie"],
                    relief=tk.SOLID,
                    bd=1,
                    font=("Segoe UI", 10),
                    cursor="hand2",
                )
                button.grid(row=0, column=1, padx=(8, 0), sticky="w")
                self.date_buttons[key] = button

            row_frame.grid_columnconfigure(1, weight=1)

        actions = tk.Frame(card, bg=PALETA["superficie"])
        actions.pack(fill="x", padx=18, pady=(0, 4))

        self.process_btn = tk.Button(
            actions,
            text="PROCESAR ETAPA",
            width=24,
            height=2,
            command=self._process,
            bg=PALETA["primario_oscuro"],
            fg=PALETA["superficie"],
            activebackground=PALETA["primario"],
            activeforeground=PALETA["superficie"],
            relief=tk.SOLID,
            bd=1,
            font=("Segoe UI", 11, "bold"),
            cursor="hand2",
        )
        self.process_btn.pack(pady=10)

        status = tk.Label(
            card,
            textvariable=self.status_var,
            anchor="w",
            justify="left",
            bg=PALETA["superficie"],
            fg=PALETA["gris_oscuro"],
            font=("Segoe UI", 10),
            wraplength=910,
        )
        status.pack(fill="x", padx=18, pady=(6, 2))

        tk.Label(
            card,
            text="Archivos temporales por etapa: <salida>/_temp/etapa_*",
            anchor="w",
            justify="left",
            bg=PALETA["superficie"],
            fg=PALETA["gris_oscuro"],
            font=("Segoe UI", 10),
        ).pack(fill="x", padx=18, pady=(0, 12))

    def _select_file(self, key: str) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
        )
        if path:
            self.vars[key].set(path)

    def _select_diccionario(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar diccionario Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls")],
        )
        if path:
            self.vars["matriz"].set(path)
            self._save_saved_diccionario_path(path)
            self._refresh_diccionario_badge()

    def _select_dir(self, key: str) -> None:
        path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if path:
            self.vars[key].set(path)

    def _open_calendar_for(self, key: str) -> None:
        initial_date = self._parse_date_optional(self.vars[key].get(), CAMPO_META[key]["label"])
        if initial_date is None:
            initial_date = dt.date.today()

        def _set_selected(value: dt.date) -> None:
            self.vars[key].set(value.strftime("%d/%m/%Y"))
            self._refresh_calendar_button_labels()

        DatePickerDialog(self, initial_date=initial_date, on_select=_set_selected)

    def _calendar_button_text(self, key: str) -> str:
        value = self.vars[key].get().strip()
        return value if value else "Calendario"

    def _refresh_calendar_button_labels(self) -> None:
        for key, button in self.date_buttons.items():
            button.config(text=self._calendar_button_text(key))

    def _selected_stage(self) -> str:
        return ETAPAS[self.stage_label_var.get()]

    def _load_saved_diccionario_path(self) -> str:
        try:
            if not CONFIG_PATH.exists():
                return ""
            payload = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if not isinstance(payload, dict):
                return ""
            value = payload.get("ruta_diccionario", "")
            return str(value).strip()
        except Exception:
            return ""

    def _save_saved_diccionario_path(self, path: str) -> None:
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            payload = {"ruta_diccionario": str(path).strip()}
            CONFIG_PATH.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            # No bloquea la UI si no puede guardar configuracion.
            pass

    def _refresh_diccionario_badge(self) -> None:
        dic_path = self.vars["matriz"].get().strip()
        if not dic_path:
            self.diccionario_badge_var.set("Ultima actualizacion: --/--/--")
            return
        p = Path(dic_path)
        if not p.exists():
            self.diccionario_badge_var.set("Ultima actualizacion: archivo no encontrado")
            return
        date_str = dt.datetime.fromtimestamp(p.stat().st_mtime).strftime("%d/%m/%y")
        self.diccionario_badge_var.set(f"Ultima actualizacion: {date_str}")

    def _required_files_for_stage(self, stage: str) -> list[str]:
        if stage == "etapa_1":
            return ["msewjo"]
        if stage == "etapa_2":
            return ["turno", "mensual"]
        if stage == "etapa_3":
            return []
        return ["msewjo", "turno", "mensual"]

    def _optional_fields_for_stage(self, stage: str) -> list[str]:
        if stage in {"etapa_2", "completo"}:
            return ["fecha_desde", "fecha_hasta"]
        return []

    def _on_stage_change(self, *args: object) -> None:
        self._refresh_fields_visibility()

    def _refresh_fields_visibility(self) -> None:
        stage = self._selected_stage()
        required_keys = set(self._required_files_for_stage(stage))
        optional_keys = set(self._optional_fields_for_stage(stage))
        visible_keys = required_keys | optional_keys | {"output"}

        for key, row in self.field_rows.items():
            if key in visible_keys:
                row.grid()
            else:
                row.grid_remove()

        if stage == "completo":
            self.process_btn.config(text="PROCESAR FLUJO COMPLETO")
        else:
            self.process_btn.config(text="PROCESAR ETAPA")

        required_labels = ", ".join(
            CAMPO_META[k]["label"] for k in self._required_files_for_stage(stage)
        )
        if stage in {"etapa_1", "etapa_2", "etapa_3", "completo"}:
            required_labels = (
                f"{required_labels}, Diccionarios (boton superior)"
                if required_labels
                else "Diccionarios (boton superior)"
            )
        if stage in {"etapa_2", "completo"}:
            required_labels = (
                f"{required_labels}, Rango semanal opcional (dd/mm/aaaa)"
                if required_labels
                else "Rango semanal opcional (dd/mm/aaaa)"
            )
        self.requirements_var.set(f"Requeridos para esta etapa: {required_labels}.")
        self._refresh_diccionario_badge()
        self._refresh_calendar_button_labels()
        self.status_var.set("Listo para procesar.")

    def _parse_date_optional(self, value: str, field_label: str) -> dt.date | None:
        text = value.strip()
        if not text:
            return None
        try:
            return dt.datetime.strptime(text, "%d/%m/%Y").date()
        except ValueError as exc:
            raise ValueError(
                f"Fecha invalida en '{field_label}': '{text}'. Usa formato dd/mm/aaaa."
            ) from exc

    def _validate_inputs(self, stage: str) -> None:
        for key in self._required_files_for_stage(stage):
            value = self.vars[key].get().strip()
            if not value:
                raise ValueError(f"Falta seleccionar: {CAMPO_META[key]['label']}")
            if not Path(value).exists():
                raise FileNotFoundError(f"No existe el archivo: {value}")
            if key == "matriz":
                self._save_saved_diccionario_path(value)

        output = self.vars["output"].get().strip()
        if not output:
            raise ValueError("Falta seleccionar carpeta de salida.")

        if stage in {"etapa_1", "etapa_2", "etapa_3", "completo"}:
            diccionario = self.vars["matriz"].get().strip()
            if not diccionario:
                raise ValueError("Falta seleccionar Diccionarios (boton superior).")
            if not Path(diccionario).exists():
                raise FileNotFoundError(f"No existe archivo de diccionario: {diccionario}")
            self._save_saved_diccionario_path(diccionario)

        if stage == "etapa_3":
            base = Path(output) / TEMP_DIRNAME
            cierre_base = base / ETAPA_1_DIRNAME / "cierre_ot_base_tecnico.xlsx"
            turno_aplicado = base / ETAPA_2_DIRNAME / "registros_turno_aplicado.xlsx"
            if not cierre_base.exists():
                raise FileNotFoundError(f"No existe archivo previo: {cierre_base}")
            if not turno_aplicado.exists():
                raise FileNotFoundError(f"No existe archivo previo: {turno_aplicado}")

        if stage in {"etapa_2", "completo"}:
            fecha_desde = self._parse_date_optional(
                self.vars["fecha_desde"].get(), CAMPO_META["fecha_desde"]["label"]
            )
            fecha_hasta = self._parse_date_optional(
                self.vars["fecha_hasta"].get(), CAMPO_META["fecha_hasta"]["label"]
            )
            if fecha_desde is not None and fecha_hasta is None:
                fecha_hasta = fecha_desde
            if fecha_hasta is not None and fecha_desde is None:
                fecha_desde = fecha_hasta
            if (
                fecha_desde is not None
                and fecha_hasta is not None
                and fecha_desde > fecha_hasta
            ):
                raise ValueError("Fecha desde no puede ser mayor que fecha hasta.")

    def _format_result_lines(self, result: dict) -> str:
        lines = []
        for path in result.values():
            lines.append(f"- {path}")
        return "\n".join(lines)

    def _reset_inputs(self) -> None:
        for key in ["msewjo", "turno", "mensual", "fecha_desde", "fecha_hasta"]:
            self.vars[key].set("")
        self._refresh_fields_visibility()
        self.status_var.set("Formulario recargado. Selecciona nuevos archivos para procesar.")

    def _reset_after_success(self) -> None:
        for key in ["msewjo", "turno", "mensual", "fecha_desde", "fecha_hasta"]:
            self.vars[key].set("")
        self._refresh_fields_visibility()
        self.status_var.set("Listo para un nuevo procesamiento.")

    def _process(self) -> None:
        try:
            stage = self._selected_stage()
            self._validate_inputs(stage=stage)

            self.process_btn.config(state=tk.DISABLED)
            self.status_var.set(f"Procesando {self.stage_label_var.get()}...")
            self.update_idletasks()

            if stage == "etapa_1":
                result = ejecutar_etapa_1_limpieza_base(
                    ruta_msewjo=self.vars["msewjo"].get(),
                    ruta_matriz_clasificacion=self.vars["matriz"].get(),
                    carpeta_salida=self.vars["output"].get(),
                )
            elif stage == "etapa_2":
                result = ejecutar_etapa_2_actualizar_mensual(
                    ruta_programa_turno=self.vars["turno"].get(),
                    ruta_programa_mensual=self.vars["mensual"].get(),
                    ruta_matriz_clasificacion=self.vars["matriz"].get(),
                    carpeta_salida=self.vars["output"].get(),
                    fecha_desde=self.vars["fecha_desde"].get(),
                    fecha_hasta=self.vars["fecha_hasta"].get(),
                )
            elif stage == "etapa_3":
                result = ejecutar_etapa_3_clasificacion(
                    ruta_matriz_clasificacion=self.vars["matriz"].get(),
                    carpeta_salida=self.vars["output"].get(),
                )
            else:
                result = ejecutar_flujo(
                    ruta_msewjo=self.vars["msewjo"].get(),
                    ruta_programa_turno=self.vars["turno"].get(),
                    ruta_programa_mensual=self.vars["mensual"].get(),
                    ruta_matriz_clasificacion=self.vars["matriz"].get(),
                    carpeta_salida=self.vars["output"].get(),
                    fecha_desde=self.vars["fecha_desde"].get(),
                    fecha_hasta=self.vars["fecha_hasta"].get(),
                )

            self.status_var.set(f"{self.stage_label_var.get()} completada.")
            messagebox.showinfo(
                "Proceso completado",
                "Se generaron los siguientes archivos:\n\n"
                f"{self._format_result_lines(result)}",
            )
            self._reset_after_success()
        except Exception as exc:
            self.status_var.set("Error durante el procesamiento.")
            output_dir = Path(self.vars["output"].get().strip() or "output")
            error_log = output_dir / TEMP_DIRNAME / "last_error.log"
            error_log.parent.mkdir(parents=True, exist_ok=True)
            error_log.write_text(traceback.format_exc(), encoding="utf-8")
            messagebox.showerror("Error", f"{exc}\n\nDetalle: {error_log}")
        finally:
            self.process_btn.config(state=tk.NORMAL)


if __name__ == "__main__":
    app = AutomatizadorGUI()
    app.mainloop()
