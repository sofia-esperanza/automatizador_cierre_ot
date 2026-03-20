from __future__ import annotations

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
    "matriz": {"label": "Matriz Clasificacion", "selector": "file"},
    "output": {"label": "Carpeta Salida", "selector": "dir"},
}


class AutomatizadorGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Automatizador Cierre OT")
        self.geometry("980x540")
        self.resizable(False, False)
        self.configure(bg=PALETA["superficie"])

        self.vars = {
            "msewjo": tk.StringVar(),
            "turno": tk.StringVar(),
            "mensual": tk.StringVar(),
            "matriz": tk.StringVar(),
            "output": tk.StringVar(value=str(Path("output").resolve())),
        }
        first_stage = next(iter(ETAPAS.keys()))
        self.stage_label_var = tk.StringVar(value=first_stage)
        self.requirements_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(
            value="Selecciona etapa, archivos y presiona PROCESAR."
        )
        self.field_rows: dict[str, tk.Frame] = {}
        self.process_btn: tk.Button

        self._build_ui()
        self.stage_label_var.trace_add("write", self._on_stage_change)
        self._refresh_fields_visibility()

    def _build_ui(self) -> None:
        root_container = tk.Frame(self, bg=PALETA["superficie"])
        root_container.pack(fill="both", expand=True, padx=20, pady=18)

        tk.Frame(root_container, bg=PALETA["primario_oscuro"], height=6).pack(fill="x", pady=(0, 12))

        title = tk.Label(
            root_container,
            text="Automatizador Cierre OT",
            bg=PALETA["superficie"],
            fg=PALETA["primario_oscuro"],
            font=("Segoe UI", 20, "bold"),
            anchor="w",
        )
        title.pack(fill="x", pady=(0, 6))

        subtitle = tk.Label(
            root_container,
            text="Selecciona una etapa y adjunta solo los archivos requeridos",
            bg=PALETA["superficie"],
            fg=PALETA["gris_oscuro"],
            font=("Segoe UI", 11),
            anchor="w",
        )
        subtitle.pack(fill="x", pady=(0, 14))

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

        field_order = ["msewjo", "turno", "mensual", "matriz", "output"]
        for idx, key in enumerate(field_order):
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

            handler = self._select_file if meta["selector"] == "file" else self._select_dir
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

    def _select_dir(self, key: str) -> None:
        path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if path:
            self.vars[key].set(path)

    def _selected_stage(self) -> str:
        return ETAPAS[self.stage_label_var.get()]

    def _required_files_for_stage(self, stage: str) -> list[str]:
        if stage == "etapa_1":
            return ["msewjo"]
        if stage == "etapa_2":
            return ["turno", "mensual"]
        if stage == "etapa_3":
            return ["matriz"]
        return ["msewjo", "turno", "mensual", "matriz"]

    def _on_stage_change(self, *args: object) -> None:
        self._refresh_fields_visibility()

    def _refresh_fields_visibility(self) -> None:
        stage = self._selected_stage()
        required_keys = set(self._required_files_for_stage(stage))
        visible_keys = required_keys | {"output"}

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
        self.requirements_var.set(f"Requeridos para esta etapa: {required_labels}.")
        self.status_var.set("Listo para procesar.")

    def _validate_inputs(self, stage: str) -> None:
        for key in self._required_files_for_stage(stage):
            value = self.vars[key].get().strip()
            if not value:
                raise ValueError(f"Falta seleccionar: {CAMPO_META[key]['label']}")
            if not Path(value).exists():
                raise FileNotFoundError(f"No existe el archivo: {value}")

        output = self.vars["output"].get().strip()
        if not output:
            raise ValueError("Falta seleccionar carpeta de salida.")

        if stage == "etapa_3":
            base = Path(output) / TEMP_DIRNAME
            cierre_base = base / ETAPA_1_DIRNAME / "cierre_ot_base_tecnico.xlsx"
            turno_aplicado = base / ETAPA_2_DIRNAME / "registros_turno_aplicado.xlsx"
            if not cierre_base.exists():
                raise FileNotFoundError(f"No existe archivo previo: {cierre_base}")
            if not turno_aplicado.exists():
                raise FileNotFoundError(f"No existe archivo previo: {turno_aplicado}")

    def _format_result_lines(self, result: dict) -> str:
        lines = []
        for path in result.values():
            lines.append(f"- {path}")
        return "\n".join(lines)

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
                    carpeta_salida=self.vars["output"].get(),
                )
            elif stage == "etapa_2":
                result = ejecutar_etapa_2_actualizar_mensual(
                    ruta_programa_turno=self.vars["turno"].get(),
                    ruta_programa_mensual=self.vars["mensual"].get(),
                    carpeta_salida=self.vars["output"].get(),
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
                )

            self.status_var.set(f"{self.stage_label_var.get()} completada.")
            messagebox.showinfo(
                "Proceso completado",
                "Se generaron los siguientes archivos:\n\n"
                f"{self._format_result_lines(result)}",
            )
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
