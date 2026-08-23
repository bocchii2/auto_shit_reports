from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

import customtkinter as ctk

from app.services import storage
from app.services.csv_import import load_csv
from app.services.excel_generator import generate_excel
from app.services.filters import apply_filters
from app.services.report_code import (
    MES_NOMBRE_TITLE,
    build_codigo,
    build_periodo_corto,
    build_periodo_label,
    format_fecha_informe,
)
from app.paths import exports_dir, templates_dir
from app.services.word_generator import generate_word

TEMPLATES = templates_dir()
EXPORTS = exports_dir()

MESES = [(i, MES_NOMBRE_TITLE[i]) for i in range(1, 13)]


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.title("Generador de Informes de Actividades")
        self.geometry("1180x740")
        self.minsize(980, 640)

        storage.init_db()

        self.all_rows: list[dict[str, Any]] = []
        self.filtered_rows: list[dict[str, Any]] = []
        self.current_batch_id: int | None = None
        self.profiles: list[dict[str, Any]] = []
        self.current_profile_id: int | None = None
        self._profile_label_to_id: dict[str, int] = {}
        self._loading_profile = False

        today = date.today()
        # default period = previous month often used for monthly reports
        self.var_mes = tk.IntVar(value=today.month)
        self.var_anio = tk.IntVar(value=today.year)
        self.var_query = tk.StringVar()
        self.var_profile = tk.StringVar(value="(Sin perfil)")
        self.var_nombre = tk.StringVar()
        self.var_cargo = tk.StringVar()
        self.var_supervisor = tk.StringVar(value="Ing. Patricio Guzman")
        self.var_cargo_firma = tk.StringVar()
        self.var_cargo_supervisor = tk.StringVar(value="Supervisor Inmediato")
        self.var_empresa = tk.StringVar(value="Altura S.A.")
        self.var_equipo = tk.StringVar(value="IDS")
        self.var_ciudad = tk.StringVar(value="Manta")
        self.var_codigo = tk.StringVar()
        self.var_status = tk.StringVar(value="Listo. Importa un CSV para comenzar.")

        self._build_ui()
        self._refresh_codigo()
        self._load_latest_batch()
        self._reload_profiles(select_default=True)

    # ── UI construction ──────────────────────────────────────────────
    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 4))
        self.tabs.add("Actividades")
        self.tabs.add("Cabecera y generar")
        self.tabs.add("Historial")

        self._build_activities_tab(self.tabs.tab("Actividades"))
        self._build_header_tab(self.tabs.tab("Cabecera y generar"))
        self._build_history_tab(self.tabs.tab("Historial"))

        status = ctk.CTkLabel(self, textvariable=self.var_status, anchor="w")
        status.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

    def _build_activities_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(2, weight=1)

        top = ctk.CTkFrame(parent)
        top.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        top.grid_columnconfigure(6, weight=1)

        ctk.CTkButton(top, text="Importar CSV", command=self._import_csv, width=140).grid(
            row=0, column=0, padx=4, pady=4
        )
        ctk.CTkButton(top, text="Recargar batch", command=self._load_latest_batch, width=130).grid(
            row=0, column=1, padx=4, pady=4
        )

        ctk.CTkLabel(top, text="Mes").grid(row=0, column=2, padx=(16, 4))
        mes_values = [f"{i:02d} - {n}" for i, n in MESES]
        self.cmb_mes = ctk.CTkComboBox(
            top,
            values=mes_values,
            width=140,
            command=lambda _v: self._on_period_change(),
        )
        self.cmb_mes.set(f"{self.var_mes.get():02d} - {MES_NOMBRE_TITLE[self.var_mes.get()]}")
        self.cmb_mes.grid(row=0, column=3, padx=4)

        ctk.CTkLabel(top, text="Año").grid(row=0, column=4, padx=(12, 4))
        years = [str(y) for y in range(date.today().year - 3, date.today().year + 2)]
        self.cmb_anio = ctk.CTkComboBox(
            top,
            values=years,
            width=90,
            command=lambda _v: self._on_period_change(),
        )
        self.cmb_anio.set(str(self.var_anio.get()))
        self.cmb_anio.grid(row=0, column=5, padx=4)

        filt = ctk.CTkFrame(parent)
        filt.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        filt.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(filt, text="Buscar").grid(row=0, column=0, padx=4, pady=6)
        self.entry_query = ctk.CTkEntry(
            filt, textvariable=self.var_query, placeholder_text="actividad, cliente, descripción..."
        )
        self.entry_query.grid(row=0, column=1, sticky="ew", padx=4, pady=6)
        self.entry_query.bind("<KeyRelease>", lambda _e: self._apply_filters())

        ctk.CTkButton(filt, text="Aplicar filtros", width=120, command=self._apply_filters).grid(
            row=0, column=2, padx=4
        )
        ctk.CTkButton(filt, text="Limpiar búsqueda", width=130, command=self._clear_query).grid(
            row=0, column=3, padx=4
        )

        self.lbl_count = ctk.CTkLabel(filt, text="0 actividades")
        self.lbl_count.grid(row=0, column=4, padx=8)

        table_frame = ctk.CTkFrame(parent)
        table_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        cols = (
            "numero",
            "actividad",
            "cliente_proyecto",
            "fecha_desde",
            "fecha_hasta",
            "descripcion",
            "estatus",
        )
        self.tree = ttk.Treeview(
            table_frame,
            columns=cols,
            show="headings",
            selectmode="extended",
        )
        headings = {
            "numero": ("No.", 50),
            "actividad": ("Actividad", 180),
            "cliente_proyecto": ("Cliente / Proyecto", 140),
            "fecha_desde": ("Desde", 95),
            "fecha_hasta": ("Hasta", 95),
            "descripcion": ("Descripción", 360),
            "estatus": ("Estatus", 100),
        }
        for key, (label, width) in headings.items():
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width, anchor="w")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def _build_header_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        form = ctk.CTkFrame(parent)
        form.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        form.grid_columnconfigure(1, weight=1)

        # Profile selector + CRUD
        ctk.CTkLabel(form, text="Perfil", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, sticky="w", padx=8, pady=(10, 4)
        )
        self.cmb_profile = ctk.CTkComboBox(
            form,
            variable=self.var_profile,
            values=["(Sin perfil)"],
            command=self._on_profile_selected,
            width=280,
        )
        self.cmb_profile.grid(row=0, column=1, sticky="ew", padx=8, pady=(10, 4))

        profile_btns = ctk.CTkFrame(form, fg_color="transparent")
        profile_btns.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        ctk.CTkButton(profile_btns, text="Nuevo", width=70, command=self._profile_new).pack(
            side="left", padx=2
        )
        ctk.CTkButton(profile_btns, text="Guardar", width=80, command=self._profile_save).pack(
            side="left", padx=2
        )
        ctk.CTkButton(
            profile_btns, text="Guardar como…", width=110, command=self._profile_save_as
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            profile_btns, text="Eliminar", width=80, fg_color="#8B3A3A", hover_color="#6E2E2E",
            command=self._profile_delete,
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            profile_btns, text="Por defecto", width=100, command=self._profile_set_default
        ).pack(side="left", padx=2)

        fields = [
            ("Nombre del empleado", self.var_nombre),
            ("Cargo", self.var_cargo),
            ("Cargo (firma)", self.var_cargo_firma),
            ("Supervisor", self.var_supervisor),
            ("Cargo supervisor", self.var_cargo_supervisor),
            ("Empresa", self.var_empresa),
            ("Equipo (Excel)", self.var_equipo),
            ("Ciudad (Excel)", self.var_ciudad),
        ]
        for i, (label, var) in enumerate(fields):
            row = i + 2
            ctk.CTkLabel(form, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=6)
            ctk.CTkEntry(form, textvariable=var).grid(
                row=row, column=1, sticky="ew", padx=8, pady=6
            )

        code_row = len(fields) + 2
        ctk.CTkLabel(form, text="Código de informe").grid(
            row=code_row, column=0, sticky="w", padx=8, pady=6
        )
        ctk.CTkEntry(form, textvariable=self.var_codigo, state="readonly").grid(
            row=code_row, column=1, sticky="ew", padx=8, pady=6
        )

        right = ctk.CTkFrame(parent)
        right.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)

        ctk.CTkLabel(
            right,
            text="Generación de informes",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(12, 8))

        self.lbl_preview = ctk.CTkLabel(right, text="", justify="left", anchor="w")
        self.lbl_preview.pack(fill="x", padx=12, pady=8)

        ctk.CTkButton(
            right, text="Generar Word + Excel", height=40, command=self._generate_both
        ).pack(fill="x", padx=12, pady=6)
        ctk.CTkButton(right, text="Solo Word (.docx)", command=lambda: self._generate(word=True, excel=False)).pack(
            fill="x", padx=12, pady=4
        )
        ctk.CTkButton(right, text="Solo Excel (.xlsx)", command=lambda: self._generate(word=False, excel=True)).pack(
            fill="x", padx=12, pady=4
        )
        ctk.CTkButton(right, text="Abrir carpeta de exports", command=self._open_exports).pack(
            fill="x", padx=12, pady=(16, 4)
        )

        tip = (
            "Selecciona un perfil para rellenar la cabecera.\n"
            "Guardar mantiene el perfil actual; Guardar como crea uno nuevo.\n"
            "Se usan las actividades filtradas de la pestaña Actividades."
        )
        ctk.CTkLabel(right, text=tip, justify="left", text_color="gray").pack(
            anchor="w", padx=12, pady=16
        )

        self._update_preview()
        for var in (
            self.var_nombre,
            self.var_cargo,
            self.var_supervisor,
            self.var_equipo,
            self.var_ciudad,
        ):
            var.trace_add("write", lambda *_: self._update_preview())

    def _build_history_tab(self, parent: ctk.CTkFrame) -> None:
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)

        bar = ctk.CTkFrame(parent)
        bar.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        ctk.CTkButton(bar, text="Actualizar", command=self._reload_history, width=120).pack(
            side="left", padx=4, pady=4
        )
        ctk.CTkButton(bar, text="Abrir seleccionado", command=self._open_selected_report, width=150).pack(
            side="left", padx=4, pady=4
        )

        frame = ctk.CTkFrame(parent)
        frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)

        cols = ("id", "created_at", "codigo", "nombre", "count", "docx", "xlsx")
        self.hist = ttk.Treeview(frame, columns=cols, show="headings")
        labels = {
            "id": ("ID", 50),
            "created_at": ("Creado", 160),
            "codigo": ("Código", 200),
            "nombre": ("Empleado", 180),
            "count": ("Actividades", 90),
            "docx": ("Word", 260),
            "xlsx": ("Excel", 260),
        }
        for k, (lab, w) in labels.items():
            self.hist.heading(k, text=lab)
            self.hist.column(k, width=w, anchor="w")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.hist.yview)
        self.hist.configure(yscrollcommand=vsb.set)
        self.hist.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._reload_history()

    # ── helpers ───────────────────────────────────────────────────────
    def _profile_fields(self) -> dict[str, str]:
        return {
            "nombre": self.var_nombre.get().strip(),
            "cargo": self.var_cargo.get().strip(),
            "cargo_firma": self.var_cargo_firma.get().strip(),
            "supervisor": self.var_supervisor.get().strip(),
            "cargo_supervisor": self.var_cargo_supervisor.get().strip(),
            "empresa": self.var_empresa.get().strip(),
            "equipo": self.var_equipo.get().strip(),
            "ciudad": self.var_ciudad.get().strip(),
        }

    def _reload_profiles(self, select_default: bool = False) -> None:
        self._loading_profile = True
        try:
            self.profiles = storage.list_profiles()
            self._profile_label_to_id = {}
            values = ["(Sin perfil)"]
            for p in self.profiles:
                label = p["name"]
                if p.get("is_default"):
                    label = f"{p['name']} ⭐"
                self._profile_label_to_id[label] = p["id"]
                values.append(label)
            self.cmb_profile.configure(values=values)

            current = None
            if select_default:
                default = storage.get_default_profile() or (self.profiles[0] if self.profiles else None)
                current = default
            if current:
                self.current_profile_id = int(current["id"])
                label = current["name"]
                if current.get("is_default"):
                    label = f"{current['name']} ⭐"
                self.var_profile.set(label)
                self._load_profile_fields(current)
            elif select_default:
                self.var_profile.set("(Sin perfil)")
        finally:
            self._loading_profile = False

    def _load_profile_fields(self, profile: dict[str, Any]) -> None:
        self.var_nombre.set(profile.get("nombre") or "")
        self.var_cargo.set(profile.get("cargo") or "")
        self.var_cargo_firma.set(profile.get("cargo_firma") or "")
        self.var_supervisor.set(profile.get("supervisor") or "")
        self.var_cargo_supervisor.set(profile.get("cargo_supervisor") or "")
        self.var_empresa.set(profile.get("empresa") or "")
        self.var_equipo.set(profile.get("equipo") or "")
        self.var_ciudad.set(profile.get("ciudad") or "")

    def _profile_to_storage_dict(self) -> dict[str, str]:
        return self._profile_fields()

    def _on_profile_selected(self, choice: str | None = None) -> None:
        if self._loading_profile:
            return
        label = self.var_profile.get()
        if label == "(Sin perfil)":
            self.current_profile_id = None
            return
        pid = self._profile_label_to_id.get(label)
        if pid is None:
            return
        profile = storage.get_profile(pid)
        if profile:
            self.current_profile_id = int(profile["id"])
            self._load_profile_fields(profile)

    def _profile_new(self) -> None:
        name = ctk.CTkInputDialog(
            text="Nombre del nuevo perfil:",
            title="Nuevo perfil",
        ).get_input()
        if not name or not name.strip():
            return
        name = name.strip()
        try:
            self.current_profile_id = storage.create_profile(name=name, **self._profile_to_storage_dict())
            self._reload_profiles()
            # select it
            for label, pid in self._profile_label_to_id.items():
                if pid == self.current_profile_id:
                    self.var_profile.set(label)
                    break
            self._set_status(f"Perfil '{name}' creado.")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", str(exc))

    def _profile_save(self) -> None:
        if not self.current_profile_id:
            messagebox.showwarning("Guardar", "Selecciona o crea un perfil primero.")
            return
        name = self._resolve_selected_name()
        if not name:
            messagebox.showwarning("Guardar", "Selecciona un perfil.")
            return
        try:
            storage.update_profile(
                self.current_profile_id,
                name=name,
                **self._profile_to_storage_dict(),
            )
            self._reload_profiles(select_default=False)
            self._set_status(f"Perfil '{name}' actualizado.")
            self._select_profile_by_id(self.current_profile_id)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", str(exc))

    def _profile_save_as(self) -> None:
        name = ctk.CTkInputDialog(
            text="Nombre del nuevo perfil:",
            title="Guardar como",
        ).get_input()
        if not name or not name.strip():
            return
        name = name.strip()
        try:
            self.current_profile_id = storage.create_profile(
                name=name, **self._profile_to_storage_dict()
            )
            self._reload_profiles()
            self._select_profile_by_id(self.current_profile_id)
            self._set_status(f"Perfil '{name}' guardado.")
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error", str(exc))

    def _profile_delete(self) -> None:
        if not self.current_profile_id:
            messagebox.showwarning("Eliminar", "Selecciona un perfil.")
            return
        sel = self._resolve_selected_name()
        if not messagebox.askyesno(
            "Eliminar perfil",
            f"¿Eliminar el perfil '{sel}'? Esta acción no se puede deshacer.",
        ):
            return
        storage.delete_profile(self.current_profile_id)
        self.current_profile_id = None
        self._reload_profiles()
        self._set_status("Perfil eliminado.")

    def _profile_set_default(self) -> None:
        if not self.current_profile_id:
            messagebox.showwarning("Default", "Selecciona un perfil.")
            return
        storage.set_default_profile(self.current_profile_id)
        self._reload_profiles()
        self._select_profile_by_id(self.current_profile_id)
        self._set_status("Perfil marcado como predeterminado.")

    def _select_profile_by_id(self, pid: int) -> None:
        for label, lpid in self._profile_label_to_id.items():
            if lpid == pid:
                self.var_profile.set(label)
                break

    def _resolve_selected_name(self) -> str:
        label = self.var_profile.get()
        if label == "(Sin perfil)":
            return ""
        return label.replace(" ⭐", "").strip()

    def _period(self) -> tuple[int, int]:
        mes_txt = self.cmb_mes.get()
        try:
            mes = int(mes_txt.split("-")[0].strip())
        except ValueError:
            mes = self.var_mes.get()
        try:
            anio = int(self.cmb_anio.get())
        except ValueError:
            anio = self.var_anio.get()
        self.var_mes.set(mes)
        self.var_anio.set(anio)
        return mes, anio

    def _refresh_codigo(self) -> None:
        mes, anio = self._period()
        self.var_codigo.set(build_codigo(mes, anio))
        self._update_preview()

    def _on_period_change(self) -> None:
        self._refresh_codigo()
        self._apply_filters()

    def _clear_query(self) -> None:
        self.var_query.set("")
        self._apply_filters()

    def _update_preview(self) -> None:
        mes, anio = self._period()
        codigo = build_codigo(mes, anio)
        self.var_codigo.set(codigo)
        text = (
            f"Código: {codigo}\n"
            f"Periodo Word: {build_periodo_label(mes, anio)}\n"
            f"Periodo Excel: {build_periodo_corto(mes, anio)}\n"
            f"Empleado: {self.var_nombre.get() or '—'}\n"
            f"Cargo: {self.var_cargo.get() or '—'}\n"
            f"Supervisor: {self.var_supervisor.get() or '—'}\n"
            f"Actividades filtradas: {len(self.filtered_rows)}"
        )
        if hasattr(self, "lbl_preview"):
            self.lbl_preview.configure(text=text)

    def _set_status(self, msg: str) -> None:
        self.var_status.set(msg)

    # ── data ──────────────────────────────────────────────────────────
    def _import_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar CSV de actividades",
            filetypes=[("CSV", "*.csv"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            rows, warnings = load_csv(path)
            batch_id = storage.save_import_batch(Path(path).name, rows)
            self.current_batch_id = batch_id
            self.all_rows = storage.get_activities(batch_id)
            self._apply_filters()
            msg = f"Importadas {len(rows)} filas (batch #{batch_id})."
            if warnings:
                msg += f" Avisos: {len(warnings)}."
            self._set_status(msg)
            if warnings:
                messagebox.showwarning("Importación con avisos", "\n".join(warnings[:12]))
            else:
                messagebox.showinfo("Importación", msg)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error al importar", str(exc))
            self._set_status(f"Error: {exc}")

    def _load_latest_batch(self) -> None:
        batches = storage.list_batches()
        if not batches:
            self.all_rows = []
            self.filtered_rows = []
            self._refresh_table()
            self._set_status("No hay lotes importados todavía.")
            return
        self.current_batch_id = int(batches[0]["id"])
        self.all_rows = storage.get_activities(self.current_batch_id)
        self._apply_filters()
        self._set_status(
            f"Batch #{self.current_batch_id}: {batches[0]['filename']} "
            f"({batches[0]['row_count']} filas)."
        )

    def _apply_filters(self) -> None:
        mes, anio = self._period()
        self.filtered_rows = apply_filters(
            self.all_rows,
            mes=mes,
            anio=anio,
            query=self.var_query.get(),
        )
        self._refresh_table()
        self._update_preview()

    def _refresh_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.filtered_rows:
            self.tree.insert(
                "",
                "end",
                iid=str(row.get("id") or row.get("numero")),
                values=(
                    row.get("numero") or "",
                    row.get("actividad") or "",
                    row.get("cliente_proyecto") or "",
                    row.get("fecha_desde") or "",
                    row.get("fecha_hasta") or "",
                    row.get("descripcion") or "",
                    row.get("estatus") or "",
                ),
            )
        self.lbl_count.configure(
            text=f"{len(self.filtered_rows)} de {len(self.all_rows)} actividades"
        )

    def _header_dict(self) -> dict[str, Any]:
        mes, anio = self._period()
        today = date.today()
        cargo = self.var_cargo.get().strip()
        cargo_firma = self.var_cargo_firma.get().strip() or cargo
        return {
            "periodo_mes": mes,
            "periodo_anio": anio,
            "codigo_informe": build_codigo(mes, anio),
            "nombre": self.var_nombre.get().strip(),
            "cargo": cargo,
            "cargo_firma": cargo_firma,
            "supervisor": self.var_supervisor.get().strip(),
            "cargo_supervisor": self.var_cargo_supervisor.get().strip(),
            "empresa": self.var_empresa.get().strip(),
            "equipo": self.var_equipo.get().strip(),
            "ciudad": self.var_ciudad.get().strip(),
            "periodo": build_periodo_label(mes, anio),
            "periodo_corto": build_periodo_corto(mes, anio),
            "fecha_informe": format_fecha_informe(today.day, today.month, today.year),
        }

    def _generate_both(self) -> None:
        self._generate(word=True, excel=True)

    def _generate(self, *, word: bool, excel: bool) -> None:
        if not self.filtered_rows:
            messagebox.showwarning(
                "Sin actividades",
                "No hay actividades filtradas para el periodo seleccionado.",
            )
            return
        header = self._header_dict()
        if not header["nombre"]:
            messagebox.showwarning("Cabecera incompleta", "Indica el nombre del empleado.")
            return
        if not header["cargo"]:
            messagebox.showwarning("Cabecera incompleta", "Indica el cargo.")
            return

        codigo = header["codigo_informe"]
        exports = exports_dir()
        out_dir = exports / codigo
        # avoid overwrite: add timestamp subfolder if exists with files
        if out_dir.exists() and any(out_dir.iterdir()):
            from datetime import datetime

            out_dir = exports / f"{codigo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        out_dir.mkdir(parents=True, exist_ok=True)

        path_docx = None
        path_xlsx = None
        templates = templates_dir()
        try:
            if word:
                tpl = templates / "informe.docx"
                if not tpl.exists():
                    raise FileNotFoundError(f"Falta plantilla Word: {tpl}")
                path_docx = generate_word(
                    tpl,
                    out_dir / f"{codigo}.docx",
                    header=header,
                    activities=self.filtered_rows,
                )
            if excel:
                tpl = templates / "matriz.xlsx"
                if not tpl.exists():
                    raise FileNotFoundError(f"Falta plantilla Excel: {tpl}")
                path_xlsx = generate_excel(
                    tpl,
                    out_dir / f"Matriz-{codigo}.xlsx",
                    header=header,
                    activities=self.filtered_rows,
                )

            activity_ids = [int(r["id"]) for r in self.filtered_rows if r.get("id") is not None]
            storage.save_report(
                codigo_informe=codigo,
                periodo_mes=int(header["periodo_mes"]),
                periodo_anio=int(header["periodo_anio"]),
                nombre=header["nombre"],
                cargo=header["cargo"],
                supervisor=header["supervisor"],
                header=header,
                filters={
                    "mes": header["periodo_mes"],
                    "anio": header["periodo_anio"],
                    "query": self.var_query.get(),
                },
                activity_ids=activity_ids,
                path_docx=str(path_docx) if path_docx else None,
                path_xlsx=str(path_xlsx) if path_xlsx else None,
            )
            self._reload_history()
            # Persist the current header as a profile (auto-merge to a named one
            # if selected, otherwise save as the "Sin perfil" draft implicitly)
            try:
                if self.current_profile_id:
                    storage.update_profile(
                        self.current_profile_id,
                        **self._profile_to_storage_dict(),
                    )
            except Exception as exc:  # noqa: BLE001
                print("perfil no guardado:", exc)
            self._set_status(f"Informe generado en {out_dir}")
            messagebox.showinfo(
                "Listo",
                f"Informe generado:\n{out_dir}\n\n"
                + (f"Word: {path_docx.name}\n" if path_docx else "")
                + (f"Excel: {path_xlsx.name}" if path_xlsx else ""),
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Error al generar", str(exc))
            self._set_status(f"Error al generar: {exc}")

    def _open_exports(self) -> None:
        self._open_path(exports_dir())

    def _reload_history(self) -> None:
        for item in self.hist.get_children():
            self.hist.delete(item)
        for r in storage.list_reports():
            self.hist.insert(
                "",
                "end",
                values=(
                    r["id"],
                    (r["created_at"] or "")[:19].replace("T", " "),
                    r["codigo_informe"],
                    r["nombre"] or "",
                    r["activity_count"] or 0,
                    r["path_docx"] or "",
                    r["path_xlsx"] or "",
                ),
            )

    def _open_selected_report(self) -> None:
        sel = self.hist.selection()
        if not sel:
            messagebox.showinfo("Historial", "Selecciona un informe.")
            return
        vals = self.hist.item(sel[0], "values")
        docx = vals[5]
        xlsx = vals[6]
        path = docx or xlsx
        if not path:
            messagebox.showwarning("Historial", "Sin rutas guardadas.")
            return
        folder = Path(path).parent
        self._open_path(folder)

    @staticmethod
    def _open_path(path: Path) -> None:
        path = Path(path)
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)


def run() -> None:
    app = App()
    app.mainloop()
