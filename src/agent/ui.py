# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import json
import time
import zipfile
import threading
import webbrowser
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Dict, Any, Tuple

import pandas as pd
import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Calendario (opcional)
HAS_TKCALENDAR = False
try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except Exception:
    HAS_TKCALENDAR = False

# ====== Importar lógica ======
try:
    from .logic import (
        load_config,
        load_dataset,
        compute_window_from_27,
        classify_all,
        save_outputs,
        build_html_report,        # informe base
        build_html_report_kpi,    # (opcional) KPI simple
        build_email_report_kpi,   # ✅ KPI embebible para cuerpo de correo
        filtrar_sin_carga,
        save_sin_carga,
        build_html_sin_carga,
        normalize_nombre_series,
        format_rut_series,
        reorder_name_rut_first,
        load_valid_ruts,
        normalize_rut_series,
    )
except Exception:
    from agent.logic import (
        load_config,
        load_dataset,
        compute_window_from_27,
        classify_all,
        save_outputs,
        build_html_report,
        build_html_report_kpi,
        build_email_report_kpi,   # ✅ KPI embebible para cuerpo de correo
        filtrar_sin_carga,
        save_sin_carga,
        build_html_sin_carga,
        normalize_nombre_series,
        format_rut_series,
        reorder_name_rut_first,
        load_valid_ruts,
        normalize_rut_series,
    )

# =========================
# Constantes del proyecto
# =========================

ROOT = Path(__file__).resolve().parents[2]  # ...\AgenteDisponibilidadPC
LISTAS_DIR = ROOT / "Listas"
AUTH_DIR = ROOT / "auth"
OUT_DIR = ROOT / "out"
DATA_DIR = ROOT / "data"

GITHUB_PAT_FILE = AUTH_DIR / "github_pat.txt"
GH_OWNER = "applubrisider"
GH_REPO = "-actualizaciones-lubrisider"
GH_TAG = "v5.0.0"

ASSET_DISP = "disponibilidad_tidy.csv"
ASSET_RUTS = "Rut-Validos.csv"
ASSET_IDOP_ZIP = "id_operaciones.zip"
ASSET_IDOP_JSON = "id_operaciones.json"

# =========================
# Utilidades I/O & GitHub
# =========================

def log_safe(ui: "MainUI", msg: str) -> None:
    if ui:
        ui.append_log(msg)
    else:
        print(msg, flush=True)

def read_github_pat() -> Optional[str]:
    try:
        if GITHUB_PAT_FILE.exists():
            return GITHUB_PAT_FILE.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return None

def gh_headers(token: Optional[str]) -> Dict[str, str]:
    h = {"Accept": "application/vnd.github+json"}
    if token:
        h["Authorization"] = f"Bearer {token}"
    return h

def ensure_dirs() -> None:
    for d in (LISTAS_DIR, OUT_DIR, DATA_DIR, AUTH_DIR):
        d.mkdir(parents=True, exist_ok=True)

def gh_get_release_by_tag(tag: str, token: Optional[str]) -> Optional[Dict[str, Any]]:
    url = f"https://api.github.com/repos/{GH_OWNER}/{GH_REPO}/releases/tags/{tag}"
    r = requests.get(url, headers=gh_headers(token), timeout=30)
    if r.status_code == 200:
        return r.json()
    return None

def gh_download_asset(asset_name: str, dst: Path, token: Optional[str], ui: "MainUI" = None, *, force: bool = True) -> bool:
    ensure_dirs()
    if force and dst.exists():
        try:
            dst.unlink()
        except Exception:
            pass
    rel = gh_get_release_by_tag(GH_TAG, token)
    if not rel:
        log_safe(ui, f"⚠ No pude obtener el release {GH_TAG}.")
        return False
    assets = rel.get("assets", []) or []
    asset = next((a for a in assets if a.get("name") == asset_name), None)
    if not asset:
        log_safe(ui, f"⚠ No encontré el asset '{asset_name}' en {GH_TAG}.")
        return False
    url = asset.get("browser_download_url")
    if not url:
        log_safe(ui, f"⚠ Asset '{asset_name}' sin URL de descarga.")
        return False
    log_safe(ui, f"↓ Descargando {asset_name} …")
    r = requests.get(url, headers=gh_headers(token), timeout=180)
    if r.status_code != 200:
        log_safe(ui, f"✖ Error http {r.status_code} al descargar {asset_name}")
        return False
    dst.write_bytes(r.content)
    log_safe(ui, f"✓ Guardado {dst.name} ({dst.stat().st_size} bytes)")
    return True

def unzip_if_needed(zip_path: Path, dst_dir: Path, ui: "MainUI" = None) -> Optional[Path]:
    if not zip_path.exists():
        return None
    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(dst_dir)
        log_safe(ui, f"✓ Descomprimido en {dst_dir}")
    except Exception as e:
        log_safe(ui, f"✖ Error al descomprimir: {e}")
        return None
    json_path = dst_dir / ASSET_IDOP_JSON
    return json_path if json_path.exists() else None

# =========================
# Enriquecimiento de proyectos (opcional)
# =========================

def load_proyectos_mapping(ui: "MainUI" = None) -> Dict[str, Dict[str, str]]:
    ensure_dirs()
    json_local = LISTAS_DIR / ASSET_IDOP_JSON
    if json_local.exists():
        try:
            data = json.loads(json_local.read_text(encoding="utf-8"))
            return {row.get("Code", ""): {
                "desc": row.get("Desc", ""), "att1": row.get("Att1", ""), "att2": row.get("Att2", "")
            } for row in data if isinstance(row, dict)}
        except Exception as e:
            log_safe(ui, f"⚠ No pude leer {json_local.name}: {e}")
    zip_local = LISTAS_DIR / ASSET_IDOP_ZIP
    json_path = unzip_if_needed(zip_local, LISTAS_DIR, ui=ui)
    if json_path and json_path.exists():
        try:
            data = json.loads(json_path.read_text(encoding="utf-8"))
            return {row.get("Code", ""): {
                "desc": row.get("Desc", ""), "att1": row.get("Att1", ""), "att2": row.get("Att2", "")
            } for row in data if isinstance(row, dict)}
        except Exception as e:
            log_safe(ui, f"⚠ No pude leer {json_path.name}: {e}")
    log_safe(ui, "ℹ No hay lista de proyectos para enriquecer (opcional).")
    return {}

def enrich_with_proyectos(df: pd.DataFrame, mapping: Dict[str, Dict[str, str]], ui: "MainUI" = None) -> pd.DataFrame:
    if not mapping or "proyecto_actual" not in df.columns:
        return df
    def row_enrich(code: str) -> Dict[str, str]:
        return mapping.get(str(code).strip(), {"desc": "", "att1": "", "att2": ""})
    ext = df["proyecto_actual"].astype(str).map(row_enrich)
    ext_df = pd.json_normalize(ext)
    ext_df.columns = ["proyecto_desc", "proyecto_cliente", "proyecto_contacto"]
    out = pd.concat([df.reset_index(drop=True), ext_df.reset_index(drop=True)], axis=1)
    log_safe(ui, "✓ Enriquecimiento por proyectos aplicado.")
    return out

# =========================
# Pipeline de procesamiento
# =========================

@dataclass
class RunOptions:
    cfg_path: Optional[str]
    start_date: Optional[str]
    end_date: Optional[str]
    as_of: Optional[str]
    normalize_names: bool
    format_rut: bool
    reorder_name_rut: bool
    filter_valid_ruts: bool
    download_from_github: bool
    make_kpi_report: bool  # usar KPI embebible

def _df_with_calc_rut(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    if "RUT_fmt" in d.columns:
        if "RUT" in d.columns:
            d = d.drop(columns=["RUT"])
        d = d.rename(columns={"RUT_fmt": "RUT"})
    d = d.loc[:, ~d.columns.duplicated()]
    return d

def _parse_date_safe(s: Optional[str]) -> Optional[pd.Timestamp]:
    if not s:
        return None
    try:
        return pd.to_datetime(s).date()
    except Exception:
        return None

def run_pipeline(ui: "MainUI" | None, opts: RunOptions):
    ensure_dirs()
    token = read_github_pat()

    # 1) Descargas — SIEMPRE sobrescribe cuando está activado
    disp_path = DATA_DIR / ASSET_DISP
    ruts_path = LISTAS_DIR / ASSET_RUTS
    idop_zip  = LISTAS_DIR / ASSET_IDOP_ZIP

    if opts.download_from_github:
        gh_download_asset(ASSET_DISP, disp_path, token, ui=ui, force=True)
        gh_download_asset(ASSET_RUTS, ruts_path, token, ui=ui, force=True)
        gh_download_asset(ASSET_IDOP_ZIP, idop_zip,  token, ui=ui, force=True)
        unzip_if_needed(idop_zip, LISTAS_DIR, ui=ui)

    # 2) Config
    cfg_file = Path(opts.cfg_path) if opts.cfg_path else ROOT / "config.yml"
    cfg = load_config(str(cfg_file))
    log_safe(ui, f"✓ Config cargada: {cfg_file.name if cfg_file.exists() else '(por defecto)'}")

    # 3) Dataset disponibilidad
    if not disp_path.exists():
        raise FileNotFoundError(
            f"No se encontró {ASSET_DISP} en {DATA_DIR}. "
            f"Activa 'Descargar desde GitHub' o selecciona el archivo localmente."
        )
    df = load_dataset(csv_file=str(disp_path))
    if df is None or df.empty:
        raise RuntimeError("No pude cargar el dataset de disponibilidad.")
    log_safe(ui, f"✓ Disponibilidad: {len(df):,} filas")

    # 4) Normalizaciones
    if opts.normalize_names and "Colaborador" in df.columns:
        df["Colaborador"] = normalize_nombre_series(df["Colaborador"])
        log_safe(ui, "✓ Nombres normalizados")

    if "RUT" in df.columns:
        df["RUT"] = normalize_rut_series(df["RUT"])

    # 5) Filtrar por RUT válidos
    invalid_out = None
    if opts.filter_valid_ruts and ruts_path.exists():
        valid = load_valid_ruts(ruts_path)
        if valid:
            before = len(df)
            mask_valid = df["RUT"].isin(valid)
            invalid = df.loc[~mask_valid].copy()
            df = df.loc[mask_valid].copy()
            after = len(df)
            log_safe(ui, f"✓ Filtrado por RUT válidos: {after:,}/{before:,} filas")
            if len(invalid):
                invalid_out = OUT_DIR / "ruts_fuera_de_lista.csv"
                invalid[["Colaborador", "RUT", "Departamento", "fecha", "actividad"]].to_csv(
                    invalid_out, index=False, encoding="utf-8-sig"
                )
                log_safe(ui, f"→ Guardado: {invalid_out.name}")

    # 6) RUT legible para UI
    if opts.format_rut and "RUT" in df.columns:
        df["RUT_fmt"] = format_rut_series(df["RUT"])
    else:
        df["RUT_fmt"] = df["RUT"]

    # 7) Reordenar
    if opts.reorder_name_rut:
        df = reorder_name_rut_first(df, prefer_name=("Colaborador", "nombre", "Nombre"), rut_col="RUT_fmt")
        log_safe(ui, "✓ Reordenadas columnas: Nombre/RUT al inicio")

    # 8) Enriquecer proyectos (opcional)
    mapping = load_proyectos_mapping(ui=ui)
    if mapping:
        df = enrich_with_proyectos(df, mapping, ui=ui)

    # 9) Guardar dataset normalizado
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    norm_path = OUT_DIR / "disponibilidad_normalizada.csv"
    df.to_csv(norm_path, index=False, encoding="utf-8-sig")
    log_safe(ui, f"✓ Guardado: {norm_path.name}")

    # 10) Ventana de análisis
    s_override = _parse_date_safe(opts.start_date)
    e_override = _parse_date_safe(opts.end_date)
    if s_override and e_override:
        start, end = s_override, e_override
    else:
        start, end = compute_window_from_27(opts.as_of)
    log_safe(ui, f"✓ Ventana: {start} → {end}")

    # 11) Sin carga
    df_calc = _df_with_calc_rut(df)
    sin_df = filtrar_sin_carga(df_calc, date_range=(start, end))
    sin_csv_path = Path(save_sin_carga(sin_df, OUT_DIR))
    sin_html_path = Path(build_html_sin_carga(sin_df, {"rango": (str(start), str(end))}, OUT_DIR))
    log_safe(ui, f"✓ Sin carga CSV: {sin_csv_path.name}")
    log_safe(ui, f"✓ Sin carga HTML: {sin_html_path.name}")

    # 12) Clasificación general
    resumen, detalle, meta = classify_all(df_calc, cfg, ml=None, date_range=(start, end))
    resumen_path = OUT_DIR / "resumen_disponibilidad.csv"
    detalle_path = OUT_DIR / "detalle_disponibilidad.csv"
    save_outputs(resumen, detalle, OUT_DIR)
    html_path = Path(build_html_report(resumen, meta, cfg, OUT_DIR))
    log_safe(ui, "✓ Resumen/Detalle CSV generados")
    log_safe(ui, f"✓ Informe HTML: {html_path.name}")

    # 13) KPI extendido — HTML embebible (para cuerpo de correo)
    html_kpi_path = None
    if opts.make_kpi_report:
        html_kpi_path = Path(build_email_report_kpi(resumen, detalle, sin_df, meta, OUT_DIR))
        log_safe(ui, f"✓ Informe KPI (cuerpo HTML): {html_kpi_path.name}")

    # 14) Abrir carpeta salida
    try:
        os.startfile(str(OUT_DIR))
    except Exception:
        webbrowser.open(OUT_DIR.as_uri())

    outputs = {
        "norm": norm_path,
        "resumen": resumen_path,
        "detalle": detalle_path,
        "informe": html_path,
        "sin_carga_csv": sin_csv_path,
        "sin_carga_html": sin_html_path,
    }
    if html_kpi_path:
        outputs["informe_kpi"] = html_kpi_path
    if invalid_out:
        outputs["ruts_fuera"] = invalid_out

    return (str(start), str(end)), outputs

# =========================
# Interfaz Tkinter
# =========================

class MainUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Agente – Disponibilidad Operativa (PC)")
        self.geometry("1120x720")
        self.minsize(1000, 620)

        ensure_dirs()

        # Vars
        self.var_cfg = tk.StringVar(value=str((ROOT / "config.yml")))
        self.var_asof = tk.StringVar(value="")
        self.var_start = tk.StringVar()
        self.var_end = tk.StringVar()
        self.var_normalize_names = tk.BooleanVar(value=True)
        self.var_format_rut = tk.BooleanVar(value=True)
        self.var_reorder = tk.BooleanVar(value=True)
        self.var_filter_valid = tk.BooleanVar(value=True)
        self.var_download = tk.BooleanVar(value=True)
        self.var_make_kpi = tk.BooleanVar(value=True)   # ✅ generar KPI embebible
        self.var_to = tk.StringVar(value="")
        self.var_cc = tk.StringVar(value="")

        # refs a DateEntry (si existen)
        self.de_start = None
        self.de_end = None

        self.outputs: Dict[str, Path] = {}
        self._build_ui()
        self._binds()

        # Set default window 27→hoy (refresca widgets y vars)
        self._set_27_today()

    def _pick_locale(self) -> str:
        for loc in ("es_CL", "es_ES", "es"):
            return loc
        return "en_US"

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 8}

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        # === Config & opciones ===
        frm = ttk.LabelFrame(container, text="Opciones de ejecución")
        frm.pack(fill="x", **pad)

        r = 0
        ttk.Label(frm, text="Config YAML:").grid(row=r, column=0, sticky="w", padx=6, pady=6)
        ttk.Entry(frm, textvariable=self.var_cfg, width=66).grid(row=r, column=1, columnspan=3, sticky="we", padx=6, pady=6)
        ttk.Button(frm, text="Buscar…", command=self._pick_cfg).grid(row=r, column=4, padx=6, pady=6)

        # --- Fechas en columnas separadas (nada de place()) ---
        r += 1
        ttk.Label(frm, text="Desde:").grid(row=r, column=0, sticky="e", padx=6)
        if HAS_TKCALENDAR:
            self.de_start = DateEntry(frm, date_pattern="yyyy-mm-dd", locale=self._pick_locale(), width=12)
            self.de_start.grid(row=r, column=1, sticky="w", padx=6)
            ttk.Label(frm, text="Hasta:").grid(row=r, column=2, sticky="e", padx=6)
            self.de_end = DateEntry(frm, date_pattern="yyyy-mm-dd", locale=self._pick_locale(), width=12)
            self.de_end.grid(row=r, column=3, sticky="w", padx=6)
        else:
            ttk.Entry(frm, textvariable=self.var_start, width=14).grid(row=r, column=1, sticky="w", padx=6)
            ttk.Label(frm, text="Hasta:").grid(row=r, column=2, sticky="e", padx=6)
            ttk.Entry(frm, textvariable=self.var_end, width=14).grid(row=r, column=3, sticky="w", padx=6)

        ttk.Button(frm, text="Usar 27 → Hoy", command=self._set_27_today).grid(row=r, column=4, sticky="w", padx=6)

        r += 1
        ttk.Checkbutton(frm, text="Descargar desde GitHub (siempre sobrescribe)", variable=self.var_download).grid(row=r, column=0, sticky="w", padx=6)
        ttk.Checkbutton(frm, text="Normalizar nombres (Colaborador)", variable=self.var_normalize_names).grid(row=r, column=1, sticky="w", padx=6)
        ttk.Checkbutton(frm, text="Formatear RUT con puntos (columna RUT_fmt)", variable=self.var_format_rut).grid(row=r, column=2, sticky="w", padx=6)

        r += 1
        ttk.Checkbutton(frm, text="Nombre + RUT primeros", variable=self.var_reorder).grid(row=r, column=0, sticky="w", padx=6)
        ttk.Checkbutton(frm, text="Filtrar por RUTs válidos (Rut-Validos.csv)", variable=self.var_filter_valid).grid(row=r, column=1, sticky="w", padx=6)
        ttk.Checkbutton(frm, text="Generar Informe KPI + Gráficos", variable=self.var_make_kpi).grid(row=r, column=2, sticky="w", padx=6)

        # === Destinatarios ===
        r += 1
        ttk.Label(frm, text="Para (adicionales):").grid(row=r, column=0, sticky="e", padx=6)
        ttk.Entry(frm, textvariable=self.var_to).grid(row=r, column=1, sticky="we", padx=6)
        ttk.Label(frm, text="CC:").grid(row=r, column=2, sticky="e", padx=6)
        ttk.Entry(frm, textvariable=self.var_cc).grid(row=r, column=3, sticky="we", padx=6)

        for c in range(5):
            frm.grid_columnconfigure(c, weight=1)

        # === Acciones ===
        actions = ttk.Frame(container)
        actions.pack(fill="x", **pad)

        ttk.Button(actions, text="Descargar ahora", command=self.on_download).pack(side="left", padx=4)
        ttk.Button(actions, text="Procesar", command=self.on_run).pack(side="left", padx=4)
        ttk.Button(actions, text="Abrir carpeta de salida", command=self.open_out).pack(side="left", padx=4)
        ttk.Button(actions, text="Abrir carpeta Listas", command=self.open_listas).pack(side="left", padx=4)
        ttk.Button(actions, text="Enviar correo", command=self.on_email).pack(side="right", padx=4)
        ttk.Button(actions, text="Salir", command=self.destroy).pack(side="right", padx=4)

        # === Archivos generados ===
        files_frame = ttk.LabelFrame(container, text="Archivos generados (abrir)")
        files_frame.pack(fill="x", **pad)
        self.btn_open_informe = ttk.Button(files_frame, text="Informe HTML", command=lambda: self._open_file_key("informe"))
        self.btn_open_informe_k = ttk.Button(files_frame, text="Informe KPI (HTML)", command=lambda: self._open_file_key("informe_kpi"))
        self.btn_open_sin_html = ttk.Button(files_frame, text="Sin Carga (HTML)", command=lambda: self._open_file_key("sin_carga_html"))
        self.btn_open_sin_csv = ttk.Button(files_frame, text="Sin Carga (CSV)", command=lambda: self._open_file_key("sin_carga_csv"))
        self.btn_open_resumen = ttk.Button(files_frame, text="Resumen CSV", command=lambda: self._open_file_key("resumen"))
        self.btn_open_detalle = ttk.Button(files_frame, text="Detalle CSV", command=lambda: self._open_file_key("detalle"))
        self.btn_open_norm = ttk.Button(files_frame, text="Normalizado CSV", command=lambda: self._open_file_key("norm"))

        for i, b in enumerate([self.btn_open_informe, self.btn_open_informe_k, self.btn_open_sin_html,
                               self.btn_open_sin_csv, self.btn_open_resumen, self.btn_open_detalle, self.btn_open_norm]):
            b.grid(row=0, column=i, padx=6, pady=6, sticky="w")
        self._set_file_buttons_state(False)

        # === Log ===
        log_frame = ttk.LabelFrame(container, text="Registro / Log")
        log_frame.pack(fill="both", expand=True, **pad)

        self.txt = tk.Text(log_frame, height=18, wrap="word")
        self.txt.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt.yview)
        yscroll.pack(side="right", fill="y")
        self.txt.configure(yscrollcommand=yscroll.set)

        self.append_log("Listo. Presiona 'Procesar' para generar los informes.")

    def _binds(self) -> None:
        self.bind("<Control-Return>", lambda e: self.on_run())

    # ==== helpers UI ====

    def _pick_cfg(self) -> None:
        p = filedialog.askopenfilename(
            title="Seleccionar config.yml",
            filetypes=[("YAML", "*.yml *.yaml"), ("Todos", "*.*")]
        )
        if p:
            self.var_cfg.set(p)

    def open_out(self) -> None:
        try:
            os.startfile(str(OUT_DIR))
        except Exception:
            webbrowser.open(OUT_DIR.as_uri())

    def open_listas(self) -> None:
        try:
            os.startfile(str(LISTAS_DIR))
        except Exception:
            webbrowser.open(LISTAS_DIR.as_uri())

    def append_log(self, msg: str) -> None:
        ts = time.strftime("%H:%M:%S")
        self.txt.insert("end", f"[{ts}] {msg}\n")
        self.txt.see("end")
        self.update_idletasks()

    def _set_27_today(self):
        """Pone por defecto la ventana [último 27 → hoy] tanto en vars como en widgets."""
        s, e = compute_window_from_27(None)
        self.var_start.set(str(s))
        self.var_end.set(str(e))
        if HAS_TKCALENDAR:
            try:
                if self.de_start:
                    self.de_start.set_date(s)
                if self.de_end:
                    self.de_end.set_date(e)
            except Exception:
                pass

    def _sync_dates_from_widgets(self):
        if HAS_TKCALENDAR:
            try:
                if self.de_start:
                    self.var_start.set(self.de_start.get_date().strftime("%Y-%m-%d"))
                if self.de_end:
                    self.var_end.set(self.de_end.get_date().strftime("%Y-%m-%d"))
            except Exception:
                pass

    def _set_file_buttons_state(self, enabled: bool):
        for b in (self.btn_open_informe, self.btn_open_informe_k, self.btn_open_sin_html,
                  self.btn_open_sin_csv, self.btn_open_resumen, self.btn_open_detalle, self.btn_open_norm):
            b.configure(state=("normal" if enabled else "disabled"))

    def _open_file_key(self, key: str):
        p = self.outputs.get(key)
        if not p or not Path(p).exists():
            messagebox.showwarning("Abrir archivo", "El archivo aún no existe.")
            return
        try:
            os.startfile(str(p))
        except Exception:
            webbrowser.open(Path(p).as_uri())

    # ==== acciones ====

    def on_download(self) -> None:
        def _task():
            try:
                token = read_github_pat()
                ensure_dirs()
                gh_download_asset(ASSET_DISP, DATA_DIR / ASSET_DISP, token, ui=self, force=True)
                gh_download_asset(ASSET_RUTS, LISTAS_DIR / ASSET_RUTS, token, ui=self, force=True)
                ok3 = gh_download_asset(ASSET_IDOP_ZIP, LISTAS_DIR / ASSET_IDOP_ZIP, token, ui=self, force=True)
                if ok3:
                    unzip_if_needed(LISTAS_DIR / ASSET_IDOP_ZIP, LISTAS_DIR, ui=self)
                messagebox.showinfo("Descargas", "Completado.")
            except Exception as e:
                messagebox.showerror("Descargas", f"Error: {e}")

        threading.Thread(target=_task, daemon=True).start()

    def on_run(self) -> None:
        self._sync_dates_from_widgets()
        opts = RunOptions(
            cfg_path=self.var_cfg.get().strip() or None,
            start_date=self.var_start.get().strip() or None,
            end_date=self.var_end.get().strip() or None,
            as_of=self.var_asof.get().strip() or None,
            normalize_names=self.var_normalize_names.get(),
            format_rut=self.var_format_rut.get(),
            reorder_name_rut=self.var_reorder.get(),
            filter_valid_ruts=self.var_filter_valid.get(),
            download_from_github=self.var_download.get(),
            make_kpi_report=self.var_make_kpi.get(),
        )

        def _task():
            try:
                (start, end), outputs = run_pipeline(self, opts)
                self.outputs = outputs
                self._set_file_buttons_state(True)
                self.append_log(f"✓ Rango final usado: {start} – {end}")
            except Exception as e:
                self.append_log(f"✖ Error: {e}")
                messagebox.showerror("Error", str(e))

        threading.Thread(target=_task, daemon=True).start()

    def on_email(self):
        PRIMARY_TO = "areyes@lubrisider.cl"
        if not self.outputs:
            messagebox.showwarning("Enviar correo", "Primero genera los informes (Procesar).")
            return

        start = self.var_start.get() or ""
        end = self.var_end.get() or ""
        subject = f"Disponibilidad Operativa — {start} a {end}"

        kpi_path = self.outputs.get("informe_kpi")
        if not kpi_path or not Path(kpi_path).exists():
            messagebox.showwarning("Enviar correo", "No se encontró el KPI HTML. Vuelve a procesar con 'Generar Informe KPI + Gráficos'.")
            return

        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.Subject = subject

            html_file = self.outputs.get("informe_kpi") or self.outputs.get("informe")
            if html_file and Path(html_file).exists():
                mail.HTMLBody = Path(html_file).read_text(encoding="utf-8")
            else:
                mail.Body = "Adjunto informe."

            extras = [e.strip() for e in self.var_to.get().split(";") if e.strip()]
            to_list = [PRIMARY_TO] + extras
            mail.To = "; ".join(dict.fromkeys(to_list))
            if self.var_cc.get().strip():
                mail.CC = self.var_cc.get().strip()

            mail.Display()
            self.append_log("✓ Borrador de correo abierto en Outlook.")
            return
        except Exception as e:
            self.append_log(f"ℹ No se pudo usar Outlook ({e}). Abriendo el HTML para copiar/pegar manualmente…")
            try:
                os.startfile(str(kpi_path))
            except Exception:
                webbrowser.open(Path(kpi_path).as_uri())
            messagebox.showinfo(
                "Enviar correo",
                "Outlook no está disponible. Se abrió el HTML; copia/pega su contenido como cuerpo del correo."
            )

# =========================
# CLI fallback (sin UI)
# =========================

def launch_ui() -> None:
    app = MainUI()
    app.mainloop()
