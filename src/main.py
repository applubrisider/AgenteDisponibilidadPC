# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
import datetime as dt

# dotenv es opcional
try:
    from dotenv import load_dotenv
except Exception:
    def load_dotenv():  # no-op si no está instalado
        return None

from agent import logic
from agent.logic import (
    load_config, load_dataset, classify_all, save_outputs,
    build_html_report, build_html_report_kpi, compute_window_from_27,
    filtrar_sin_carga, save_sin_carga, build_html_sin_carga,
    normalize_rut_series
)

APP_NAME = "AgenteDisponibilidadPC"

def _parse_date(s: str | None):
    if not s:
        return None
    return dt.date.fromisoformat(s)

def get_base_dir(app_name: str = APP_NAME) -> Path:
    """
    Carpeta base escribible del usuario:
    %LOCALAPPDATA%\AgenteDisponibilidadPC (Windows)
    """
    return Path(os.getenv("LOCALAPPDATA", str(Path.home() / "AppData" / "Local"))) / app_name

def resolve_under_base(path_str: str, base: Path) -> Path:
    """
    Si la ruta es absoluta, se respeta.
    Si es relativa (p.ej. 'out' o 'models'), se coloca bajo la base.
    """
    p = Path(path_str)
    return p if p.is_absolute() else (base / p)

def ensure_dirs(*dirs: Path):
    for d in dirs:
        d.mkdir(parents=True, exist_ok=True)

def parse_args():
    p = argparse.ArgumentParser(description="Agente de Disponibilidad (PC)")
    # Origen de datos
    p.add_argument("--csv-url", type=str, help="URL del CSV en Google Drive (link de vista)")
    p.add_argument("--csv-file", type=str, help="Ruta local a CSV")
    p.add_argument("--ruts-file", type=str, help="(Opcional) CSV con RUTs válidos para filtrar")

    # Config & salida
    p.add_argument("--out", type=str, default="out", help="Carpeta de salida (relativa => LOCALAPPDATA/AgenteDisponibilidadPC/out)")
    p.add_argument("--config", type=str, default="config.yaml", help="Ruta a config.yaml")

    # Rango de fechas
    p.add_argument("--start", type=str, help="Inicio (YYYY-MM-DD)")
    p.add_argument("--end", type=str, help="Fin (YYYY-MM-DD)")

    # ML offline (opcional)
    p.add_argument("--use-ml", action="store_true", help="Usar clasificador ML offline")
    p.add_argument("--train-ml", action="store_true", help="Entrenar/actualizar el modelo ML")
    p.add_argument("--model-dir", type=str, default="models", help="Carpeta para el modelo ML (relativa => LOCALAPPDATA/AgenteDisponibilidadPC/models)")

    # Reporte “sin carga”
    p.add_argument("--null-csv", action="store_true", help="Exportar sin_carga.csv")
    p.add_argument("--null-html", action="store_true", help="Exportar sin_carga.html")

    # KPI extendido
    p.add_argument("--kpi-html", action="store_true", help="Generar informe KPI con gráficos")

    # Modos
    p.add_argument("--ui", action="store_true", help="Abrir interfaz gráfica")

    return p.parse_args()

def main():
    load_dotenv()
    args = parse_args()

    # UI
    if args.ui:
        from agent.ui import launch_ui
        launch_ui()
        return

    # Siempre que sean relativas, colocamos bajo %LOCALAPPDATA%\AgenteDisponibilidadPC
    base_dir  = get_base_dir()
    outdir    = resolve_under_base(args.out, base_dir)
    model_dir = resolve_under_base(args.model_dir, base_dir)

    # Crear carpetas escribibles
    try:
        ensure_dirs(base_dir, outdir, model_dir)
    except PermissionError as e:
        # Mensaje claro si por alguna razón el usuario no puede crear en LOCALAPPDATA
        raise SystemExit(f"No se pudo crear carpetas de trabajo en {base_dir} ({e}). "
                         "Ejecuta la app como usuario normal y verifica permisos.") from e

    # Cargar config (se respeta tal cual; puede estar junto al exe)
    cfg = load_config(args.config)

    # Cargar dataset
    df = load_dataset(csv_url=args.csv_url, csv_file=args.csv_file)
    if df is None or df.empty:
        raise SystemExit("No se pudo cargar el dataset. Verifica el origen (--csv-url o --csv-file).")

    # (Opcional) filtrar por RUTs válidos
    if args.ruts_file and os.path.exists(args.ruts_file):
        valid = logic.load_valid_ruts(args.ruts_file)
        if valid:
            df["RUT"] = normalize_rut_series(df["RUT"])
            df = df[df["RUT"].isin(valid)]

    # Rango de fechas
    start = _parse_date(args.start)
    end   = _parse_date(args.end)
    if not start or not end:
        start, end = compute_window_from_27()  # regla 27 → hoy

    # ML offline (opcional)
    ml = None
    if args.use_ml or args.train_ml:
        try:
            from agent.ml import ActivityClassifier
            model_path = model_dir / "activity_clf.joblib"

            if args.train_ml or not model_path.exists():
                ml = ActivityClassifier.train_from_dataframe(df, cfg)
                if ml is not None:
                    ml.save(model_path)
            if ml is None and model_path.exists():
                ml = ActivityClassifier.load(model_path)
        except Exception as e:
            print(f"[ML] Continuo sin ML por error: {e}")

    # Pipeline principal
    resumen, detalle, meta = classify_all(df, cfg, ml=ml, date_range=(start, end))

    # Guardar CSVs en outdir (siempre escribible)
    save_outputs(resumen, detalle, outdir)

    # Reporte principal
    html_path = build_html_report(resumen, meta, cfg, outdir)

    # KPI extendido
    if args.kpi_html:
        build_html_report_kpi(resumen, meta, outdir)

    # “Sin carga” (opcional)
    if args.null_csv or args.null_html:
        sc_df = filtrar_sin_carga(df, start=start, end=end)
        if args.null_csv:
            save_sin_carga(sc_df, outdir)
        if args.null_html:
            build_html_sin_carga(sc_df, {"rango": (start.isoformat(), end.isoformat())}, outdir)

    # Salida informativa
    print("OK. Directorios:")
    print(f"  BASE : {base_dir.resolve()}")
    print(f"  OUT  : {outdir.resolve()}")
    print(f"  MODEL: {model_dir.resolve()}")
    print("Archivos generados:")
    print(f"- {html_path}")
    print("- resumen_disponibilidad.csv")
    print("- detalle_disponibilidad.csv")
    if args.kpi_html:
        print("- informe_disponibilidad_kpi.html")
    if args.null_csv:
        print("- sin_carga.csv")
    if args.null_html:
        print("- sin_carga.html")

if __name__ == "__main__":
    main()
