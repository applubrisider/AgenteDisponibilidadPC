# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import re
import math
import base64
import datetime as dt
from pathlib import Path
from typing import Tuple, Dict, Any, Optional

import pandas as pd
import requests
import yaml

# matplotlib para gráficos (modo no interactivo)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# =========================
# Utilidades de fechas
# =========================

def compute_window_from_27(as_of: Optional[str] = None) -> tuple[dt.date, dt.date]:
    """
    Ventana: comienza el día 27 más cercano hacia atrás y termina en 'as_of' (o hoy).
    - Si 'as_of'.day >= 27 -> inicio = (as_of.year, as_of.month, 27)
    - Si 'as_of'.day < 27   -> inicio = 27 del mes anterior
    """
    today = dt.date.fromisoformat(as_of) if as_of else dt.date.today()
    if today.day >= 27:
        start = today.replace(day=27)
    else:
        first = today.replace(day=1)
        prev_month_last = first - dt.timedelta(days=1)
        start = prev_month_last.replace(day=27)
    return start, today

# =========================
# Config
# =========================

_DEFAULT_CFG = {
    "availability": {
        "whitelist_keywords": ["disponible"],
        "blacklist_prefixes": ["ser-", "con-", "lab-"],
        "blacklist_exact": [
            "descanso", "vacaciones", "licencia médica", "licencia medica",
            "descanso en zona"
        ],
        "neutral_keywords": [
            "oficina", "actividad interna", "teletrabajo", "capacitacion",
            "capacitación", "academia", "capacitaciones presenciales",
            "oficina central", "oficina central sucre"
        ],
    },
    "rules": {
        "high_days": 7,
        "high_streak": 3,
        "low_days": 4
    },
    # Filtros operativos (puedes cambiarlos en config.yml)
    "filters": {
        "allowed_departments": ["OPER", "OPER_OF"]
    }
}

def load_config(path: str | Path) -> Dict[str, Any]:
    p = Path(path)
    if not p.exists():
        return _DEFAULT_CFG.copy()
    with open(p, "r", encoding="utf-8") as f:
        try:
            cfg = yaml.safe_load(f) or {}
        except Exception:
            cfg = {}
    out = _DEFAULT_CFG.copy()
    out.update(cfg or {})
    if "availability" in cfg:
        out["availability"].update(cfg["availability"] or {})
    if "rules" in cfg:
        out["rules"].update(cfg["rules"] or {})
    if "filters" in cfg:
        out["filters"].update(cfg["filters"] or {})
    return out

# =========================
# RUT helpers
# =========================

_RUT_CLEAN_RE = re.compile(r"[^0-9kK]")
_RUT_VALUE_RE = re.compile(r"^\s*\d{6,9}-[0-9Kk]\s*$")  # 6-9 dígitos + DV

def _strip_bom(s: str) -> str:
    # quita BOM real (\ufeff) y la secuencia mal-decodificada "ï»¿"
    return s.replace("\ufeff", "").replace("ï»¿", "")

def normalize_rut_value(x: Any) -> str:
    """
    Normaliza a formato simple: 12345678-9 (sin puntos; DV mayúscula).
    Si no se puede, devuelve "".
    """
    if x is None:
        return ""
    s = str(x).strip()
    if not s:
        return ""
    s = _RUT_CLEAN_RE.sub("", s)  # solo dígitos y k/K
    if len(s) < 2:
        return ""
    s = f"{s[:-1]}-{s[-1].upper()}"
    return s if _RUT_VALUE_RE.match(s) else ""

def normalize_rut_series(s: pd.Series) -> pd.Series:
    return s.astype(str).map(normalize_rut_value)

def _read_any_header_csv(path: str | Path) -> pd.DataFrame | None:
    # prueba codificaciones y con/sin encabezado
    for enc in ("utf-8-sig", "utf-8", "latin1"):
        for header in (0, None):
            try:
                df = pd.read_csv(path, sep=None, engine="python", dtype=str, encoding=enc, header=header)
                if df is not None:
                    df.columns = [_strip_bom(str(c)) for c in df.columns]
                    return df
            except Exception:
                pass
    return None

def _score_rut_column(series: pd.Series) -> float:
    """% de filas que parecen RUT en esta serie (heurística por valores)."""
    vals = series.dropna().astype(str)
    if len(vals) == 0:
        return 0.0
    def norm_probe(s: str) -> str:
        s2 = _RUT_CLEAN_RE.sub("", s)
        if len(s2) >= 2:
            s2 = f"{s2[:-1]}-{s2[-1]}"
        return s2
    matches = vals.map(lambda s: bool(_RUT_VALUE_RE.match(norm_probe(s))))
    return float(matches.mean())

def load_valid_ruts(path: str | Path) -> set[str]:
    """
    Lee CSV con o sin encabezado; detecta la columna RUT por nombre
    ('rut', 'RUT', 'ï»¿rut', etc.) o por patrón de valores y normaliza.
    """
    p = Path(path)
    if not p.exists():
        return set()

    df = _read_any_header_csv(p)
    if df is None or df.empty:
        return set()

    # 1) por nombre
    def norm_name(c: str) -> str:
        x = _strip_bom(c).lower().strip()
        x = (x.replace("á","a").replace("é","e").replace("í","i")
               .replace("ó","o").replace("ú","u").replace("ñ","n"))
        return re.sub(r"\s+", " ", x)

    name_map = {c: norm_name(str(c)) for c in df.columns}
    by_name = [c for c, nc in name_map.items() if nc in ("rut","ruts","rut valido","ruts validos")]
    if by_name:
        col = by_name[0]
    else:
        # 2) por patrón de valores
        best_col, best_score = None, 0.0
        for c in df.columns:
            score = _score_rut_column(df[c])
            if score > best_score:
                best_col, best_score = c, score
        col = best_col if best_col is not None and best_score >= 0.2 else df.columns[0]

    vals = df[col].dropna().astype(str)
    out = set()
    for s in vals:
        n = normalize_rut_value(s)
        if n:
            out.add(n)
    return out

# =========================
# Carga robusta del dataset
# =========================

_EXPECTED_COLS = ["Colaborador", "Departamento", "RUT", "Ciudad Residencia", "fecha", "actividad"]

def _read_csv_robust(source, is_bytes: bool = False) -> Optional[pd.DataFrame]:
    seps = [",", ";", "\t", "|"]
    encs = ["utf-8-sig", "utf-8", "latin1"]
    for enc in encs:
        for sep in seps:
            try:
                if is_bytes:
                    df = pd.read_csv(io.BytesIO(source), sep=sep, encoding=enc, dtype=str, engine="python")
                else:
                    df = pd.read_csv(source, sep=sep, encoding=enc, dtype=str, engine="python")
                if df is not None and len(df.columns) >= 5:
                    return df
            except Exception:
                continue
    return None

def _collapse_duplicate_rut_columns(d: pd.DataFrame) -> pd.DataFrame:
    """Si hay columnas duplicadas llamadas 'RUT', consolida en una sola."""
    rut_idxs = [i for i, c in enumerate(d.columns) if c == "RUT"]
    if len(rut_idxs) <= 1:
        return d
    # usar la primera como base y completar con las siguientes (no nulas / no vacías)
    base = d.iloc[:, rut_idxs[0]].copy()
    for idx in rut_idxs[1:]:
        aux = d.iloc[:, idx]
        mask = base.isna() | (base.astype(str).str.strip() == "")
        base = base.where(~mask, aux)
    d = d.drop(columns=[d.columns[i] for i in rut_idxs[1:]])
    d["RUT"] = normalize_rut_series(base)
    d = d.loc[:, ~d.columns.duplicated()]
    return d

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    def norm(s: str) -> str:
        s = s.strip().lower()
        s = re.sub(r"\s+", " ", s)
        s = (s.replace("á", "a").replace("é", "e").replace("í", "i")
               .replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
        return s

    colmap = {}
    for c in df.columns:
        nc = norm(c)
        if nc == "colaborador":
            colmap[c] = "Colaborador"
        elif nc == "departamento":
            colmap[c] = "Departamento"
        elif nc == "rut":
            colmap[c] = "RUT"
        elif "ciudad" in nc and "residenc" in nc:
            colmap[c] = "Ciudad Residencia"
        elif nc == "fecha":
            colmap[c] = "fecha"
        elif nc == "actividad":
            colmap[c] = "actividad"

    d = df.rename(columns=colmap).copy()

    if (pd.Series(list(d.columns)) == "RUT").sum() > 1:
        d = _collapse_duplicate_rut_columns(d)

    missing = [c for c in _EXPECTED_COLS if c not in d.columns]
    if missing:
        raise ValueError(f"Faltan columnas requeridas: {missing}. Detectadas: {list(d.columns)}")

    for c in ["Colaborador", "Departamento", "RUT", "Ciudad Residencia", "actividad"]:
        d[c] = d[c].astype(str).str.strip()

    d["RUT"] = normalize_rut_series(d["RUT"])
    d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce").dt.date
    d = d.dropna(subset=["fecha"])
    return d

def _drive_view_to_download(url: str) -> str:
    m = re.search(r"/d/([A-Za-z0-9_\-]+)/", url)
    if not m:
        return url
    fid = m.group(1)
    return f"https://drive.google.com/uc?export=download&id={fid}"

def load_dataset(csv_url: Optional[str] = None, csv_file: Optional[str] = None) -> Optional[pd.DataFrame]:
    df = None
    if csv_file and Path(csv_file).exists():
        df_raw = _read_csv_robust(csv_file, is_bytes=False)
        if df_raw is not None:
            df = _normalize_columns(df_raw)
            return df

    if csv_url:
        url = _drive_view_to_download(csv_url.strip())
        try:
            r = requests.get(url, timeout=30)
            r.raise_for_status()
            df_raw = _read_csv_robust(r.content, is_bytes=True)
            if df_raw is not None:
                df = _normalize_columns(df_raw)
        except Exception:
            df = None
    return df

# =========================
# Clasificación de actividad
# =========================

def classify_activity_flag(text: str, cfg: Dict[str, Any]) -> int:
    """
    1 si 'Disponible', 0 en otro caso (SER-/CON-/LAB-, blacklist, neutrales...).
    """
    if not isinstance(text, str) or not text:
        return 0
    t = text.strip().lower()

    wl = [w.lower() for w in cfg["availability"].get("whitelist_keywords", [])]
    if any(w in t for w in wl):
        return 1

    blp = [p.lower() for p in cfg["availability"].get("blacklist_prefixes", [])]
    if any(t.startswith(p) for p in blp):
        return 0

    ble = [w.lower() for w in cfg["availability"].get("blacklist_exact", [])]
    if any(w == t for w in ble):
        return 0

    ne = [w.lower() for w in cfg["availability"].get("neutral_keywords", [])]
    if any(w in t for w in ne):
        return 0

    return 0

def _max_consecutive_days(dates_sorted: list[dt.date]) -> int:
    if not dates_sorted:
        return 0
    best = 1
    cur = 1
    for i in range(1, len(dates_sorted)):
        if dates_sorted[i] == dates_sorted[i-1] + dt.timedelta(days=1):
            cur += 1
            best = max(best, cur)
        else:
            cur = 1
    return best

# =========================
# Sin carga helpers
# =========================

def is_nullish(x: Any) -> bool:
    if x is None:
        return True
    try:
        if pd.isna(x):
            return True
    except Exception:
        pass
    s = str(x).strip().lower()
    return s in ("", "null", "none", "nan")

def build_sin_carga(df: pd.DataFrame, date_range: Optional[Tuple[dt.date, dt.date]]) -> pd.DataFrame:
    d = df.copy()
    if "actividad" not in d.columns:
        raise ValueError("El dataset no tiene la columna 'actividad'.")

    if date_range is not None:
        a, b = date_range
        d = d[(d["fecha"] >= a) & (d["fecha"] <= b)]

    mask = d["actividad"].map(is_nullish)
    cols = [c for c in ["Colaborador","RUT","Departamento","Ciudad Residencia","fecha","actividad"] if c in d.columns]
    out = d.loc[mask, cols].sort_values(["fecha","Colaborador"] if "fecha" in cols else cols[:1])
    return out.reset_index(drop=True)

def filtrar_sin_carga(
    df: pd.DataFrame,
    date_range: Optional[Tuple[dt.date, dt.date]] = None,
    start: Optional[dt.date] = None,
    end: Optional[dt.date] = None,
) -> pd.DataFrame:
    """Compatibilidad: aceptar date_range=(ini,fin) o start=..., end=..."""
    if date_range is None and (start is not None and end is not None):
        date_range = (start, end)
    return build_sin_carga(df, date_range)

def save_sin_carga(sin_df: pd.DataFrame, outdir: Path) -> str:
    outdir.mkdir(parents=True, exist_ok=True)
    out_path = outdir / "sin_carga.csv"
    cols = [c for c in ["Colaborador","RUT","Departamento","Ciudad Residencia","fecha","actividad"] if c in sin_df.columns]
    (sin_df[cols] if cols else sin_df).to_csv(out_path, index=False, encoding="utf-8-sig")
    return str(out_path.resolve())

def build_html_sin_carga(sin_df: pd.DataFrame, meta: Dict[str, Any] | None, outdir: Path) -> str:
    outdir.mkdir(parents=True, exist_ok=True)

    if meta and "rango" in meta:
        ini, fin = meta["rango"]
    else:
        if "fecha" in sin_df.columns and not sin_df.empty:
            fechas = pd.to_datetime(sin_df["fecha"], errors="coerce").dt.date.dropna()
            ini = fechas.min().isoformat() if not fechas.empty else ""
            fin = fechas.max().isoformat() if not fechas.empty else ""
        else:
            ini, fin = "", ""

    total = len(sin_df)
    cols = [c for c in ["Colaborador","RUT","Departamento","Ciudad Residencia","fecha","actividad"] if c in sin_df.columns]
    df_view = sin_df[cols].copy() if cols else sin_df.copy()
    if "fecha" in df_view.columns:
        df_view = df_view.sort_values(["fecha","Colaborador"] if "Colaborador" in df_view.columns else ["fecha"])

    tabla = df_view.to_html(index=False, justify="center", classes="tabla", border=0)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<title>Informe — Sin Carga</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; }}
  h1 {{ margin-bottom: 4px; }}
  .meta {{ color:#444; margin-bottom: 16px; }}
  .kpis {{ display:flex; gap:16px; margin: 16px 0; }}
  .kpi {{ padding:12px 16px; border-radius:10px; background:#f4f6f8; }}
  .tabla {{ border-collapse: collapse; width: 100%; font-size: 14px; }}
  .tabla th, .tabla td {{ border-bottom: 1px solid #e5e5e5; padding: 8px 10px; text-align: left; }}
  .tabla tr:nth-child(even) {{ background: #fafafa; }}
</style>
</head>
<body>
  <h1>Informe — Personas sin carga</h1>
  <div class="meta">
    <div><b>Rango analizado:</b> {ini} a {fin}</div>
  </div>

  <div class="kpis">
    <div class="kpi"><b>Total sin carga</b><br><span style="font-weight:bold;">{total}</span></div>
  </div>

  {tabla}
</body>
</html>"""

    out_path = outdir / "sin_carga.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return str(out_path.resolve())

# =========================
# Filtros de negocio
# =========================

def _apply_filters(df: pd.DataFrame, cfg: Dict[str, Any]) -> pd.DataFrame:
    """Aplica filtros de negocio (departamentos permitidos, etc.)."""
    out = df.copy()
    allowed = [str(x).upper() for x in cfg.get("filters", {}).get("allowed_departments", [])]
    if allowed and "Departamento" in out.columns:
        out["Departamento_UP"] = out["Departamento"].astype(str).str.upper().str.strip()
        out = out[out["Departamento_UP"].isin(allowed)].drop(columns=["Departamento_UP"])
    return out

# =========================
# Reporte base + KPI/Charts
# =========================

def save_outputs(resumen: pd.DataFrame, detalle: pd.DataFrame, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    resumen.to_csv(outdir / "resumen_disponibilidad.csv", index=False, encoding="utf-8-sig")
    detalle.to_csv(outdir / "detalle_disponibilidad.csv", index=False, encoding="utf-8-sig")

def build_html_report(resumen: pd.DataFrame, meta: Dict[str, Any], cfg: Dict[str, Any], outdir: Path) -> str:
    outdir.mkdir(parents=True, exist_ok=True)
    ini, fin = meta.get("rango", ("", ""))
    um = meta.get("umbrales", {})
    window_days = um.get("window_days", 0)

    counts = resumen["criticidad"].value_counts().to_dict() if "criticidad" in resumen.columns else {}
    c_alta = counts.get("ALTA", 0); c_media = counts.get("MEDIA", 0); c_baja = counts.get("BAJA", 0)

    tbl = resumen.to_html(index=False, justify="center", classes="tabla", border=0)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<title>Informe de Disponibilidad</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; }}
  h1 {{ margin-bottom: 4px; }}
  .meta {{ color:#444; margin-bottom: 16px; }}
  .kpis {{ display:flex; gap:16px; margin: 16px 0; }}
  .kpi {{ padding:12px 16px; border-radius:10px; background:#f4f6f8; }}
  .tabla {{ border-collapse: collapse; width: 100%; font-size: 14px; }}
  .tabla th, .tabla td {{ border-bottom: 1px solid #e5e5e5; padding: 8px 10px; text-align: left; }}
  .tabla tr:nth-child(even) {{ background: #fafafa; }}
  .crit-ALTA {{ color: #0c7; font-weight:bold; }}
  .crit-MEDIA {{ color: #f90; font-weight:bold; }}
  .crit-BAJA {{ color: #e33; font-weight:bold; }}
</style>
</head>
<body>
  <h1>Informe de Disponibilidad</h1>
  <div class="meta">
    <div><b>Rango analizado:</b> {ini} a {fin} ({window_days} días)</div>
    <div><b>Umbrales:</b> ALTA ≥{um.get('th_high_days')} días o racha ≥{um.get('th_streak')} &nbsp;|&nbsp; BAJA ≤{um.get('th_low_days')} días</div>
  </div>

  <div class="kpis">
    <div class="kpi"><b>ALTA</b><br><span class="crit-ALTA">{c_alta}</span></div>
    <div class="kpi"><b>MEDIA</b><br><span class="crit-MEDIA">{c_media}</span></div>
    <div class="kpi"><b>BAJA</b><br><span class="crit-BAJA">{c_baja}</span></div>
  </div>

  {tbl}
</body>
</html>"""

    out_path = outdir / "informe_disponibilidad.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return str(out_path)

def _fig_to_base64(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("ascii")

def build_html_report_kpi(resumen: pd.DataFrame, meta: Dict[str, Any], outdir: Path) -> str:
    """Informe KPI simple con 2 gráficos (para referencia)."""
    outdir.mkdir(parents=True, exist_ok=True)
    ini, fin = meta.get("rango", ("", ""))

    # --- Chart 1: pie criticidad ---
    counts = resumen["criticidad"].value_counts().reindex(["ALTA","MEDIA","BAJA"]).fillna(0)
    fig1, ax1 = plt.subplots(figsize=(5,5))
    ax1.pie(counts.values,
            labels=counts.index,
            autopct="%1.0f%%",
            startangle=90,
            colors=["#12b886","#f59f00","#fa5252"])
    ax1.set_title("Distribución por criticidad")
    img1 = _fig_to_base64(fig1)

    # --- Chart 2: top-15 por días disponibles ---
    top = (resumen.sort_values(["dias_disponibles","streak_max","porc_disponible"], ascending=[False,False,False])
                  .head(15)[["Colaborador","dias_disponibles"]])
    fig2, ax2 = plt.subplots(figsize=(8,6))
    ax2.barh(top["Colaborador"][::-1], top["dias_disponibles"][::-1])
    ax2.set_xlabel("Días disponibles")
    ax2.set_title("Top 15 colaboradores por días disponibles")
    img2 = _fig_to_base64(fig2)

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8" />
<title>Informe KPI de Disponibilidad (simple)</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; margin: 24px; }}
  h1 {{ margin-bottom: 4px; }}
  .meta {{ color:#444; margin-bottom: 16px; }}
  .grid {{ display:grid; grid-template-columns: 1fr 1fr; gap:24px; }}
  .card {{ background:#f8f9fb; padding:16px; border-radius:12px; }}
  img {{ max-width:100%; height:auto; }}
</style>
</head>
<body>
  <h1>Informe KPI de Disponibilidad (simple)</h1>
  <div class="meta"><b>Rango:</b> {ini} a {fin}</div>

  <div class="grid">
    <div class="card">
      <h3>Distribución por criticidad</h3>
      <img src="data:image/png;base64,{img1}" alt="Criticidad">
    </div>
    <div class="card">
      <h3>Top 15 por días disponibles</h3>
      <img src="data:image/png;base64,{img2}" alt="Top 15">
    </div>
  </div>
</body>
</html>"""

    out_path = outdir / "informe_disponibilidad_kpi_simple.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return str(out_path)

# =========================
# Normalización de nombres/RUT (presentación)
# =========================

_PARTICULAS = {
    "de","del","la","las","los","y","da","das","do","dos","di","du",
    "van","von","mac","mc"
}

def _smart_title_es(s: str) -> str:
    """
    Título “inteligente” para español: respeta partículas (de/del/la/…),
    maneja guiones y múltiples espacios.
    """
    if not s:
        return ""
    s = " ".join(str(s).split())
    tokens = re.split(r"(\s+|-)", s.lower())
    out = []
    for t in tokens:
        if not t.strip() or t in (" ", "\t", "-"):
            out.append(t)
        elif t in _PARTICULAS:
            out.append(t)
        else:
            out.append(t[:1].upper() + t[1:])
    return "".join(out).replace(" - ", "-")

def normalize_nombre(n: str) -> str:
    """
    - “APELLIDO, NOMBRE” -> “Nombre Apellido”
    - Corrige all-caps / all-lower
    - Limpia espacios
    """
    if n is None:
        return ""
    n = " ".join(str(n).split())
    if "," in n:
        ap, no = [p.strip() for p in n.split(",", 1)]
        n = f"{no} {ap}"
    n = _smart_title_es(n)
    return n

def normalize_nombre_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).map(normalize_nombre)

def format_rut(rut: str) -> str:
    """
    Asume rut ya normalizado (solo dígitos y K/k) y devuelve '12.345.678-5'
    """
    if rut is None:
        return ""
    s = re.sub(r"[^0-9kK]", "", str(rut))
    if len(s) < 2:
        return str(rut)
    dv = s[-1].upper()
    num = s[:-1]
    grupos = []
    while num:
        grupos.append(num[-3:])
        num = num[:-3]
    num_fmt = ".".join(reversed(grupos))
    return f"{num_fmt}-{dv}"

def format_rut_series(s: pd.Series) -> pd.Series:
    return s.astype(str).map(format_rut)

def reorder_name_rut_first(
    df: pd.DataFrame,
    prefer_name=("Colaborador","nombre","Nombre"),
    rut_col: str = "RUT"
) -> pd.DataFrame:
    """
    Empuja Nombre y RUT al frente. Mantiene el resto de columnas igual.
    """
    cols = list(df.columns)
    name_col = next((c for c in prefer_name if c in df.columns), None)
    first = [c for c in (name_col, rut_col) if c and c in df.columns]
    tail = [c for c in cols if c not in first]
    return df[first + tail]

# =========================
# Pipeline principal
# =========================

def classify_all(
    df: pd.DataFrame,
    cfg: Dict[str, Any],
    ml=None,
    date_range: Optional[Tuple[dt.date, dt.date]] = None
) -> tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """
    Devuelve:
      - resumen (por RUT/Colaborador)
      - detalle (filas con flag de disponibilidad)
      - meta (dict con rango y umbrales usados)
    """
    d = _apply_filters(df.copy(), cfg)
    if "fecha" not in d.columns:
        raise ValueError("El dataset no contiene columna 'fecha'.")

    # Filtrado por ventana
    if date_range is not None:
        start, end = date_range
        mask = (d["fecha"] >= start) & (d["fecha"] <= end)
        d = d.loc[mask].copy()
    else:
        start = d["fecha"].min() if not d.empty else dt.date.today()
        end = d["fecha"].max() if not d.empty else dt.date.today()

    # Vacío en ventana
    if d.empty:
        resumen = pd.DataFrame(columns=[
            "RUT","Colaborador","Departamento","Ciudad Residencia",
            "dias_ventana","dias_disponibles","porc_disponible","streak_max","criticidad"
        ])
        detalle = d.assign(disp_flag=pd.Series(dtype=int))[[
            "Colaborador","Departamento","RUT","Ciudad Residencia","fecha","actividad","disp_flag"
        ]] if len(d.columns) else pd.DataFrame(columns=[
            "Colaborador","Departamento","RUT","Ciudad Residencia","fecha","actividad","disp_flag"
        ])
        meta = {
            "rango": (start.isoformat(), end.isoformat()),
            "umbrales": {
                "th_high_days": _DEFAULT_CFG["rules"]["high_days"],
                "th_low_days": _DEFAULT_CFG["rules"]["low_days"],
                "th_streak": _DEFAULT_CFG["rules"]["high_streak"],
                "window_days": (end - start).days + 1 if end and start else 0,
                "base_rules": cfg.get("rules", {})
            }
        }
        return resumen, detalle, meta

    # Flag de disponibilidad
    if ml is not None:
        try:
            d["disp_flag"] = ml.predict_flags(d["actividad"].astype(str))
        except Exception:
            d["disp_flag"] = d["actividad"].astype(str).map(lambda x: classify_activity_flag(x, cfg))
    else:
        d["disp_flag"] = d["actividad"].astype(str).map(lambda x: classify_activity_flag(x, cfg))

    window_days = (end - start).days + 1
    rules = cfg.get("rules", {})
    base_high = max(1, int(rules.get("high_days", 7)))
    base_low = max(0, int(rules.get("low_days", 4)))
    base_streak = max(1, int(rules.get("high_streak", 3)))

    if window_days >= 30:
        th_high_days = base_high
        th_low_days = base_low
        th_streak = base_streak
    else:
        th_high_days = max(1, math.ceil(base_high * window_days / 30.0))
        th_low_days = max(0, math.floor(base_low * window_days / 30.0))
        th_streak = min(base_streak, window_days)

    d = d.sort_values(["RUT", "fecha"]).reset_index(drop=True)

    rows = []
    for rut, g in d.groupby("RUT", dropna=False):
        g = g.copy()
        nombre = g["Colaborador"].iloc[0] if "Colaborador" in g.columns and not g.empty else ""
        depto = g["Departamento"].iloc[0] if "Departamento" in g.columns and not g.empty else ""
        ciudad = g["Ciudad Residencia"].iloc[0] if "Ciudad Residencia" in g.columns and not g.empty else ""

        disp_dates = sorted(g.loc[g["disp_flag"] == 1, "fecha"].dropna().unique().tolist())
        dias_disponibles = len(disp_dates)
        streak_max = _max_consecutive_days(disp_dates)

        if dias_disponibles >= th_high_days or streak_max >= th_streak:
            criticidad = "ALTA"
        elif dias_disponibles <= th_low_days:
            criticidad = "BAJA"
        else:
            criticidad = "MEDIA"

        porc = (dias_disponibles / window_days) * 100.0 if window_days > 0 else 0.0

        rows.append({
            "RUT": rut,
            "Colaborador": nombre,
            "Departamento": depto,
            "Ciudad Residencia": ciudad,
            "dias_ventana": window_days,
            "dias_disponibles": dias_disponibles,
            "porc_disponible": round(porc, 1),
            "streak_max": streak_max,
            "criticidad": criticidad,
        })

    resumen = pd.DataFrame(rows)

    # Orden requerido: ALTA → MEDIA → BAJA, luego días↓, racha↓, %
    if not resumen.empty:
        map_crit = {"ALTA": 0, "MEDIA": 1, "BAJA": 2}
        resumen["criticidad_orden"] = resumen["criticidad"].map(map_crit).fillna(99)
        resumen = resumen.sort_values(
            ["criticidad_orden", "dias_disponibles", "streak_max", "porc_disponible"],
            ascending=[True, False, False, False]
        ).drop(columns=["criticidad_orden"]).reset_index(drop=True)

    detalle = d[[
        "Colaborador", "Departamento", "RUT", "Ciudad Residencia", "fecha", "actividad", "disp_flag"
    ]].copy()

    meta = {
        "rango": (start.isoformat(), end.isoformat()),
        "umbrales": {
            "th_high_days": th_high_days,
            "th_low_days": th_low_days,
            "th_streak": th_streak,
            "window_days": window_days,
            "base_rules": rules,
        }
    }
    return resumen, detalle, meta

# =========================
# Informe KPI “bonito” (cuerpo de correo)
# =========================

def build_email_report_kpi(resumen: pd.DataFrame,
                           detalle: pd.DataFrame,
                           sin_df: pd.DataFrame,
                           meta: Dict[str, Any],
                           outdir: Path) -> str:
    """
    Genera 'informe_disponibilidad_kpi.html' con gráficos (pie, stacked por departamento,
    top ALTA y serie diaria) + resumen ejecutivo y firma. HTML listo para usar como
    cuerpo del correo (sin adjuntos).
    """
    import numpy as np  # local
    from datetime import datetime, timedelta

    def fig_to_b64(fig):
        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight", dpi=160)
        plt.close(fig)
        buf.seek(0)
        return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")

    # ---- Ventana / meta
    ini, fin = meta.get("rango", ("", ""))
    window_days = meta.get("umbrales", {}).get("window_days", 0)

    # ---- KPIs
    counts = resumen["criticidad"].value_counts().to_dict()
    alta = counts.get("ALTA", 0)
    media = counts.get("MEDIA", 0)
    baja = counts.get("BAJA", 0)
    total = max(1, alta + media + baja)
    pct_alta = round(alta * 100 / total, 1)
    pct_baja = round(baja * 100 / total, 1)

    # ---- Top ALTA / BAJA persistente
    top_alta = (resumen[resumen["criticidad"] == "ALTA"]
                .sort_values(["dias_disponibles", "streak_max", "porc_disponible"],
                             ascending=[False, False, False])
                .head(12)[["RUT", "Colaborador", "Departamento",
                           "dias_disponibles", "streak_max", "porc_disponible"]])

    baja_persist = resumen[(resumen["criticidad"] == "BAJA") &
                           (resumen["dias_disponibles"] <= 1)][
        ["RUT", "Colaborador", "Departamento", "dias_disponibles", "streak_max", "porc_disponible"]
    ]

    # ---- Sin carga
    sc_count = len(sin_df) if isinstance(sin_df, pd.DataFrame) else 0
    sc_top = (sin_df.groupby(["Colaborador", "RUT"]).size()
              .reset_index(name="dias_sin_carga")
              .sort_values("dias_sin_carga", ascending=False).head(10)) if sc_count else None

    # ---- Gráficos
    imgs = {}

    # Pie criticidad
    fig = plt.figure(figsize=(4.2, 4.2))
    labels, vals = ["ALTA", "MEDIA", "BAJA"], [alta, media, baja]
    colors = ["#12b886", "#f59f00", "#fa5252"]
    plt.pie(vals, labels=labels, autopct="%1.0f%%", startangle=140, colors=colors)
    plt.title("Distribución por criticidad")
    imgs["pie"] = fig_to_b64(fig)

    # Stacked por departamento
    if "Departamento" in resumen.columns and not resumen.empty:
        fig = plt.figure(figsize=(7.5, 4.6))
        piv = (resumen.pivot_table(index="Departamento", columns="criticidad",
                                   values="RUT", aggfunc="count", fill_value=0)
               .reindex(columns=["ALTA", "MEDIA", "BAJA"])
               .sort_values(by=["ALTA", "MEDIA"], ascending=False))
        X = np.arange(len(piv))
        plt.bar(X, piv["ALTA"], color="#12b886", label="ALTA")
        plt.bar(X, piv["MEDIA"], bottom=piv["ALTA"], color="#f59f00", label="MEDIA")
        plt.bar(X, piv["BAJA"], bottom=piv["ALTA"] + piv["MEDIA"], color="#fa5252", label="BAJA")
        plt.xticks(X, piv.index, rotation=30, ha="right")
        plt.ylabel("Personas"); plt.title("Criticidad por departamento"); plt.legend()
        imgs["stack_dep"] = fig_to_b64(fig)

    # Top disponibilidad (ALTA)
    if not top_alta.empty:
        fig = plt.figure(figsize=(8, 4.4))
        y = top_alta["dias_disponibles"].tolist()
        names = (top_alta["Colaborador"] + " (" + top_alta["RUT"].astype(str) + ")").tolist()
        plt.barh(range(len(y)), y, color="#228be6")
        plt.yticks(range(len(y)), names); plt.gca().invert_yaxis()
        plt.xlabel("Días disponibles en ventana"); plt.title("Top disponibilidad (ALTA)")
        imgs["top_alta"] = fig_to_b64(fig)

    # Serie diaria (si hay detalle con disp_flag)
    daily_img = ""
    if "fecha" in detalle.columns and "disp_flag" in detalle.columns and not detalle.empty:
        di = (detalle[detalle["disp_flag"] == 1]
              .groupby("fecha")["RUT"].nunique().reset_index(name="personas"))
        if len(di) > 0:
            fig = plt.figure(figsize=(8, 3.8))
            plt.plot(di["fecha"], di["personas"], linewidth=2.5)
            plt.xticks(rotation=30, ha="right")
            plt.ylabel("Personas disponibles"); plt.title("Disponibilidad diaria")
            daily_img = fig_to_b64(fig)

    # ---- Recomendaciones
    recs = []
    if "fecha" in detalle.columns and "disp_flag" in detalle.columns and not detalle.empty:
        di = (detalle[detalle["disp_flag"] == 1]
              .groupby("fecha")["RUT"].nunique().reset_index(name="personas"))
        if len(di) >= 2:
            x = np.arange(len(di)); y = di["personas"].values
            slope = np.polyfit(x, y, 1)[0]
            if slope > 0.05:
                recs.append("La disponibilidad diaria va <b>en aumento</b> — adelantar tareas si es posible.")
            elif slope < -0.05:
                recs.append("La disponibilidad muestra <b>tendencia a la baja</b> — planificar refuerzos y contingencias.")
            else:
                recs.append("Disponibilidad <b>estable</b> — mantener planificación y revisar cuellos puntuales.")
    if len(baja_persist) > 0:
        recs.append(f"<b>{len(baja_persist)}</b> personas con <b>BAJA</b> y ≤1 día disponible — abrir causa-raíz y reasignar.")
    if sc_count > 0:
        recs.append(f"<b>{sc_count}</b> registros sin actividad — reforzar captura/estandarización de ‘actividad’.")
    if pct_alta < 25:
        recs.append(f"ALTA bajo (<b>{pct_alta}%</b>) — evaluar redistribución de cargas y reemplazos.")
    else:
        recs.append(f"Buen nivel de ALTA (<b>{pct_alta}%</b>) — mantener prácticas de programación.")

    def tbl(df, cols=None, n=50):
        if df is None or df.empty:
            return '<div class="muted">Sin datos.</div>'
        if cols:
            df = df[cols]
        return df.head(n).to_html(index=False, border=0, classes="table")

    # ---- Metadatos “Resumen Ejecutivo”
    ahora = datetime.now()
    generado = ahora.strftime("%d %B %Y, %H:%M hrs")
    proxima = (ahora + timedelta(days=2)).strftime("%d %B %Y")
    periodo = f"{ini} - {fin}" if ini and fin else (ini or fin or "")

    # ---- Firma (personaliza aquí)
    nombre    = "Ariel Reyes"
    cargo     = "Jefe de Operaciones – Lubrisider Chile S.A."
    telefono  = "+56 9 8833 4248"
    email     = "areyes@lubrisider.cl"
    web       = "https://lubrisider.cl"
    direccion = "Sucre 1234, Antofagasta, Chile"
    whatsapp  = "https://wa.me/56988334248"
    linkedin  = "https://www.linkedin.com/company/lubrisider"
    brand     = "Lubrisider | Trafo Energy"

    firma_html = f"""
    <div class="card">
      <div style="font-size:16px;font-weight:700;margin-bottom:2px;">{nombre}</div>
      <div class="muted" style="margin-bottom:8px;">{cargo}</div>
      <div>📞 <a href="tel:{telefono}" style="color:#93c5fd;text-decoration:none">{telefono}</a>
          &nbsp;·&nbsp; 💬 <a href="{whatsapp}" style="color:#93c5fd;text-decoration:none">WhatsApp</a></div>
      <div>✉️ <a href="mailto:{email}" style="color:#93c5fd;text-decoration:none">{email}</a></div>
      <div>🌐 <a href="{web}" style="color:#93c5fd;text-decoration:none">{brand}</a>
          &nbsp;·&nbsp; 🔗 <a href="{linkedin}" style="color:#93c5fd;text-decoration:none">LinkedIn</a></div>
      <div>📍 {direccion}</div>
    </div>
    """

    # ---- HTML final (sin “Listo para pegar…”)
    html = f"""
<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Disponibilidad — Informe KPI</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body {{ font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin:0; padding:0; background:#0f172a; color:#e2e8f0; }}
.wrap {{ max-width:1100px; margin:0 auto; padding:24px; }}
.card {{ background:#111827; padding:20px; border-radius:16px; box-shadow:0 4px 20px rgba(0,0,0,.35); margin-bottom:16px; }}
.title {{ font-size:22px; margin:0 0 6px; }}
.muted {{ color:#94a3b8; }}
.kpis {{ display:flex; gap:12px; flex-wrap:wrap; margin-top:8px; }}
.chip {{ border-radius:14px; padding:10px 14px; font-weight:600; display:inline-block; }}
.chip.alta {{ background:#0ea5e91a; border:1px solid #10b98155; color:#34d399; }}
.chip.media {{ background:#f59e0b1a; border:1px solid #f59e0b55; color:#fbbf24; }}
.chip.baja {{ background:#ef44441a; border:1px solid #ef444455; color:#fb7185; }}
.chip.win  {{ background:#1f2937; border:1px solid #475569; color:#cbd5e1; }}
.grid {{ display:grid; grid-template-columns: 1fr 1fr; gap:16px; }}
.table {{ width:100%; border-collapse:collapse; font-size:13px; }}
.table th,.table td {{ padding:8px 10px; border-bottom:1px solid #1f2937; }}
.table thead th {{ color:#93c5fd; text-align:left; }}
img {{ border-radius:12px; background:#0b1220; }}
@media(max-width:900px){{ .grid{{ grid-template-columns:1fr; }} }}
</style></head>
<body><div class="wrap">

  <div class="card">
    <div class="title">Disponibilidad Operativa</div>
    <div class="muted">Rango: <b>{ini}</b> a <b>{fin}</b> · Ventana: <b>{window_days} días</b></div>
    <div class="kpis">
      <div class="chip alta">ALTA: {alta} ({pct_alta}%)</div>
      <div class="chip media">MEDIA: {media}</div>
      <div class="chip baja">BAJA: {baja} ({pct_baja}%)</div>
      <div class="chip win">Sin carga: {sc_count}</div>
    </div>
  </div>

  <div class="card">
    <div class="title">Resumen Ejecutivo</div>
    <div><b>Reporte Generado:</b> {generado} &nbsp;|&nbsp; <b>Próxima Actualización:</b> {proxima}</div>
    <div><b>Período Analizado:</b> {periodo} ({window_days} días completados de período dinámico)</div>
    <div><b>Metodología:</b> Solo días con Cargas confirmadas · Personal con Cargo Específico = TRUE · Alertas automáticas &gt; 2 días</div>
  </div>

  <div class="grid">
    <div class="card"><img src="{imgs.get('pie','')}" style="width:100%"></div>
    <div class="card"><img src="{imgs.get('stack_dep','')}" style="width:100%"></div>
  </div>

  <div class="grid">
    <div class="card"><img src="{imgs.get('top_alta','')}" style="width:100%"></div>
    <div class="card">{('<img src="'+daily_img+'" style="width:100%">') if daily_img else '<div class="muted">Sin serie diaria.</div>'}</div>
  </div>

  <div class="card">
    <div class="title">Sugerencias basadas en datos</div>
    <ul>{"".join(f"<li>{r}</li>" for r in recs)}</ul>
  </div>

  <div class="grid">
    <div class="card">
      <div class="title">Top disponibilidad (ALTA)</div>
      {tbl(top_alta)}
    </div>
    <div class="card">
      <div class="title">BAJA (≤1 día disponible)</div>
      {tbl(baja_persist)}
    </div>
  </div>

  <div class="card">
    <div class="title">Personas con más días sin carga</div>
    {tbl(sc_top, ["Colaborador","RUT","dias_sin_carga"]) if sc_top is not None else '<div class="muted">No aplica.</div>'}
  </div>

  {firma_html}

</div></body></html>
"""
    outdir.mkdir(parents=True, exist_ok=True)
    out_path = outdir / "informe_disponibilidad_kpi.html"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return str(out_path)
