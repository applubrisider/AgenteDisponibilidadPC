# src/agent/reporting.py
from __future__ import annotations
from pathlib import Path
from typing import Iterable, Optional, Dict
import pandas as pd
from datetime import datetime, date
from jinja2 import Environment, BaseLoader, select_autoescape

def _is_nullish(value) -> bool:
    if pd.isna(value):
        return True
    s = str(value).strip().lower()
    return s in ("", "null", "none", "nan")

def _parse_date(d: Optional[str]) -> Optional[date]:
    if not d:
        return None
    return datetime.fromisoformat(d).date()

def _normalize_fecha(df: pd.DataFrame) -> pd.DataFrame:
    if "fecha" in df.columns:
        df = df.copy()
        df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.date
    return df

def _apply_filters(
    df: pd.DataFrame,
    desde: Optional[str] = None,
    hasta: Optional[str] = None,
    deptos: Optional[Iterable[str]] = None,
) -> pd.DataFrame:
    df = _normalize_fecha(df)

    if desde and "fecha" in df.columns:
        d = _parse_date(desde)
        if d:
            df = df[df["fecha"] >= d]

    if hasta and "fecha" in df.columns:
        h = _parse_date(hasta)
        if h:
            df = df[df["fecha"] <= h]

    if deptos and "Departamento" in df.columns:
        allow = {s.strip() for s in deptos if str(s).strip()}
        if allow:
            df = df[df["Departamento"].isin(allow)]

    return df

def extract_sin_carga(df: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve solo filas sin carga (actividad nula/vacía/'null'/NaN).
    Conserva columnas clave si existen.
    """
    if "actividad" not in df.columns:
        raise ValueError("El DataFrame no tiene la columna 'actividad'.")

    cols = [
        c for c in ["Colaborador", "RUT", "Departamento", "Ciudad Residencia", "fecha", "actividad"]
        if c in df.columns
    ]
    mask = df["actividad"].map(_is_nullish)
    out = df.loc[mask, cols]
    out = _normalize_fecha(out)

    # Ordenar por fecha si existe
    if "fecha" in out.columns:
        sort_cols = ["fecha"]
        if "Departamento" in out.columns: sort_cols.append("Departamento")
        if "Colaborador" in out.columns: sort_cols.append("Colaborador")
        out = out.sort_values(sort_cols, na_position="last")
    return out

def _render_html_sin_carga(
    base: pd.DataFrame,
    por_persona: pd.DataFrame,
    por_dia: Optional[pd.DataFrame],
    por_depto: Optional[pd.DataFrame],
    outdir: Path,
    filtros: Dict[str, str | list | None],
) -> Path:
    """Construye un HTML compacto con métricas + tablas y links a los CSV."""
    env = Environment(
        loader=BaseLoader(),
        autoescape=select_autoescape(["html", "xml"]),
        trim_blocks=True,
        lstrip_blocks=True,
    )

    template = env.from_string(r"""
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Informe — Sin Carga</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{
  --bg:#0b0f14; --card:#10161d; --muted:#8aa0b4; --fg:#e6eef6; --accent:#59c3ff;
  --ok:#2ecc71; --warn:#f1c40f; --bad:#e74c3c; --line:#1e2935;
}
*{box-sizing:border-box}
body{margin:0; font-family:Segoe UI, Roboto, Arial; background:var(--bg); color:var(--fg);}
.wrap{max-width:1200px; margin:24px auto; padding:0 16px;}
h1{font-size:26px; margin:8px 0 16px;}
.card{background:var(--card); border:1px solid var(--line); border-radius:14px; padding:16px; margin:16px 0; box-shadow:0 4px 16px #00000033;}
.kpis{display:grid; grid-template-columns: repeat(4, minmax(180px,1fr)); gap:12px;}
.kpi{background:#0e141b; border:1px solid var(--line); border-radius:12px; padding:12px;}
.kpi .label{font-size:12px; color:var(--muted);}
.kpi .value{font-size:20px; font-weight:600;}
.meta{color:var(--muted); font-size:13px;}
table{width:100%; border-collapse:collapse; border:1px solid var(--line); border-radius:12px; overflow:hidden;}
th, td{padding:10px 12px; border-bottom:1px solid var(--line); text-align:left; font-size:14px;}
th{background:#0d131a; color:#b6c6d6; position:sticky; top:0;}
tbody tr:hover{background:#0f171f;}
.badge{display:inline-block; padding:2px 8px; border-radius:999px; font-size:12px; border:1px solid var(--line); color:var(--muted);}
.links a{color:var(--accent); text-decoration:none;}
.links a:hover{text-decoration:underline;}
.footer{color:var(--muted); font-size:12px; margin-top:24px;}
.small{font-size:12px;}
@media (max-width:900px){ .kpis{grid-template-columns: repeat(2, 1fr);} }
</style>
</head>
<body>
<div class="wrap">
  <h1>Informe de Personas <span class="badge">Sin Carga</span></h1>
  <div class="meta">
    Generado: {{ now }} · Archivos en: <span class="small">{{ outdir }}</span><br>
    Filtros — Desde: <strong>{{ filtros.desde or '—' }}</strong>,
    Hasta: <strong>{{ filtros.hasta or '—' }}</strong>,
    Deptos: <strong>{{ filtros.deptos if filtros.deptos else 'Todos' }}</strong>
  </div>

  <div class="card">
    <div class="kpis">
      <div class="kpi"><div class="label">Filas “sin carga”</div><div class="value">{{ kpi.filas }}</div></div>
      <div class="kpi"><div class="label">Personas únicas</div><div class="value">{{ kpi.personas }}</div></div>
      <div class="kpi"><div class="label">Días distintos</div><div class="value">{{ kpi.dias }}</div></div>
      <div class="kpi"><div class="label">Departamentos</div><div class="value">{{ kpi.deptos }}</div></div>
    </div>
    <div class="links" style="margin-top:8px">
      CSV: <a href="sin_carga.csv">sin_carga.csv</a> ·
      <a href="sin_carga_por_persona.csv">sin_carga_por_persona.csv</a>
      {% if has_por_dia %} · <a href="sin_carga_por_dia.csv">sin_carga_por_dia.csv</a>{% endif %}
      {% if has_por_depto %} · <a href="sin_carga_por_depto.csv">sin_carga_por_depto.csv</a>{% endif %}
    </div>
  </div>

  <div class="card">
    <h2 style="margin:0 0 8px">Top personas con más días sin carga</h2>
    <table>
      <thead>
        <tr>
          {% if 'Colaborador' in por_persona.columns %}<th>Colaborador</th>{% endif %}
          {% if 'RUT' in por_persona.columns %}<th>RUT</th>{% endif %}
          {% if 'Departamento' in por_persona.columns %}<th>Departamento</th>{% endif %}
          <th>Días sin carga</th>
        </tr>
      </thead>
      <tbody>
      {% for _, r in por_persona.head(150).iterrows() %}
        <tr>
          {% if 'Colaborador' in por_persona.columns %}<td>{{ r['Colaborador'] or '' }}</td>{% endif %}
          {% if 'RUT' in por_persona.columns %}<td>{{ r['RUT'] or '' }}</td>{% endif %}
          {% if 'Departamento' in por_persona.columns %}<td>{{ r['Departamento'] or '' }}</td>{% endif %}
          <td><strong>{{ r['dias_sin_carga'] }}</strong></td>
        </tr>
      {% endfor %}
      </tbody>
    </table>
    <div class="small" style="margin-top:6px">Mostrando top 150. Ver CSV para listado completo.</div>
  </div>

  {% if por_dia is not none %}
  <div class="card">
    <h2 style="margin:0 0 8px">Sin carga por día{% if 'Departamento' in por_dia.columns %} y departamento{% endif %}</h2>
    <table>
      <thead>
        <tr>
          <th>Fecha</th>
          {% if 'Departamento' in por_dia.columns %}<th>Departamento</th>{% endif %}
          <th>Sin carga</th>
        </tr>
      </thead>
      <tbody>
      {% for _, r in por_dia.iterrows() %}
        <tr>
          <td>{{ r['fecha'] }}</td>
          {% if 'Departamento' in por_dia.columns %}<td>{{ r['Departamento'] or '' }}</td>{% endif %}
          <td><strong>{{ r['sin_carga'] }}</strong></td>
        </tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
  {% endif %}

  {% if por_depto is not none %}
  <div class="card">
    <h2 style="margin:0 0 8px">Acumulado por departamento</h2>
    <table>
      <thead><tr><th>Departamento</th><th>Sin carga</th></tr></thead>
      <tbody>
      {% for _, r in por_depto.iterrows() %}
        <tr><td>{{ r['Departamento'] }}</td><td><strong>{{ r['sin_carga'] }}</strong></td></tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
  {% endif %}

  <div class="footer">
    Lubrisider · Reporte generado automáticamente · {{ now }}
  </div>
</div>
</body>
</html>
    """)

    # KPIs
    kpi = {
        "filas": len(base),
        "personas": base["RUT"].nunique() if "RUT" in base.columns else base["Colaborador"].nunique() if "Colaborador" in base.columns else len(base),
        "dias": base["fecha"].nunique() if "fecha" in base.columns else "—",
        "deptos": base["Departamento"].nunique() if "Departamento" in base.columns else "—",
    }

    has_por_dia = por_dia is not None
    has_por_depto = por_depto is not None

    html = template.render(
        now=datetime.now().strftime("%Y-%m-%d %H:%M"),
        outdir=str(outdir),
        filtros=filtros,
        kpi=kpi,
        por_persona=por_persona,
        por_dia=por_dia,
        por_depto=por_depto,
        has_por_dia=has_por_dia,
        has_por_depto=has_por_depto,
    )
    out_path = outdir / "informe_sin_carga.html"
    out_path.write_text(html, encoding="utf-8")
    return out_path

def export_sin_carga_reports(
    df: pd.DataFrame,
    outdir: Path,
    desde: Optional[str] = None,
    hasta: Optional[str] = None,
    deptos: Optional[Iterable[str]] = None,
) -> Dict[str, Path]:
    """
    Genera:
      - sin_carga.csv
      - sin_carga_por_persona.csv
      - sin_carga_por_dia.csv (si hay 'fecha')
      - sin_carga_por_depto.csv (si hay 'Departamento')
      - informe_sin_carga.html
    """
    outdir.mkdir(parents=True, exist_ok=True)

    base = extract_sin_carga(df)
    base = _apply_filters(base, desde=desde, hasta=hasta, deptos=deptos)

    paths: Dict[str, Path] = {}

    # 1) Detalle sin carga
    p_det = outdir / "sin_carga.csv"
    base.to_csv(p_det, index=False, encoding="utf-8-sig")
    paths["sin_carga"] = p_det

    # 2) Por persona
    cols_persona = [c for c in ["Colaborador", "RUT", "Departamento"] if c in base.columns]
    por_persona = base.groupby(cols_persona, dropna=False).size().reset_index(name="dias_sin_carga")
    por_persona = por_persona.sort_values("dias_sin_carga", ascending=False)
    p_persona = outdir / "sin_carga_por_persona.csv"
    por_persona.to_csv(p_persona, index=False, encoding="utf-8-sig")
    paths["sin_carga_por_persona"] = p_persona

    # 3) Por día
    por_dia = None
    if "fecha" in base.columns:
        cols = ["fecha"] + (["Departamento"] if "Departamento" in base.columns else [])
        por_dia = base.groupby(cols, dropna=False).size().reset_index(name="sin_carga")
        por_dia = por_dia.sort_values(cols)
        p_dia = outdir / "sin_carga_por_dia.csv"
        por_dia.to_csv(p_dia, index=False, encoding="utf-8-sig")
        paths["sin_carga_por_dia"] = p_dia

    # 4) Por departamento total
    por_depto = None
    if "Departamento" in base.columns:
        por_depto = base.groupby(["Departamento"], dropna=False).size().reset_index(name="sin_carga")
        por_depto = por_depto.sort_values("sin_carga", ascending=False)
        p_depto = outdir / "sin_carga_por_depto.csv"
        por_depto.to_csv(p_depto, index=False, encoding="utf-8-sig")
        paths["sin_carga_por_depto"] = p_depto

    # 5) HTML
    filtros = {
        "desde": desde,
        "hasta": hasta,
        "deptos": ", ".join(deptos) if deptos else None
    }
    p_html = _render_html_sin_carga(
        base=base,
        por_persona=por_persona,
        por_dia=por_dia,
        por_depto=por_depto,
        outdir=outdir,
        filtros=filtros,
    )
    paths["informe_sin_carga_html"] = p_html

    return paths
