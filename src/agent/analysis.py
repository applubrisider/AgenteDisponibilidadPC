# src/agent/analysis.py
from __future__ import annotations
import re
import math
import datetime as dt
from pathlib import Path
from typing import Dict, Any, Tuple, Optional, List

import pandas as pd


# ---------- Helpers de normalización ----------
def _norm(s: str) -> str:
    if not isinstance(s, str):
        return ""
    t = s.strip().lower()
    t = (t.replace("á","a").replace("é","e").replace("í","i")
           .replace("ó","o").replace("ú","u").replace("ñ","n"))
    return t

def _is_oper(depto: str) -> bool:
    d = _norm(depto)
    # flexibilizamos por si vienen variaciones: "oper_of", "oper of", etc.
    return ("oper" in d) or ("oper_of" in d) or ("oper of" in d)

def _short_name(fullname: str) -> str:
    # Nombre Apellido (si viene muy largo)
    parts = [p for p in str(fullname).split() if p]
    if len(parts) >= 2:
        return f"{parts[0]} {parts[-1]}"
    return str(fullname).strip()

def _fmt_date(d: dt.date) -> str:
    return d.strftime("%d-%m-%Y")

# ---------- Detección de proyectos ----------
_PROY_PAT = re.compile(r"\b(?:SER|CON|LAB)-\d{4}-\d{4}\b", re.IGNORECASE)

def _extract_projects(series: pd.Series) -> List[str]:
    found: set[str] = set()
    for txt in series.dropna().astype(str):
        for m in _PROY_PAT.findall(txt):
            found.add(m.upper())
    return sorted(found)

# ---------- Disponibles por día ----------
def _daily_availability(detalle: pd.DataFrame) -> pd.DataFrame:
    """
    Recibe 'detalle' con columnas ['fecha','disp_flag','Colaborador'].
    Devuelve df con columnas: fecha, disponibles, nombres (string corto con tope).
    """
    d = detalle.copy()
    d = d[d["disp_flag"] == 1].copy()
    if d.empty:
        return pd.DataFrame(columns=["fecha","disponibles","nombres"])

    # nombres cortos únicos por día, en orden alfabético
    grp = d.groupby("fecha")
    rows = []
    for f, g in grp:
        names = sorted({_short_name(n) for n in g["Colaborador"].dropna().astype(str)})
        disp = len(names)
        # Limita la lista para que el correo sea legible
        MAXN = 12
        if len(names) > MAXN:
            vis = ", ".join(names[:MAXN])
            vis = f"{vis}, y {len(names) - MAXN} más"
        else:
            vis = ", ".join(names)
        rows.append({"fecha": f, "disponibles": disp, "nombres": vis})
    out = pd.DataFrame(rows).sort_values("fecha").reset_index(drop=True)
    return out

# ---------- Alertas (licencia / vacaciones todo el período) ----------
def _alerts_full_period(detalle: pd.DataFrame, start: dt.date, end: dt.date) -> List[str]:
    """
    Persona cae en alerta si TODAS sus actividades del rango son licencia médica
    o vacaciones. (Se ignoran filas sin actividad)
    """
    d = detalle.copy()
    d = d[(d["fecha"] >= start) & (d["fecha"] <= end)]
    if d.empty:
        return []

    lic_set = {"licencia medica", "licencia médica"}
    vac_set = {"vacaciones"}

    al_names: List[str] = []

    for rut, g in d.groupby("RUT"):
        acts = {_norm(x) for x in g["actividad"].dropna().astype(str)}
        if not acts:
            continue
        # Si todos los valores están dentro de licencia o todos dentro de vacaciones:
        if acts.issubset(lic_set) or acts.issubset(vac_set):
            nombre = g["Colaborador"].iloc[0] if "Colaborador" in g.columns else str(rut)
            al_names.append(nombre)

    # Orden alfabético
    return sorted(al_names, key=lambda s: _short_name(s))

# ---------- Tendencia (simple) ----------
def _trend_from_series(counts: List[int]) -> str:
    if len(counts) < 2:
        return "Estable"
    first, last = counts[0], counts[-1]
    diff = last - first
    if diff > 1:
        return "Creciente"
    elif diff < -1:
        return "Decreciente"
    else:
        return "Estable"

# ---------- Construcción del análisis ----------
def build_exec_analysis(
    df: pd.DataFrame,
    detalle: pd.DataFrame,
    start: dt.date,
    end: dt.date
) -> Dict[str, Any]:
    """
    Devuelve dict con:
      - 'html' (str)
      - 'text' (str)
      - 'path_html' (ruta si se guardó)
      - 'meta' (fechas, cobertura, totales)
    """
    # Cobertura de datos
    datos_hasta: Optional[dt.date] = None
    if "fecha" in df.columns and not df.empty:
        datos_hasta = max(df["fecha"])
    periodo_fin_efectivo = min(end, datos_hasta) if datos_hasta else end

    # Personal operativo
    oper_df = df[df["Departamento"].map(_is_oper, na_action="ignore")].copy() \
             if "Departamento" in df.columns else df.copy()
    oper_personas = oper_df[["RUT","Colaborador"]].drop_duplicates()
    total_oper = len(oper_personas)

    # Proyectos activos detectados en el rango
    df_rango = df[(df["fecha"] >= start) & (df["fecha"] <= periodo_fin_efectivo)].copy()
    proyectos = _extract_projects(df_rango["actividad"]) if "actividad" in df_rango.columns else []

    # Disponibles por día (solo dentro del rango efectivo)
    det_rango = detalle[(detalle["fecha"] >= start) & (detalle["fecha"] <= periodo_fin_efectivo)].copy()
    daily = _daily_availability(det_rango)
    prom = round(daily["disponibles"].mean(), 1) if not daily.empty else 0.0
    max_row = daily.loc[daily["disponibles"].idxmax()] if not daily.empty else None
    min_row = daily.loc[daily["disponibles"].idxmin()] if not daily.empty else None
    tendencia = _trend_from_series(daily["disponibles"].tolist()) if not daily.empty else "Estable"

    # Alertas
    alertas = _alerts_full_period(detalle, start, periodo_fin_efectivo)

    # Lista mostrable de personal (primeros 6 + “y X+”)
    nombres_oper = [_short_name(n) for n in oper_personas["Colaborador"].dropna().astype(str)]
    nombres_oper = sorted(set(nombres_oper))
    head = ", ".join(nombres_oper[:6])
    tail = f" — y {max(0, len(nombres_oper)-6)} más" if len(nombres_oper) > 6 else ""

    # Fechas bonitas
    hoy_ts = dt.datetime.now()
    fecha_analisis = hoy_ts.strftime("%d de %B %Y, %H:%M hrs")
    datos_hasta_str = _fmt_date(datos_hasta) if datos_hasta else "s/d"
    per_ini_str = _fmt_date(start)
    per_fin_str = _fmt_date(periodo_fin_efectivo)

    # Construcción de la tabla por fecha
    filas_tbl = []
    for _, r in daily.iterrows():
        filas_tbl.append(
            f"<tr><td>{_fmt_date(r['fecha'])}</td><td style='text-align:center'>{int(r['disponibles'])}</td>"
            f"<td>{r['nombres']}</td></tr>"
        )
    html_tbl = (
        "<table style='border-collapse:collapse;width:100%;font-size:14px'>"
        "<thead><tr>"
        "<th style='text-align:left;border-bottom:1px solid #e5e5e5;padding:6px 8px'>Fecha</th>"
        "<th style='text-align:center;border-bottom:1px solid #e5e5e5;padding:6px 8px'>Disponibles</th>"
        "<th style='text-align:left;border-bottom:1px solid #e5e5e5;padding:6px 8px'>Colaboradores específicos</th>"
        "</tr></thead><tbody>"
        + "".join(f"<tr><td style='padding:6px 8px'>{row.split('</td>')[0][4:]}"
                  + "</td>" + row.split("</td>",1)[1] for row in filas_tbl)  # ya tiene <td>...<td>...
        + "</tbody></table>"
    ) if filas_tbl else "<i>No hay disponibilidad diaria en el rango solicitado.</i>"

    # Proyectos como bullets
    html_proy = "<ul>" + "".join(f"<li>{p}</li>" for p in proyectos) + "</ul>" if proyectos else "<i>Sin códigos detectados.</i>"

    # Alertas como bullets
    html_alertas = "<ul>" + "".join(f"<li>{_short_name(n)}</li>" for n in alertas) + "</ul>" if alertas else "<i>Sin alertas en el período.</i>"

    # HTML principal (estilo parecido a tu ejemplo)
    html = f"""
<div style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; line-height:1.4">
  <h2>📊 ANÁLISIS PERSONAL OPERATIVO LUBRISIDER CHILE</h2>
  <div><b>Fecha de Análisis:</b> {fecha_analisis}</div>
  <div>🔍 <b>DATOS DISPONIBLES HASTA:</b> {datos_hasta_str}</div>
  <div>📅 <b>PERÍODO ANALIZADO:</b> {per_ini_str} al {per_fin_str} (datos disponibles)</div>
  <hr/>

  <h3>👥 PERSONAL OPERATIVO IDENTIFICADO ({total_oper} colaboradores OPER/OPER_OF)</h3>
  <div>{head}{tail}</div>

  <h3>🎯 PROYECTOS ACTIVOS IDENTIFICADOS</h3>
  {html_proy}

  <h3>📅 DISPONIBILIDAD ESPECÍFICA POR FECHA</h3>
  {html_tbl}

  <h3>🚨 ALERTAS INDIVIDUALES CRÍTICAS</h3>
  {html_alertas}

  <h3>📈 MÉTRICAS OPERATIVAS</h3>
  <ul>
    <li><b>Promedio diario:</b> {prom} colaboradores disponibles</li>
    <li><b>Día con mayor disponibilidad:</b> {(_fmt_date(max_row['fecha']) + f" ({int(max_row['disponibles'])})") if max_row is not None else "s/d"}</li>
    <li><b>Día con menor disponibilidad:</b> {(_fmt_date(min_row['fecha']) + f" ({int(min_row['disponibles'])})") if min_row is not None else "s/d"}</li>
    <li><b>Tendencia:</b> {tendencia}</li>
    <li><b>Total colaboradores operativos:</b> {total_oper}</li>
  </ul>

  <h3>⚠️ LIMITACIÓN IMPORTANTE</h3>
  <div>Los datos están disponibles hasta <b>{datos_hasta_str}</b>. Para un análisis completo del período <b>{per_ini_str}</b> a <b>{_fmt_date(end)}</b>, se requiere actualización del archivo fuente.</div>
</div>
""".strip()

    # Versión texto (fallback)
    text = [
        f"ANÁLISIS PERSONAL OPERATIVO LUBRISIDER CHILE",
        f"Fecha de Análisis: {fecha_analisis}",
        f"DATOS DISPONIBLES HASTA: {datos_hasta_str}",
        f"PERÍODO ANALIZADO: {per_ini_str} al {per_fin_str}",
        "",
        f"PERSONAL OPERATIVO IDENTIFICADO: {total_oper}. Ejemplos: {head}{tail}",
        "",
        "PROYECTOS ACTIVOS IDENTIFICADOS:",
        (" - " + "\n - ".join(proyectos)) if proyectos else " - Sin códigos detectados.",
        "",
        "DISPONIBILIDAD POR FECHA:",
    ]
    for _, r in daily.iterrows():
        text.append(f" - {_fmt_date(r['fecha'])}: {int(r['disponibles'])} -> {r['nombres']}")
    if daily.empty:
        text.append(" - (Sin datos en el rango)")

    text += [
        "",
        "ALERTAS INDIVIDUALES CRÍTICAS:",
        (" - " + "\n - ".join(alertas)) if alertas else " - Sin alertas.",
        "",
        "MÉTRICAS:",
        f" - Promedio diario: {prom}",
        f" - Tendencia: {tendencia}",
        f" - Total colaboradores operativos: {total_oper}",
    ]
    text_body = "\n".join(text)

    return {
        "html": html,
        "text": text_body,
        "meta": {
            "ini": start.isoformat(),
            "fin": end.isoformat(),
            "fin_efectivo": periodo_fin_efectivo.isoformat(),
            "datos_hasta": datos_hasta.isoformat() if datos_hasta else None,
            "total_oper": total_oper,
            "proyectos": proyectos
        }
    }

def save_exec_analysis_html(analysis: Dict[str, Any], outdir: Path) -> str:
    outdir.mkdir(parents=True, exist_ok=True)
    path = outdir / "analisis_ejecutivo.html"
    with open(path, "w", encoding="utf-8") as f:
        f.write(analysis["html"])
    return str(path.resolve())
