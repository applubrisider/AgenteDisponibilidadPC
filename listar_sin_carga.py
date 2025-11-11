import pandas as pd
from pathlib import Path

src = r'.\disponibilidad_tidy.csv'
outdir = Path(r'.\out'); outdir.mkdir(parents=True, exist_ok=True)

# Lectura robusta del CSV (prueba codificaciones y autodetección de separador)
df = None
for enc in ('utf-8-sig', 'utf-8', 'latin1'):
    try:
        df = pd.read_csv(src, sep=None, engine='python', dtype=str, encoding=enc)
        break
    except Exception:
        pass
if df is None:
    raise SystemExit('No pude leer el CSV (revisa ruta/codificación).')

def is_nullish(x):
    if pd.isna(x): 
        return True
    s = str(x).strip().lower()
    return s in ('', 'null', 'none', 'nan')

if 'actividad' not in df.columns:
    raise SystemExit("El CSV no tiene la columna 'actividad'.")

mask = df['actividad'].map(is_nullish)
cols = [c for c in ['Colaborador','RUT','Departamento','Ciudad Residencia','fecha','actividad'] if c in df.columns]
out = df.loc[mask, cols].sort_values(['fecha'] if 'fecha' in df.columns else cols[:1])

out.to_csv(outdir / 'sin_carga.csv', index=False, encoding='utf-8-sig')
print(f"Se encontraron {len(out)} filas sin carga. Archivo: {outdir/'sin_carga.csv'}")
