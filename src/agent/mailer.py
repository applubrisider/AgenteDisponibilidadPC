import os
import requests

def fetch_valid_ruts() -> set[str]:
    """
    Devuelve el set de RUT válidos de operación (Zapier Tables).
    Implementación real: ajusta URL/headers a tu tabla/acción.
    """
    url = os.getenv("ZAPIER_TABLES_VALID_RUTS_URL")  # p.ej. un Webhook/REST que retorna JSON con {"rut": "..."}
    token = os.getenv("ZAPIER_TABLES_TOKEN")
    if not url:
        # Fallback: carga de CSV local si estás offline
        import pandas as pd
        csv_path = os.getenv("VALID_RUTS_CSV", "data/ruts_validos.csv")
        if os.path.exists(csv_path):
            ruts = pd.read_csv(csv_path, dtype=str)["RUT"].dropna().str.strip().unique().tolist()
            return set(ruts)
        return set()

    headers = {"Authorization": f"Bearer {token}"} if token else {}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    # adapta a tu estructura real
    ruts = []
    for row in data:
        rut = str(row.get("rut") or row.get("RUT") or "").strip()
        if rut:
            ruts.append(rut)
    return set(ruts)
