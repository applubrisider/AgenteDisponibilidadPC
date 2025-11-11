from __future__ import annotations
import os, re, unicodedata
from typing import Optional, Dict, Any, List, Set
import requests
from agent.logic import normalize_rut_value

def _normkey(k: str) -> str:
    k = (k or "").strip().lower()
    k = k.replace("ï»¿", "")
    k = "".join(c for c in unicodedata.normalize("NFKD", k) if not unicodedata.combining(c))
    k = re.sub(r"[\s_]+", "", k)
    return k

def _extract_table_id(table_url: str) -> Optional[str]:
    m = re.search(r"/t/([A-Za-z0-9]+)", table_url)
    return m.group(1) if m else None

def _pick_records(js: Any) -> List[Dict[str, Any]]:
    if isinstance(js, list):
        return js
    if isinstance(js, dict):
        for key in ("records", "items", "data", "rows"):
            val = js.get(key)
            if isinstance(val, list):
                return val
    return []

def fetch_valid_ruts(
    table_url: Optional[str] = None,
    table_id: Optional[str] = None,
    api_key: Optional[str] = None,
    rut_field: str = "rut",
    filter_field: Optional[str] = "estado_disponibilidad",
    filter_truthy: Optional[bool | int | str] = True,
    page_size: int = 200,
) -> Set[str]:
    if not table_id:
        table_id = _extract_table_id(table_url or "")
    if not table_id:
        return set()

    api_key = (api_key or os.getenv("ZAPIER_TABLES_API_KEY", "")).strip()
    if not api_key:
        raise RuntimeError("Falta ZAPIER_TABLES_API_KEY (.env o variable de entorno).")

    base = f"https://tables.zapier.com/api/v1/tables/{table_id}/records"
    session = requests.Session()

    headers_list = [
        {"Authorization": f"Bearer {api_key}"},
        {"X-API-Key": api_key},
    ]

    ruts: Set[str] = set()
    page = 1
    while True:
        ok = False
        last_exc: Exception | None = None
        for hdrs in headers_list:
            try:
                resp = session.get(base, params={"limit": page_size, "page": page}, headers=hdrs, timeout=30)
                if resp.status_code == 401 and "X-API-Key" not in hdrs:
                    continue
                resp.raise_for_status()
                data = resp.json()
                recs = _pick_records(data)
                ok = True
                break
            except Exception as e:
                last_exc = e
                continue
        if not ok:
            raise RuntimeError(f"No pude leer Zapier Tables (página {page}): {last_exc}")

        if not recs:
            break

        for rec in recs:
            content = rec.get("content", rec) if isinstance(rec, dict) else {}

            val = None
            for k, v in content.items():
                if _normkey(k) == _normkey(rut_field):
                    val = v; break
            if val is None:
                for k, v in content.items():
                    if _normkey(k) in ("rut","ruts","rutvalido","rutválido","rutvalidos","rutsvalidos"):
                        val = v; break

            if filter_field:
                fv = None
                for k, v in content.items():
                    if _normkey(k) == _normkey(filter_field):
                        fv = v; break
                if filter_truthy is not None:
                    want = str(filter_truthy).strip().lower()
                    got  = str(fv).strip().lower()
                    truthy = ("1","true","yes","si","sí")
                    if not (got == want or (want in truthy and got in truthy)):
                        continue

            rut_norm = normalize_rut_value(val)
            if rut_norm:
                ruts.add(rut_norm)

        if len(recs) < page_size:
            break
        page += 1

    return ruts
