# server.py
# -*- coding: utf-8 -*-
import os, hmac, json, traceback
from pathlib import Path
from types import SimpleNamespace
from typing import Optional, Dict, Any

from fastapi import FastAPI, Header, HTTPException, Request

app = FastAPI(title="Agente Disponibilidad – Webhook")
OUT_DIR = Path("out"); OUT_DIR.mkdir(parents=True, exist_ok=True)

WEBHOOK_TOKEN = os.getenv("WEBHOOK_TOKEN", "xLq8R7p2ZkFhJvY9m3N4bW1eT5sU6gC0rA7oB9h").strip()

def _log_error(where: str, err: Exception, extra: Dict[str, Any] | None = None) -> None:
    fp = OUT_DIR / "webhook_errors.log"
    with fp.open("a", encoding="utf-8") as fh:
        fh.write("\n" + "="*80 + "\n")
        fh.write(f"[{where}] ERROR:\n")
        fh.write("".join(traceback.format_exception(type(err), err, err.__traceback__)))
        fh.write("\nEXTRA:\n")
        fh.write(json.dumps(extra or {}, ensure_ascii=False, indent=2, default=str) + "\n")

def _check_token(x_token: Optional[str]):
    if not x_token: raise HTTPException(status_code=401, detail="Falta X-Webhook-Token")
    if not hmac.compare_digest(x_token, WEBHOOK_TOKEN):
        raise HTTPException(status_code=403, detail="Token inválido")

def _get_pipeline():
    try:
        from agent.ui import run_pipeline
    except Exception as e:
        raise RuntimeError("No pude importar run_pipeline desde agent.ui") from e
    RunOptions = None
    try:
        from agent.ui import RunOptions as _RunOptions
        RunOptions = _RunOptions
    except Exception:
        pass
    return run_pipeline, RunOptions

def _build_options(payload: Dict[str, Any], RunOptions):
    common = dict(
        cfg_path=str(Path("config.yml")),
        start_date=payload.get("start"),
        end_date=payload.get("end"),
        as_of=payload.get("as_of"),
        normalize_names=payload.get("normalize_names", True),
        format_rut=payload.get("format_rut", True),
        reorder_name_rut=payload.get("reorder_name_rut", True),
        filter_valid_ruts=payload.get("filter_valid_ruts", True),
        download_from_github=payload.get("download_from_github", True),
        make_kpi_report=payload.get("make_kpi_report", True),
    )
    if RunOptions is not None:
        try:
            return RunOptions(**common)
        except Exception as e:
            _log_error("build_options.RunOptions", e, {"common": common})
    return SimpleNamespace(**common)

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/echo")
async def echo(request: Request, x_webhook_token: Optional[str] = Header(None)):
    _check_token(x_webhook_token)
    body = await request.json()
    headers = {k: v for k, v in request.headers.items()
               if k.lower() in ("content-type", "x-webhook-token")}
    return {"ok": True, "headers": headers, "body": body}

@app.post("/webhook")
async def webhook(request: Request, x_webhook_token: Optional[str] = Header(None)):
    _check_token(x_webhook_token)
    try:
        payload = await request.json()
    except Exception as e:
        _log_error("parse_json", e, {})
        raise HTTPException(status_code=400, detail="JSON inválido")

    try:
        run_pipeline, RunOptions = _get_pipeline()
    except Exception as e:
        _log_error("import_pipeline", e, {"payload": payload})
        raise HTTPException(status_code=500, detail="No se pudo importar el pipeline (agent.ui.run_pipeline)")

    try:
        opts = _build_options(payload, RunOptions)
        rango, outputs = run_pipeline(ui=None, opts=opts)

        def p(k: str):
            v = outputs.get(k); return str(v) if v else None

        return {
            "ok": True,
            "rango": rango,
            "outputs": {
                "normalizado_csv": p("norm"),
                "resumen_csv": p("resumen"),
                "detalle_csv": p("detalle"),
                "informe_html": p("informe"),
                "informe_kpi_html": p("informe_kpi"),
                "sin_carga_csv": p("sin_carga_csv"),
                "sin_carga_html": p("sin_carga_html"),
                "ruts_fuera": p("ruts_fuera"),
            },
        }
    except HTTPException:
        raise
    except Exception as e:
        _log_error("run_pipeline", e, {"payload": payload})
        raise HTTPException(status_code=500, detail="Error interno ejecutando pipeline")
