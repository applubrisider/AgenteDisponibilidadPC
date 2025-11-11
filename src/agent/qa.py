import os
import pandas as pd
from dotenv import load_dotenv

# IA opcional para Q&A: intenta OpenAI; si no, intenta Ollama local (HTTP)
def _ask_openai(prompt: str) -> str | None:
    try:
        from openai import OpenAI
        key = os.getenv("OPENAI_API_KEY")
        if not key:
            return None
        client = OpenAI(api_key=key)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role":"system","content":"Eres un analista de datos conciso."},
                      {"role":"user","content":prompt}],
            temperature=0.2
        )
        return resp.choices[0].message.content
    except Exception:
        return None

def _ask_ollama(prompt: str) -> str | None:
    try:
        import requests, json
        if os.getenv("USE_OLLAMA","false").lower()!="true":
            return None
        model = os.getenv("OLLAMA_MODEL","llama3.1")
        r = requests.post("http://localhost:11434/api/generate",
                          json={"model": model, "prompt": prompt, "stream": False},
                          timeout=120)
        if r.ok:
            data = r.json()
            return data.get("response")
    except Exception:
        pass
    return None

def qa_console(resumen_csv):
    load_dotenv()
    df = pd.read_csv(resumen_csv, dtype=str)
    print("\nConsola de preguntas (escribe 'salir' para terminar).")
    while True:
        q = input("\n¿Tu pregunta sobre el resumen? > ").strip()
        if q.lower() in ("salir","exit","quit"):
            break
        # pequeño contexto tabular
        head = df.head(20).to_markdown(index=False)
        prompt = f"""Tienes el siguiente resumen (primeras filas):
{head}

Pregunta: {q}

Responde de forma concisa y, si corresponde, sugiere filtros o columnas útiles.
"""
        ans = _ask_openai(prompt) or _ask_ollama(prompt) or "No se pudo usar IA (falta configuración). Revisa .env."
        print(f"\n{ans}\n")
