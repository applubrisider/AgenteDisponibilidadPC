# src/agent/ml.py
from __future__ import annotations
import re
from pathlib import Path
from dataclasses import dataclass
from typing import Iterable, List, Optional, Tuple

import joblib
import numpy as np

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.svm import LinearSVC
    from sklearn.dummy import DummyClassifier
    from sklearn.pipeline import Pipeline
    from sklearn.base import BaseEstimator
except Exception as e:
    raise RuntimeError("Falta scikit-learn en el entorno. Instala requirements.txt dentro del venv.") from e


def _norm(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.lower().strip()
    s = s.replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").replace("ñ","n")
    s = re.sub(r"\s+", " ", s)
    return s

# Palabras clave por defecto para etiquetar "disponible" en el dataset de entrenamiento
DEFAULT_AVAILABLE_TOKENS = [
    "disponible", "stand by", "stand-by", "standby", "libre"
]

# Si aparece un código de servicio/lab/contrato generalmente NO es día disponible
DEFAULT_NOT_AVAILABLE_PATTERNS = [
    r"\bser-\d{4}-\d+\b",   # SER-2025-0154
    r"\blab-\d{4}-\d+\b",   # LAB-2025-0029
    r"\bcon-\d{4}-\d+\b",   # CON-2025-0156
]

def _label_from_text(t: str,
                     avail_tokens: List[str],
                     not_avail_ptrns: List[str]) -> int:
    """
    Devuelve 1 si disponible, 0 si no disponible.
    Heurística solo para crear etiquetas de entrenamiento a partir del propio dataset.
    """
    n = _norm(t)
    if not n:
        return 0
    for pat in not_avail_ptrns:
        if re.search(pat, n):
            return 0
    for tok in avail_tokens:
        if tok in n:
            return 1
    # Casos explícitos de NO disponible frecuentes
    no_tokens = [
        "vacaciones", "licencia", "descanso", "descanso en zona",
        "capacitacion", "capacitaciones", "capacitacion on line",
        "oficina central", "actividad interna", "teletrabajo",
        "gestion adm", "academia lubri"
    ]
    if any(tok in n for tok in no_tokens):
        return 0
    # Por defecto, conservador
    return 0

@dataclass
class ActivityClassifier:
    pipe: Pipeline
    classes_: np.ndarray

    # --------- Entrenamiento a partir del DataFrame ----------
    @staticmethod
    def train_from_dataframe(df, cfg: Optional[dict] = None) -> "ActivityClassifier":
        """
        Crea etiquetas binarias (1 disponible / 0 no disponible) a partir de df['actividad']
        usando una heurística, y entrena un clasificador sencillo (TF-IDF + LinearSVC).
        """
        if "actividad" not in df.columns:
            raise ValueError("El DataFrame no tiene la columna 'actividad'.")

        avail_tokens = DEFAULT_AVAILABLE_TOKENS.copy()
        not_avail_ptrns = DEFAULT_NOT_AVAILABLE_PATTERNS.copy()

        # Permite sobrescribir por config.yaml (opcional)
        # Ejemplo en config.yaml:
        # ml:
        #   available_tokens: ["disponible", "libre"]
        #   not_available_patterns: ["\\bser-\\d{4}-\\d+\\b"]
        if cfg and isinstance(cfg, dict):
            ml_cfg = cfg.get("ml", {})
            if isinstance(ml_cfg.get("available_tokens"), list):
                avail_tokens = [str(x).lower() for x in ml_cfg["available_tokens"]]
            if isinstance(ml_cfg.get("not_available_patterns"), list):
                not_avail_ptrns = [str(x) for x in ml_cfg["not_available_patterns"]]

        X_text = df["actividad"].astype(str).fillna("")
        y = np.array([_label_from_text(t, avail_tokens, not_avail_ptrns) for t in X_text], dtype=int)

        # Si solo hay una clase, usa Dummy (evita error en SVC)
        if len(set(y.tolist())) < 2:
            clf: BaseEstimator = DummyClassifier(strategy="most_frequent")
        else:
            clf = LinearSVC()

        pipe = Pipeline([
            ("tfidf", TfidfVectorizer(
                preprocessor=_norm,
                ngram_range=(1, 2),
                min_df=1,
                max_features=5000
            )),
            ("clf", clf),
        ])
        pipe.fit(X_text, y)
        classes = np.array([0, 1], dtype=int)
        return ActivityClassifier(pipe=pipe, classes_=classes)

    # --------- Predicción ----------
    def predict_available(self, texts: Iterable[str]) -> List[bool]:
        X = list(texts)
        if not X:
            return []
        y = self.pipe.predict(X)
        return [bool(v) for v in y]

    # --------- Persistencia ----------
    def save(self, path: Path | str) -> None:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        joblib.dump({
            "pipe": self.pipe,
            "classes_": self.classes_,
        }, path)

    @staticmethod
    def load(path: Path | str) -> "ActivityClassifier":
        obj = joblib.load(path)
        return ActivityClassifier(pipe=obj["pipe"], classes_=obj["classes_"])
