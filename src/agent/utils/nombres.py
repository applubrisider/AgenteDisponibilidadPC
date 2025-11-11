import re

def normaliza_espacios(txt: str) -> str:
    return re.sub(r"\s+", " ", (txt or "").strip())

def nombre_y_paterno(fullname: str) -> str:
    """
    Devuelve 'PrimerNombre ApellidoPaterno'.
    Si sólo hay 2 tokens, usa el segundo como apellido.
    Si hay 3 o más, toma el penúltimo como paterno (evita materno).
    """
    if not fullname:
        return ""
    tks = normaliza_espacios(fullname).split(" ")
    if len(tks) == 1:
        return tks[0].title()
    if len(tks) == 2:
        return f"{tks[0].title()} {tks[1].title()}"
    # >= 3 tokens → penúltimo es paterno en la gran mayoría de los casos
    return f"{tks[0].title()} {tks[-2].title()}"
