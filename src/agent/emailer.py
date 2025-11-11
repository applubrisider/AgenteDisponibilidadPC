# src/agent/emailer.py
from __future__ import annotations
import os
import mimetypes
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, Dict, Any, List, Tuple

def _as_list(x) -> List[str]:
    if not x:
        return []
    if isinstance(x, str):
        return [t.strip() for t in x.split(",") if t.strip()]
    return [str(t).strip() for t in x if str(t).strip()]

def _guess_filename(p: Path) -> Tuple[bytes, str, str]:
    ctype, _ = mimetypes.guess_type(str(p))
    if ctype is None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    return p.read_bytes(), maintype, subtype

def _send_with_outlook(to: List[str], cc: List[str], subject: str, html_body: str, attachments: Iterable[Path]) -> str:
    try:
        import win32com.client as win32  # requiere: pip install pywin32
    except Exception as e:
        return f"Outlook no disponible: {e}"
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.To = ";".join(to)
        if cc:
            mail.CC = ";".join(cc)
        mail.HTMLBody = html_body
        for p in attachments:
            mail.Attachments.Add(str(p))
        mail.Send()
        return "Correo enviado vía Outlook."
    except Exception as e:
        return f"Fallo enviando con Outlook: {e}"

def _send_with_smtp(cfg_email: Dict[str, Any], to: List[str], cc: List[str], subject: str, html_body: str, attachments: Iterable[Path]) -> str:
    smtp = cfg_email.get("smtp", {}) or {}
    host  = smtp.get("host", "smtp.office365.com")
    port  = int(smtp.get("port", 587))
    user  = smtp.get("user", "")
    pwd   = smtp.get("password", "") or os.environ.get("EMAIL_PASS", "")
    starttls = bool(smtp.get("starttls", True))

    if not user or not pwd:
        return "SMTP no configurado (user/password)."

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)

    msg.set_content("Su cliente no soporta HTML. Revise adjuntos.")
    msg.add_alternative(html_body, subtype="html")

    for p in attachments:
        data, maintype, subtype = _guess_filename(p)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=p.name)

    rcpts = list(set(to + cc))
    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            if starttls:
                s.starttls()
            s.login(user, pwd)
            s.send_message(msg, from_addr=user, to_addrs=rcpts)
        return "Correo enviado vía SMTP."
    except Exception as e:
        return f"Fallo enviando con SMTP: {e}"

def send_report_email(cfg: Dict[str, Any], attachments: Iterable[Path], meta: Dict[str, Any],
                      extra_to: str | Iterable[str] = "", extra_cc: str | Iterable[str] = "") -> str:
    email_cfg = (cfg or {}).get("email", {}) or {}

    to = _as_list(email_cfg.get("to"))
    cc = _as_list(email_cfg.get("cc"))
    to += _as_list(extra_to)
    cc += _as_list(extra_cc)
    to = list(dict.fromkeys(to))
    cc = list(dict.fromkeys(cc))

    if not to:
        return "Sin destinatarios. Configure email.to en config.yaml o use los campos Extra To/CC."

    ini, fin = (meta or {}).get("rango", ("", ""))
    subject = (email_cfg.get("subject_template") or "Informe de disponibilidad {ini}–{fin}").format(ini=ini, fin=fin)
    body_intro = email_cfg.get("body_intro") or "Adjunto informe de disponibilidad."
    html_body = f"""
    <html><body>
      <p>{body_intro}</p>
      <p><b>Rango analizado:</b> {ini} – {fin}</p>
      <p>Se adjuntan archivos HTML y CSV.</p>
    </body></html>
    """

    method = (email_cfg.get("method") or "outlook").lower()
    paths = [Path(p).resolve() for p in attachments if Path(p).exists()]

    if method == "smtp":
        return _send_with_smtp(email_cfg, to, cc, subject, html_body, paths)
    else:
        return _send_with_outlook(to, cc, subject, html_body, paths)
