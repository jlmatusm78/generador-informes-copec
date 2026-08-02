from __future__ import annotations

import io
import re
import smtplib
import ssl
import unicodedata
import zipfile
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable

import pandas as pd


RECIPIENT_COLUMNS = ["Transportista", "Para", "CC", "Activo"]


@dataclass(frozen=True)
class SmtpSettings:
    host: str
    port: int
    username: str
    password: str
    sender_email: str
    sender_name: str = "Torre de Control COPEC"
    use_tls: bool = True
    use_ssl: bool = False


@dataclass
class PreparedEmail:
    transportista: str
    to: list[str]
    cc: list[str]
    subject: str
    body: str
    attachment_name: str
    attachment_bytes: bytes


def normalize_text(value: object) -> str:
    text = str(value or "").upper().strip()
    text = "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", text)


def split_addresses(value: object) -> list[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    pieces = re.split(r"[;,\n]+", str(value))
    return [p.strip() for p in pieces if p.strip()]


def is_valid_email(address: str) -> bool:
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", address.strip()))


def load_recipient_catalog(uploaded_file) -> pd.DataFrame:
    """Load a recipients CSV/XLSX with Transportista, Para, CC and Activo columns."""
    if uploaded_file is None:
        return pd.DataFrame(columns=RECIPIENT_COLUMNS)

    name = getattr(uploaded_file, "name", "destinatarios.xlsx").lower()
    raw = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else Path(uploaded_file).read_bytes()
    if name.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(raw), sep=None, engine="python")
    else:
        excel = pd.ExcelFile(io.BytesIO(raw))
        preferred = next((s for s in excel.sheet_names if normalize_text(s) in {"CORREO TRANSPORTISTAS", "DESTINATARIOS"}), excel.sheet_names[0])
        df = pd.read_excel(io.BytesIO(raw), sheet_name=preferred)

    rename_map = {}
    for col in df.columns:
        key = normalize_text(col)
        if key in {"TRANSPORTISTA", "EMPRESA"}:
            rename_map[col] = "Transportista"
        elif key in {"PARA", "CORREO", "EMAIL", "DESTINATARIO"}:
            rename_map[col] = "Para"
        elif key in {"CC", "COPIA"}:
            rename_map[col] = "CC"
        elif key in {"ACTIVO", "ENVIAR", "HABILITADO"}:
            rename_map[col] = "Activo"
    df = df.rename(columns=rename_map)
    for col in RECIPIENT_COLUMNS:
        if col not in df.columns:
            df[col] = "SI" if col == "Activo" else ""
    df = df[RECIPIENT_COLUMNS].copy()
    df["Transportista"] = df["Transportista"].fillna("").astype(str).str.strip()
    df["Para"] = df["Para"].fillna("").astype(str).str.strip()
    df["CC"] = df["CC"].fillna("").astype(str).str.strip()
    df["Activo"] = df["Activo"].fillna("SI").astype(str).str.upper().str.strip()
    return df[df["Transportista"] != ""].drop_duplicates(subset=["Transportista"], keep="last")


def read_pdf_attachments(zip_bytes: bytes) -> list[dict]:
    attachments: list[dict] = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        for name in sorted(zf.namelist()):
            if name.lower().endswith(".pdf"):
                attachments.append({
                    "filename": Path(name).name,
                    "bytes": zf.read(name),
                    "transportista": transportista_from_filename(Path(name).name),
                })
    return attachments


def transportista_from_filename(filename: str) -> str:
    stem = Path(filename).stem
    if "GLOBAL_COPEC" in normalize_text(stem).replace(" ", "_"):
        return "COPEC"
    # Formats: 01_Informe_EMPRESA_0607_12072026 or 01_Informe_Semanal_EMPRESA_...
    cleaned = re.sub(r"^\d+_Informe(?:_(?:Semanal|Mensual))?_", "", stem, flags=re.IGNORECASE)
    cleaned = re.sub(r"_\d{4}_\d{8}$", "", cleaned)
    cleaned = re.sub(r"_\d{4}_\d{4,8}$", "", cleaned)
    return cleaned.replace("_", " ").strip()


def build_dispatch_table(attachments: list[dict], catalog: pd.DataFrame) -> pd.DataFrame:
    lookup = {
        normalize_text(row.Transportista): row
        for row in catalog.itertuples(index=False)
    }
    rows = []
    for item in attachments:
        transportista = item["transportista"]
        match = lookup.get(normalize_text(transportista))
        para = getattr(match, "Para", "") if match is not None else ""
        cc = getattr(match, "CC", "") if match is not None else ""
        active = getattr(match, "Activo", "SI") if match is not None else "SI"
        rows.append({
            "Enviar": str(active).upper() not in {"NO", "FALSE", "0"},
            "Transportista": transportista,
            "Para": para,
            "CC": cc,
            "Informe": item["filename"],
            "Estado": "Listo" if para else "Falta destinatario",
        })
    return pd.DataFrame(rows)


def make_subject(transportista: str, start_label: str, end_label: str, mode: str = "semanal") -> str:
    if normalize_text(transportista) == "COPEC":
        return f"Informe {mode} global COPEC | {start_label} al {end_label}"
    return f"Informe {mode} de alertas | {transportista} | {start_label} al {end_label}"


def make_body(transportista: str, start_label: str, end_label: str, mode: str = "semanal") -> str:
    if normalize_text(transportista) == "COPEC":
        return (
            "Estimados:\n\n"
            f"Junto con saludar, se adjunta el Informe {mode.capitalize()} Global COPEC correspondiente al período "
            f"comprendido entre el {start_label} y el {end_label}.\n\n"
            f"El informe contiene la evolución {mode} de alertas, empresas y conductores con mayor recurrencia, "
            "gestión de fatiga, cumplimiento del protocolo y principales hallazgos operacionales.\n\n"
            "Saludos cordiales,\nTorre de Control COPEC"
        )
    return (
        "Estimados:\n\n"
        f"Junto con saludar, se adjunta el informe {mode} de alertas de {transportista}, correspondiente al período "
        f"comprendido entre el {start_label} y el {end_label}.\n\n"
        f"El documento contiene la evolución de las alertas, comparación con períodos anteriores, conductores con mayor "
        "recurrencia y estado de cumplimiento del protocolo asociado a eventos de fatiga.\n\n"
        "Solicitamos revisar los hallazgos y gestionar las desviaciones identificadas.\n\n"
        "Saludos cordiales,\nTorre de Control COPEC"
    )


def prepare_email(
    dispatch_row: dict,
    attachment: dict,
    start_label: str,
    end_label: str,
    test_recipient: str | None = None,
) -> PreparedEmail:
    transportista = str(dispatch_row["Transportista"])
    to = [test_recipient] if test_recipient else split_addresses(dispatch_row.get("Para", ""))
    cc = [] if test_recipient else split_addresses(dispatch_row.get("CC", ""))
    invalid = [a for a in to + cc if not is_valid_email(a)]
    if invalid:
        raise ValueError(f"Correos inválidos para {transportista}: {', '.join(invalid)}")
    if not to:
        raise ValueError(f"No hay destinatario principal para {transportista}.")
    mode = "mensual" if "MENSUAL" in normalize_text(attachment.get("filename", "")) else "semanal"
    subject = make_subject(transportista, start_label, end_label, mode)
    if test_recipient:
        subject = f"[PRUEBA] {subject}"
    return PreparedEmail(
        transportista=transportista,
        to=to,
        cc=cc,
        subject=subject,
        body=make_body(transportista, start_label, end_label, mode),
        attachment_name=attachment["filename"],
        attachment_bytes=attachment["bytes"],
    )


def _build_message(settings: SmtpSettings, email: PreparedEmail) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = f"{settings.sender_name} <{settings.sender_email}>"
    msg["To"] = ", ".join(email.to)
    if email.cc:
        msg["Cc"] = ", ".join(email.cc)
    msg["Subject"] = email.subject
    msg.set_content(email.body)
    msg.add_attachment(
        email.attachment_bytes,
        maintype="application",
        subtype="pdf",
        filename=email.attachment_name,
    )
    return msg


def send_email(settings: SmtpSettings, email: PreparedEmail) -> None:
    msg = _build_message(settings, email)
    recipients = email.to + email.cc
    if settings.use_ssl:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(settings.host, settings.port, context=context, timeout=45) as smtp:
            smtp.login(settings.username, settings.password)
            smtp.send_message(msg, to_addrs=recipients)
    else:
        with smtplib.SMTP(settings.host, settings.port, timeout=45) as smtp:
            smtp.ehlo()
            if settings.use_tls:
                smtp.starttls(context=ssl.create_default_context())
                smtp.ehlo()
            smtp.login(settings.username, settings.password)
            smtp.send_message(msg, to_addrs=recipients)


def send_batch(settings: SmtpSettings, emails: Iterable[PreparedEmail]) -> list[dict]:
    results = []
    for email in emails:
        try:
            send_email(settings, email)
            results.append({
                "Transportista": email.transportista,
                "Destinatario": "; ".join(email.to),
                "Informe": email.attachment_name,
                "Resultado": "Enviado",
                "Detalle": "",
            })
        except Exception as exc:
            results.append({
                "Transportista": email.transportista,
                "Destinatario": "; ".join(email.to),
                "Informe": email.attachment_name,
                "Resultado": "Error",
                "Detalle": str(exc),
            })
    return results
