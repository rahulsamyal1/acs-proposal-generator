"""
ACS Commercial Cleaning - proposal builder (web / server version).

Renders the branded Word template with docxtpl and converts to PDF using
LibreOffice (which the host installs via packages.txt). No Microsoft Word
needed, so this runs on a Linux web server.
"""

import os
import re
import shutil
import subprocess
import tempfile
from datetime import date
from io import BytesIO

import jinja2
from docxtpl import DocxTemplate

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "template", "proposal_template.docx")


# ---------------------------------------------------------------------------
# Input shaping
# ---------------------------------------------------------------------------
def _clean_list(value):
    if value is None:
        return []
    if isinstance(value, str):
        parts = value.replace("\r\n", "\n").split("\n")
    else:
        parts = list(value)
    return [str(p).strip() for p in parts if str(p).strip()]


def safe_filename(text):
    text = re.sub(r"[^A-Za-z0-9 _-]", "", str(text)).strip()
    text = re.sub(r"\s+", "_", text)
    return text or "Proposal"


def _initials(name):
    words = re.findall(r"[A-Za-z0-9]+", str(name))
    if not words:
        return "XXX"
    return "".join(w[0] for w in words).upper()[:4]


def build_context(data):
    client_name = (data.get("client_name") or "").strip() or "[Client Name]"
    contact_name = (data.get("contact_name") or "").strip()

    contact_first_name = (data.get("contact_first_name") or "").strip()
    if not contact_first_name and contact_name:
        contact_first_name = contact_name.split()[0]

    today = date.today()
    date_str = (data.get("date") or "").strip() or today.strftime("%d %B %Y")

    reference = (data.get("reference") or "").strip()
    if not reference:
        reference = "ACS-{}-{}-001".format(_initials(client_name), today.year)

    inclusions = _clean_list(data.get("inclusions"))
    inclusion_rows = []
    for i in range(0, len(inclusions), 2):
        left = inclusions[i]
        right = inclusions[i + 1] if i + 1 < len(inclusions) else ""
        inclusion_rows.append({
            "left": left, "right": right,
            "right_mark": "✓ " if right else "",
        })

    areas = []
    for area in (data.get("areas") or []):
        name = (area.get("name") or "").strip()
        lines = _clean_list(area.get("items"))
        if name or lines:
            areas.append({"name": name, "lines": lines})

    services = []
    for svc in (data.get("services") or []):
        if not any((svc.get(k) or "").strip()
                   for k in ("name", "description", "schedule",
                             "price_ex", "price_period", "price_inc")):
            continue
        services.append({
            "name": (svc.get("name") or "").strip(),
            "description": (svc.get("description") or "").strip(),
            "schedule": (svc.get("schedule") or "").strip(),
            "price_ex": (svc.get("price_ex") or "").strip(),
            "price_period": (svc.get("price_period") or "").strip(),
            "price_inc": (svc.get("price_inc") or "").strip(),
        })

    terms = []
    for term in (data.get("terms") or []):
        label = (term.get("label") or "").strip()
        value = (term.get("value") or "").strip()
        if label or value:
            terms.append({"label": label, "value": value})

    return {
        "client_name": client_name,
        "contact_name": contact_name or "[Contact Name]",
        "contact_first_name": contact_first_name or "there",
        "contact_title": (data.get("contact_title") or "").strip(),
        "client_address": (data.get("client_address") or "").strip(),
        "contact_phone": (data.get("contact_phone") or "").strip(),
        "date": date_str,
        "reference": reference,
        "cover_paragraphs": _clean_list(data.get("cover_paragraphs")),
        "frequency": (data.get("frequency") or "").strip(),
        "duration": (data.get("duration") or "").strip(),
        "areas": areas,
        "service_notes": _clean_list(data.get("service_notes")),
        "services": services,
        "inclusion_rows": inclusion_rows,
        "terms": terms,
    }


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------
def render_docx_bytes(data):
    """Render the template and return the .docx as bytes."""
    context = build_context(data)
    tpl = DocxTemplate(TEMPLATE_PATH)
    jinja_env = jinja2.Environment(autoescape=True)
    tpl.render(context, jinja_env)
    buf = BytesIO()
    tpl.save(buf)
    return buf.getvalue()


def _find_soffice():
    for name in ("soffice", "libreoffice", "soffice.bin"):
        path = shutil.which(name)
        if path:
            return path
    # common Windows install location (for local testing)
    win = r"C:\Program Files\LibreOffice\program\soffice.exe"
    return win if os.path.exists(win) else None


def docx_bytes_to_pdf_bytes(docx_bytes):
    """Convert .docx bytes to .pdf bytes via LibreOffice. Returns None if
    LibreOffice is unavailable or conversion fails."""
    soffice = _find_soffice()
    if not soffice:
        return None

    tmp = tempfile.mkdtemp(prefix="acs_pdf_")
    try:
        docx_path = os.path.join(tmp, "proposal.docx")
        with open(docx_path, "wb") as fh:
            fh.write(docx_bytes)

        profile = os.path.join(tmp, "lo_profile")
        cmd = [
            soffice, "--headless", "--nologo", "--nofirststartwizard",
            "-env:UserInstallation=file://" + profile.replace("\\", "/"),
            "--convert-to", "pdf", "--outdir", tmp, docx_path,
        ]
        try:
            subprocess.run(cmd, capture_output=True, timeout=120, check=False)
        except Exception:  # noqa: BLE001
            return None

        pdf_path = os.path.join(tmp, "proposal.pdf")
        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as fh:
                return fh.read()
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
