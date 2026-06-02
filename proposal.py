"""
ACS Commercial Cleaning - proposal builder (web / server version).

Renders the branded Word template with docxtpl and converts to PDF using
LibreOffice (installed by the host via packages.txt). No Microsoft Word
needed, so this runs on a Linux web server.

Supports modular proposals: the four core sections (Cover Letter, Scope of
Work, Investment Options, Service Terms) plus any number of optional sections
(Executive Summary, What We Understood, etc.) placed at three anchor points.
Scope can be entered by area or by rotation frequency.
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

PLACEMENTS = ("after_cover", "after_scope", "after_investment")


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


def _process_section(raw):
    """Shape one optional section, or None if empty."""
    title = (raw.get("title") or "").strip()
    paragraphs = _clean_list(raw.get("paragraphs"))
    bullets = _clean_list(raw.get("bullets"))
    if not (title or paragraphs or bullets):
        return None
    return {
        "heading": title.upper(),        # green bar text
        "toc_title": title or "Section",  # table of contents text
        "paragraphs": paragraphs,
        "bullets": bullets,
    }


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

    # --- inclusions (paired into a 2-column grid) ---
    inclusions = _clean_list(data.get("inclusions"))
    inclusion_rows = []
    for i in range(0, len(inclusions), 2):
        left = inclusions[i]
        right = inclusions[i + 1] if i + 1 < len(inclusions) else ""
        inclusion_rows.append({
            "left": left, "right": right,
            "right_mark": "✓ " if right else "",
        })

    # --- scope rows (same shape for 'area' and 'rotation' styles) ---
    areas = []
    for area in (data.get("areas") or []):
        name = (area.get("name") or "").strip()
        lines = _clean_list(area.get("items"))
        if name or lines:
            areas.append({"name": name, "lines": lines})
    scope_style = (data.get("scope_style") or "area").strip().lower()
    scope_col = "FREQUENCY" if scope_style == "rotation" else "AREA"

    # split inclusions into two columns (column-major, to match the C2 grid)
    half = (len(inclusions) + 1) // 2
    inclusions_left = inclusions[:half]
    inclusions_right = inclusions[half:]

    # --- services / inclusions / terms ---
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

    # --- optional sections, grouped by placement ---
    buckets = {p: [] for p in PLACEMENTS}
    for raw in (data.get("extra_sections") or []):
        sec = _process_section(raw)
        if not sec:
            continue
        placement = (raw.get("placement") or "after_cover").strip().lower()
        if placement not in buckets:
            placement = "after_cover"
        buckets[placement].append(sec)

    # --- assign section numbers in document order + build the TOC ---
    toc = []
    counter = {"n": 0}

    def number(title):
        counter["n"] += 1
        toc.append({"num": counter["n"], "title": title})
        return counter["n"]

    num_cover = number("Cover Letter")
    for sec in buckets["after_cover"]:
        sec["number"] = number(sec["toc_title"])
    num_scope = number("Scope of Work")
    for sec in buckets["after_scope"]:
        sec["number"] = number(sec["toc_title"])
    num_investment = number("Investment Options")
    for sec in buckets["after_investment"]:
        sec["number"] = number(sec["toc_title"])
    num_terms = number("Service Terms")

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
        "scope_col": scope_col,
        "areas": areas,
        "service_notes": _clean_list(data.get("service_notes")),
        "services": services,
        "inclusion_rows": inclusion_rows,
        "inclusions_left": inclusions_left,
        "inclusions_right": inclusions_right,
        "terms": terms,
        "toc": toc,
        "num_cover": num_cover,
        "num_scope": num_scope,
        "num_investment": num_investment,
        "num_terms": num_terms,
        "extra_after_cover": buckets["after_cover"],
        "extra_after_scope": buckets["after_scope"],
        "extra_after_investment": buckets["after_investment"],
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
