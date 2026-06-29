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
import threading
import time
from datetime import date
from io import BytesIO

import jinja2
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

# Only one LibreOffice conversion runs at a time (process-wide) so concurrent
# users can't spawn several soffice processes and exhaust the container's RAM.
_CONVERT_LOCK = threading.Lock()
_TEMP_PREFIX = "acs_pdf_"

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "template", "proposal_template.docx")
HIA_LOGO_PATH = os.path.join(HERE, "template", "assets", "hia_logo.jpg")
COMPLIANCE_DIR = os.path.join(HERE, "template", "assets", "compliance")
# Each certificate is embedded as a full-page image inside the document (so it
# inherits the header/footer and page numbers). (context var, bundled image).
COMPLIANCE_CERTS = (
    ("cert_labour", "labour_hire.png"),
    ("cert_workcover", "workcover.png"),
    ("cert_public", "public_liability.png"),
)

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
        # testimonial sections render their paragraphs as quote boxes
        "kind": "quote" if "testimonial" in title.lower() else "normal",
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
        pad = "%02d" % counter["n"]
        toc.append({"num": pad, "title": title})
        return pad

    num_cover = number("Cover Letter")
    for sec in buckets["after_cover"]:
        sec["number"] = number(sec["toc_title"])
    num_scope = number("Scope of Work")
    for sec in buckets["after_scope"]:
        sec["number"] = number(sec["toc_title"])
    num_investment = number("Investment Options")
    for sec in buckets["after_investment"]:
        sec["number"] = number(sec["toc_title"])
    include_testimonial = bool(data.get("include_testimonial"))
    num_testimonial = number("Testimonial") if include_testimonial else ""
    include_compliance = bool(data.get("include_compliance"))
    num_compliance = number("Compliance Documents") if include_compliance else ""
    num_terms = number("Service Terms")
    num_acceptance = number("Acceptance")

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
        "num_acceptance": num_acceptance,
        "include_testimonial": include_testimonial,
        "num_testimonial": num_testimonial,
        "include_compliance": include_compliance,
        "num_compliance": num_compliance,
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
    # The HIA testimonial logo is a real embedded image (not a screenshot),
    # only attached when the testimonial section is switched on.
    if context.get("include_testimonial") and os.path.exists(HIA_LOGO_PATH):
        context["hia_logo"] = InlineImage(tpl, HIA_LOGO_PATH, height=Mm(22))
    else:
        context["hia_logo"] = ""
    # Compliance certificates: full-page embedded images (160mm wide fits one
    # US-Letter page with margin to spare) so they carry the proposal's header,
    # footer and page numbers.
    if context.get("include_compliance"):
        for var, fn in COMPLIANCE_CERTS:
            path = os.path.join(COMPLIANCE_DIR, fn)
            context[var] = InlineImage(tpl, path, width=Mm(160)) if os.path.exists(path) else ""
    else:
        for var, _fn in COMPLIANCE_CERTS:
            context[var] = ""
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


def sweep_stale_temp(max_age_seconds=3600):
    """Remove orphaned conversion temp dirs (e.g. left behind if soffice was
    OOM-killed). Safe to call any time; never raises."""
    base = tempfile.gettempdir()
    try:
        now = time.time()
        for name in os.listdir(base):
            if not name.startswith(_TEMP_PREFIX):
                continue
            path = os.path.join(base, name)
            try:
                if now - os.path.getmtime(path) > max_age_seconds:
                    shutil.rmtree(path, ignore_errors=True)
            except OSError:
                pass
    except OSError:
        pass


# Sweep once when the module is first imported (i.e. on app start).
sweep_stale_temp()


def docx_bytes_to_pdf_bytes(docx_bytes):
    """Convert .docx bytes to .pdf bytes via LibreOffice. Returns None if
    LibreOffice is unavailable or conversion fails."""
    soffice = _find_soffice()
    if not soffice:
        return None

    sweep_stale_temp()  # opportunistically clear any orphaned temp dirs
    tmp = tempfile.mkdtemp(prefix=_TEMP_PREFIX)
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
        # Serialize conversions and hard-kill a hung/over-time soffice so it
        # can't linger consuming RAM.
        with _CONVERT_LOCK:
            try:
                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                        stderr=subprocess.PIPE)
            except Exception:  # noqa: BLE001
                return None
            try:
                proc.communicate(timeout=120)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.communicate()
                return None
            except Exception:  # noqa: BLE001
                try:
                    proc.kill()
                except Exception:  # noqa: BLE001
                    pass
                return None

        pdf_path = os.path.join(tmp, "proposal.pdf")
        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as fh:
                return fh.read()
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
