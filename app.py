"""
ACS Commercial Cleaning - Proposal Generator (web app).

A single-password Streamlit app. Fill the form, click Generate, download the
branded Word doc and PDF. Deploy free on Streamlit Community Cloud.

Modular: the four core sections plus optional sections (Executive Summary,
What We Understood, etc.) placed at chosen anchor points; scope by area or by
rotation frequency.
"""

import hashlib
import json
import os

import pandas as pd
import streamlit as st

import proposal

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "assets", "acs_logo.png")
PRESETS_PATH = os.path.join(HERE, "presets.json")

st.set_page_config(page_title="ACS Proposal Generator", page_icon="🧽", layout="centered")

PLACEMENT_LABELS = {
    "After Cover Letter": "after_cover",
    "After Scope of Work": "after_scope",
    "After Investment Options": "after_investment",
}

# Quick-add templates for the common optional sections (titles + starter text).
SECTION_PRESETS = {
    "Executive Summary": {
        "placement": "After Cover Letter",
        "paragraphs": "From our walkthrough and discussions, we understand the current cleaning has become inconsistent and the standard has slipped over time. Just as importantly, you want a provider who listens, is easy to deal with, and actually responds when something needs attention.\nThis proposal is built around two priorities: reset the site to a clean, controlled standard, then maintain it consistently — a reliable, well-managed service where you are not chasing missed work or dealing with a different cleaner every week. We have listened to what matters most to you, and this proposal is built around exactly that.",
        "bullets": "A consistent, reliable standard on every visit\nClear communication and a quick response when issues are raised\nThe same trusted cleaner who gets to know your site\nA service that is genuinely easy for your team to manage",
    },
    "Our Cleaning Approach": {
        "placement": "After Cover Letter",
        "paragraphs": "Not every task needs to be done daily to keep a good standard. We separate daily cleaning priorities from rotating detail tasks.\nThe goal is not to reduce the standard - it is to focus daily cleaning time where it matters most.",
        "bullets": "",
    },
    "Transition & Site Improvement Plan": {
        "placement": "After Scope of Work",
        "paragraphs": "When ACS takes over a site, the first few weeks establish the routine, the cleaners learn the site, and any shortfalls are identified.",
        "bullets": "Confirm the final agreed scope\nComplete a site handover with the cleaning team\nPrioritise bathrooms, kitchens and high-touch areas\nStabilise the regular routine, then work through detail areas",
    },
    "How ACS Manages the Service": {
        "placement": "After Scope of Work",
        "paragraphs": "ACS focuses on stability, consistency and accountability. Wherever possible we use the same cleaners so they become familiar with the building and expectations.",
        "bullets": "Consistent cleaners wherever possible\nClear scope and proper site induction\nSupervisor involvement where required\nDirect communication and follow-up on missed items",
    },
    "Relevant Experience & References": {
        "placement": "After Investment Options",
        "paragraphs": "ACS Commercial Cleaning has experience across commercial offices, industrial sites, food production and long-term recurring contracts.",
        "bullets": "References available on request\nThe Salvation Army\nSt Pauls\nTurosi Food Solutions",
    },
    "Client Testimonials": {
        "placement": "After Investment Options",
        "paragraphs": "“[Paste a short client testimonial here — what they were struggling with, and the difference ACS made.]”\n“[Add a second client testimonial here if you have one.]”",
        "bullets": "",
    },
    "Client References": {
        "placement": "After Investment Options",
        "paragraphs": "We are happy to provide direct references from clients who have worked with us over time. We can arrange a quick phone call with any of the following.",
        "bullets": "The Salvation Army\nSt Pauls\nTurosi Food Solutions\n[Add another relevant reference]",
    },
    "First Clean Guarantee": {
        "placement": "After Investment Options",
        "paragraphs": "We back our work from day one. If anything in the first clean is not up to standard, simply let us know within 48 hours and we will return and put it right at no extra cost.\nThe goal is for you to see the difference from the very first visit.",
        "bullets": "We return and re-clean any missed areas at no charge\nA supervisor checks the quality of the first clean\nClear communication if anything needs your input",
    },
}


@st.cache_data
def load_presets():
    with open(PRESETS_PATH, "r", encoding="utf-8") as fh:
        return json.load(fh)


# ---------------------------------------------------------------------------
# Password gate
# ---------------------------------------------------------------------------
def password_ok():
    if st.session_state.get("auth_ok"):
        return True
    try:
        expected = st.secrets["app_password"]
    except Exception:  # noqa: BLE001
        expected = "changeme"
    if os.path.exists(LOGO):
        st.image(LOGO, width=260)
    st.subheader("ACS Proposal Generator")
    pw = st.text_input("Password", type="password")
    if pw:
        if pw == expected:
            st.session_state.auth_ok = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()


password_ok()

PRESETS = load_presets()
presets = PRESETS.get("presets", [])
default_terms = PRESETS.get("default_terms", [])


# ---------------------------------------------------------------------------
# State helpers
# ---------------------------------------------------------------------------
def _ensure(key, default):
    if key not in st.session_state:
        st.session_state[key] = default


def new_id():
    st.session_state.next_id = st.session_state.get("next_id", 0) + 1
    return st.session_state.next_id


def add_area(name="", items=""):
    i = new_id()
    st.session_state[f"area_name_{i}"] = name
    st.session_state[f"area_items_{i}"] = items
    st.session_state.area_ids.append(i)


def add_service(svc=None):
    svc = svc or {}
    i = new_id()
    st.session_state[f"svc_name_{i}"] = svc.get("name", "")
    st.session_state[f"svc_desc_{i}"] = svc.get("description", "")
    st.session_state[f"svc_sched_{i}"] = svc.get("schedule", "")
    try:
        amt = float(svc.get("amount_ex") or 0)
    except (TypeError, ValueError):
        amt = 0.0
    st.session_state[f"svc_amt_{i}"] = amt
    st.session_state[f"svc_period_{i}"] = svc.get("price_period", "per month")
    st.session_state.service_ids.append(i)


def add_section(title="", placement="After Cover Letter", paragraphs="", bullets=""):
    i = new_id()
    st.session_state[f"sec_title_{i}"] = title
    st.session_state[f"sec_place_{i}"] = placement
    st.session_state[f"sec_paras_{i}"] = paragraphs
    st.session_state[f"sec_bullets_{i}"] = bullets
    st.session_state.section_ids.append(i)


def apply_preset(p):
    st.session_state["frequency"] = p.get("frequency", "")
    st.session_state["duration"] = p.get("duration", "")
    st.session_state["service_notes"] = "\n".join(p.get("service_notes", []))
    st.session_state["inclusions"] = "\n".join(p.get("inclusions", []))
    st.session_state.area_ids = []
    for a in p.get("areas", []):
        add_area(a.get("name", ""), "\n".join(a.get("items", [])))
    st.session_state.service_ids = []
    for s in p.get("services", []):
        add_service(s)


# Placeholder cover letter shown on first load (edit the [bracketed] bits).
DEFAULT_COVER = "\n".join([
    "Thank you for the opportunity to inspect [your site] and submit this proposal for cleaning services at [Client name].",
    "During the walkthrough we noted [what you observed - the current standard, anything being missed, and the main concern raised].",
    "We understand you are looking for a reliable, easy-to-manage service that keeps a consistent standard each visit.",
    "Our proposal is based on [X days per week, after hours], covering the main areas, kitchens, bathrooms and general presentation.",
])

CLIENT_PLACEHOLDERS = {
    "client_name": "[Client name]",
    "site_office": "[your site]",
    "contact_name": "[Contact name]",
    "contact_title": "[Title]",
    "contact_phone": "[Phone]",
    "client_address": "[Address]",
    "date": "",
    "reference": "",
}


def seed_defaults():
    """Pre-fill the whole form with placeholder/example content on first load."""
    for key, val in CLIENT_PLACEHOLDERS.items():
        st.session_state[key] = val
    st.session_state["cover_paragraphs"] = DEFAULT_COVER
    if presets:
        apply_preset(presets[0])
    else:
        add_area()
        add_service()


def clear_all():
    """Empty the form to start from scratch."""
    for key in list(CLIENT_PLACEHOLDERS) + ["cover_paragraphs", "frequency",
                                            "duration", "service_notes", "inclusions"]:
        st.session_state[key] = ""
    st.session_state.area_ids = []
    st.session_state.service_ids = []
    st.session_state.section_ids = []
    add_area()
    add_service()


GST_RATE = 1.10


def _money(n):
    try:
        n = float(n)
    except (TypeError, ValueError):
        return ""
    s = "{:,.2f}".format(n)
    if s.endswith(".00"):
        s = s[:-3]
    return "$" + s


def _service_for_render(i):
    """Build a service row for the document, computing GST from the ex-GST amount."""
    ss = st.session_state
    try:
        amt = float(ss.get(f"svc_amt_{i}", 0) or 0)
    except (TypeError, ValueError):
        amt = 0.0
    if amt > 0:
        price_ex = _money(amt) + " plus GST"
        price_inc = _money(round(amt * GST_RATE, 2)) + " incl GST"
    else:
        price_ex = price_inc = ""
    return {
        "name": ss.get(f"svc_name_{i}", ""),
        "description": ss.get(f"svc_desc_{i}", ""),
        "schedule": ss.get(f"svc_sched_{i}", ""),
        "price_ex": price_ex,
        "price_period": ss.get(f"svc_period_{i}", ""),
        "price_inc": price_inc,
    }


def collect_form(terms_records):
    """Snapshot the whole form (raw values) — used for drafts and for generating."""
    ss = st.session_state
    return {
        "client_name": ss.get("client_name", ""),
        "site_office": ss.get("site_office", ""),
        "contact_name": ss.get("contact_name", ""),
        "contact_title": ss.get("contact_title", ""),
        "contact_phone": ss.get("contact_phone", ""),
        "client_address": ss.get("client_address", ""),
        "date": ss.get("date", ""),
        "reference": ss.get("reference", ""),
        "cover_paragraphs": ss.get("cover_paragraphs", ""),
        "frequency": ss.get("frequency", ""),
        "areas": [{"name": ss.get(f"area_name_{i}", ""), "items": ss.get(f"area_items_{i}", "")}
                  for i in ss.area_ids],
        "service_notes": ss.get("service_notes", ""),
        "services": [{"name": ss.get(f"svc_name_{i}", ""), "description": ss.get(f"svc_desc_{i}", ""),
                      "schedule": ss.get(f"svc_sched_{i}", ""), "amount_ex": ss.get(f"svc_amt_{i}", 0),
                      "price_period": ss.get(f"svc_period_{i}", "")} for i in ss.service_ids],
        "inclusions": ss.get("inclusions", ""),
        "terms": terms_records,
        "extra_sections": [{"title": ss.get(f"sec_title_{i}", ""), "placement": ss.get(f"sec_place_{i}", "After Cover Letter"),
                            "paragraphs": ss.get(f"sec_paras_{i}", ""), "bullets": ss.get(f"sec_bullets_{i}", "")}
                           for i in ss.section_ids],
    }


def restore_form(form):
    """Rebuild the form from a saved draft dict."""
    ss = st.session_state
    for k in CLIENT_PLACEHOLDERS:
        ss[k] = form.get(k, "")
    ss["cover_paragraphs"] = form.get("cover_paragraphs", "")
    ss["frequency"] = form.get("frequency", "")
    ss["service_notes"] = form.get("service_notes", "")
    ss["inclusions"] = form.get("inclusions", "")
    ss.area_ids = []
    for a in form.get("areas", []):
        add_area(a.get("name", ""), a.get("items", ""))
    if not ss.area_ids:
        add_area()
    ss.service_ids = []
    for s in form.get("services", []):
        add_service(s)
    if not ss.service_ids:
        add_service()
    ss.section_ids = []
    for s in form.get("extra_sections", []):
        add_section(s.get("title", ""), s.get("placement", "After Cover Letter"),
                    s.get("paragraphs", ""), s.get("bullets", ""))
    terms = form.get("terms", [])
    if terms:
        ss.terms_df = pd.DataFrame([{"Term": t.get("label", ""), "Detail": t.get("value", "")}
                                    for t in terms])
        ss.pop("terms_editor", None)


def _form_sig(terms_records):
    """A fingerprint of the whole form, so we can tell if anything changed
    since the last time the proposal was generated."""
    blob = json.dumps(collect_form(terms_records), sort_keys=True, default=str)
    return hashlib.md5(blob.encode("utf-8")).hexdigest()


if "area_ids" not in st.session_state:
    st.session_state.area_ids = []
    st.session_state.service_ids = []
    st.session_state.section_ids = []
    seed_defaults()


# ---------------------------------------------------------------------------
# Header + preset picker
# ---------------------------------------------------------------------------
if os.path.exists(LOGO):
    st.image(LOGO, width=260)
st.title("Proposal Generator")
st.caption("Fill in the details, click Generate, and download the branded proposal.")

with st.container(border=True):
    names = ["— Blank —"] + [p["name"] for p in presets]
    choice = st.selectbox("Start from a job type (optional)", names, index=0)
    pcol1, pcol2 = st.columns(2)
    with pcol1:
        if st.button("Apply preset", disabled=(choice == "— Blank —"), width="stretch"):
            apply_preset(presets[names.index(choice) - 1])
            st.rerun()
    with pcol2:
        if st.button("Start blank", width="stretch"):
            clear_all()
            st.rerun()
    st.caption("The form opens pre-filled with placeholder text — edit the "
               "[bracketed] parts. Use “Start blank” to clear everything.")

with st.container(border=True):
    st.markdown("**💾 Drafts** — your work is **not** auto-saved. Save a draft before "
                "refreshing or closing, then load it here to continue.")
    _up = st.file_uploader("Load a saved draft (.json)", type=["json"], key="draft_upload")
    if st.button("Load this draft", disabled=(_up is None)):
        try:
            restore_form(json.loads(_up.getvalue().decode("utf-8")))
            st.session_state["_draft_loaded"] = True
            st.rerun()
        except Exception as exc:  # noqa: BLE001
            st.error("Could not read that draft file: %s" % exc)
if st.session_state.pop("_draft_loaded", False):
    st.success("Draft loaded.")

# 1. Client details ----------------------------------------------------------
st.header("1 · Client details")
c1, c2 = st.columns(2)
with c1:
    _ensure("client_name", "")
    st.text_input("Client / Company name", key="client_name")
    _ensure("contact_name", "")
    st.text_input("Contact name", key="contact_name")
    _ensure("contact_phone", "")
    st.text_input("Contact phone", key="contact_phone")
    _ensure("date", "")
    st.text_input("Date (blank = today)", key="date")
with c2:
    _ensure("site_office", "")
    st.text_input("Site / office (for the letter)", key="site_office")
    _ensure("contact_title", "")
    st.text_input("Contact title", key="contact_title")
    _ensure("client_address", "")
    st.text_input("Client address", key="client_address")
    _ensure("reference", "")
    st.text_input("Reference (blank = auto)", key="reference")

# 2. Cover letter ------------------------------------------------------------
st.header("2 · Cover letter")
_ensure("cover_paragraphs", "")
st.text_area("Letter paragraphs (one paragraph per line)", key="cover_paragraphs", height=170)

# 3. Scope of work -----------------------------------------------------------
st.header("3 · Scope of work")
_ensure("frequency", "")
st.text_input("Cleaning frequency", key="frequency",
              placeholder="e.g. 2 days per week, after hours")

st.markdown("**Areas**")
for i in list(st.session_state.area_ids):
    with st.container(border=True):
        cc1, cc2 = st.columns([3, 1])
        with cc1:
            st.text_input("Area name", key=f"area_name_{i}")
        with cc2:
            st.write("")
            st.write("")
            if st.button("Remove", key=f"rm_area_{i}"):
                st.session_state.area_ids.remove(i)
                st.rerun()
        st.text_area("Scope items (one per line)", key=f"area_items_{i}", height=110)
if st.button("➕ Add row"):
    add_area()
    st.rerun()

_ensure("service_notes", "")
st.text_area("Service notes (one per line)", key="service_notes", height=110)

# 4. Investment options ------------------------------------------------------
st.header("4 · Investment options")
for i in list(st.session_state.service_ids):
    with st.container(border=True):
        cc1, cc2 = st.columns([3, 1])
        with cc1:
            st.text_input("Service name", key=f"svc_name_{i}")
        with cc2:
            st.write("")
            st.write("")
            if st.button("Remove", key=f"rm_svc_{i}"):
                st.session_state.service_ids.remove(i)
                st.rerun()
        st.text_area("Description", key=f"svc_desc_{i}", height=80)
        d1, d2 = st.columns(2)
        with d1:
            st.text_input("Schedule", key=f"svc_sched_{i}")
            st.text_input("Period", key=f"svc_period_{i}")
        with d2:
            amt = st.number_input("Monthly amount ex GST ($)", key=f"svc_amt_{i}",
                                  min_value=0.0, step=10.0, format="%.2f")
            if amt and amt > 0:
                st.caption("→ %s incl GST  (auto, 10%% GST added)"
                           % _money(round(amt * GST_RATE, 2)))
if st.button("➕ Add service line"):
    add_service()
    st.rerun()

_ensure("inclusions", "")
st.text_area("What the investment includes (one per line)", key="inclusions", height=150)

# 5. Service terms -----------------------------------------------------------
st.header("5 · Service terms")
if "terms_df" not in st.session_state:
    st.session_state.terms_df = pd.DataFrame(
        default_terms if default_terms else [{"label": "", "value": ""}]
    ).rename(columns={"label": "Term", "value": "Detail"})
terms_edited = st.data_editor(
    st.session_state.terms_df, num_rows="dynamic", width="stretch",
    column_config={
        "Term": st.column_config.TextColumn("Term", width="medium"),
        "Detail": st.column_config.TextColumn("Detail", width="large"),
    },
    key="terms_editor",
)

# 6. Optional extra sections -------------------------------------------------
st.header("6 · Optional sections")
st.caption("Add Executive Summary, Transition Plan, Testimonials, etc. "
           "Each is placed where you choose. Leave empty for a standard 4-section proposal.")
qc1, qc2 = st.columns([3, 1])
with qc1:
    quick = st.selectbox("Quick-add a common section", ["—"] + list(SECTION_PRESETS))
with qc2:
    st.write("")
    st.write("")
    if st.button("Add", disabled=(quick == "—")):
        sp = SECTION_PRESETS[quick]
        add_section(quick, sp["placement"], sp["paragraphs"], sp["bullets"])
        st.rerun()
if st.button("➕ Add blank section"):
    add_section()
    st.rerun()

for i in list(st.session_state.section_ids):
    with st.container(border=True):
        cc1, cc2 = st.columns([3, 1])
        with cc1:
            st.text_input("Section title", key=f"sec_title_{i}")
        with cc2:
            st.write("")
            st.write("")
            if st.button("Remove", key=f"rm_sec_{i}"):
                st.session_state.section_ids.remove(i)
                st.rerun()
        st.selectbox("Where it goes", list(PLACEMENT_LABELS), key=f"sec_place_{i}")
        st.text_area("Paragraphs (one per line)", key=f"sec_paras_{i}", height=100)
        st.text_area("Bullet points (one per line, optional)", key=f"sec_bullets_{i}", height=90)

# Save draft + Generate ------------------------------------------------------
_terms_records = [{"label": str(r.get("Term", "")), "value": str(r.get("Detail", ""))}
                  for r in terms_edited.to_dict("records")]

st.divider()
_sd1, _sd2 = st.columns(2)
with _sd1:
    st.download_button(
        "💾 Save draft (.json)",
        json.dumps(collect_form(_terms_records), indent=2, ensure_ascii=False).encode("utf-8"),
        "ACS draft - %s.json" % proposal.safe_filename(st.session_state.get("client_name") or "draft"),
        mime="application/json", width="stretch",
    )
with _sd2:
    _gen = st.button("Generate proposal", type="primary", width="stretch")
st.caption("Not auto-saved — Save a draft before refreshing or closing so you don't lose your work.")

if _gen:
    if not st.session_state.get("client_name", "").strip():
        st.error("Please enter a client name first.")
    else:
        ss = st.session_state
        data = collect_form(_terms_records)
        data["scope_style"] = "area"
        data["duration"] = ""
        data["services"] = [_service_for_render(i) for i in ss.service_ids]
        data["extra_sections"] = [
            {"title": s["title"],
             "placement": PLACEMENT_LABELS.get(s["placement"], "after_cover"),
             "paragraphs": s["paragraphs"], "bullets": s["bullets"]}
            for s in data["extra_sections"]
        ]
        with st.spinner("Building your proposal…"):
            docx_bytes = proposal.render_docx_bytes(data)
        fname = "ACS Proposal - " + proposal.safe_filename(data["client_name"])
        st.session_state.gen_count = st.session_state.get("gen_count", 0) + 1
        st.session_state.result = {
            "docx": docx_bytes, "pdf": None, "fname": fname,
            "v": st.session_state.gen_count, "sig": _form_sig(_terms_records),
        }

res = st.session_state.get("result")
if res:
    stale = res.get("sig") and res["sig"] != _form_sig(_terms_records)
    if stale:
        st.warning("⚠️ You've changed the form since this was generated. "
                   "Click **Generate proposal** again to refresh the download below "
                   "(otherwise it will be the previous version).")
    else:
        st.success("Proposal ready — download below.")
    st.download_button(
        "⬇ Download Word (.docx)", res["docx"], res["fname"] + ".docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key="dl_docx_%d" % res["v"],
    )
    if res["pdf"]:
        st.download_button("⬇ Download PDF", res["pdf"], res["fname"] + ".pdf",
                           mime="application/pdf", key="dl_pdf_%d" % res["v"])
    elif res.get("pdf_failed"):
        st.info("PDF couldn't be produced here — the Word document is ready. "
                "(On the deployed app, Create PDF works via LibreOffice.)")
    else:
        if st.button("Create PDF", key="make_pdf_%d" % res["v"]):
            with st.spinner("Creating PDF…"):
                pdf = proposal.docx_bytes_to_pdf_bytes(res["docx"])
            res["pdf_failed"] = pdf is None
            res["pdf"] = pdf
            st.session_state.result = res
            st.rerun()
        st.caption("PDF is generated only when you click — keeps the app fast and light.")
