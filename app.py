"""
ACS Commercial Cleaning - Proposal Generator (web app).

A single-password Streamlit app. Fill the form, click Generate, download the
branded Word doc and PDF. Deploy free on Streamlit Community Cloud.
"""

import json
import os

import pandas as pd
import streamlit as st

import proposal

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO = os.path.join(HERE, "assets", "acs_logo.png")
PRESETS_PATH = os.path.join(HERE, "presets.json")

st.set_page_config(page_title="ACS Proposal Generator", page_icon="🧽", layout="centered")


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
        expected = "changeme"  # local fallback; set a real one in Cloud secrets
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
# State helpers for dynamic areas / services
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
    st.session_state[f"svc_ex_{i}"] = svc.get("price_ex", "")
    st.session_state[f"svc_period_{i}"] = svc.get("price_period", "per month")
    st.session_state[f"svc_inc_{i}"] = svc.get("price_inc", "")
    st.session_state.service_ids.append(i)


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


# initialise state on first load
if "area_ids" not in st.session_state:
    st.session_state.area_ids = []
    st.session_state.service_ids = []
    add_area()
    add_service()


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
    if st.button("Apply preset", disabled=(choice == "— Blank —")):
        idx = names.index(choice) - 1
        apply_preset(presets[idx])
        st.rerun()

# ---------------------------------------------------------------------------
# 1. Client details
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# 2. Cover letter
# ---------------------------------------------------------------------------
st.header("2 · Cover letter")
_ensure("cover_paragraphs", "")
st.text_area("Letter paragraphs (one paragraph per line)", key="cover_paragraphs", height=170)

# ---------------------------------------------------------------------------
# 3. Scope of work
# ---------------------------------------------------------------------------
st.header("3 · Scope of work")
c3, c4 = st.columns(2)
with c3:
    _ensure("frequency", "")
    st.text_input("Frequency", key="frequency", placeholder="2 days per week, after hours")
with c4:
    _ensure("duration", "")
    st.text_input("Duration", key="duration", placeholder="As required to complete scope")

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
if st.button("➕ Add area"):
    add_area()
    st.rerun()

_ensure("service_notes", "")
st.text_area("Service notes (one per line)", key="service_notes", height=110)

# ---------------------------------------------------------------------------
# 4. Investment options
# ---------------------------------------------------------------------------
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
            st.text_input("Price (ex GST)", key=f"svc_ex_{i}", placeholder="$1,040 plus GST")
            st.text_input("Price (incl GST)", key=f"svc_inc_{i}", placeholder="$1,144 incl GST")
if st.button("➕ Add service line"):
    add_service()
    st.rerun()

_ensure("inclusions", "")
st.text_area("What the investment includes (one per line)", key="inclusions", height=150)

# ---------------------------------------------------------------------------
# 5. Service terms
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Generate
# ---------------------------------------------------------------------------
st.divider()
make_pdf = st.checkbox("Also create PDF", value=True)

if st.button("Generate proposal", type="primary"):
    if not st.session_state.get("client_name", "").strip():
        st.error("Please enter a client name first.")
    else:
        data = {
            "client_name": st.session_state.client_name,
            "site_office": st.session_state.site_office,
            "contact_name": st.session_state.contact_name,
            "contact_title": st.session_state.contact_title,
            "contact_phone": st.session_state.contact_phone,
            "client_address": st.session_state.client_address,
            "date": st.session_state.date,
            "reference": st.session_state.reference,
            "cover_paragraphs": st.session_state.cover_paragraphs,
            "frequency": st.session_state.frequency,
            "duration": st.session_state.duration,
            "areas": [
                {"name": st.session_state.get(f"area_name_{i}", ""),
                 "items": st.session_state.get(f"area_items_{i}", "")}
                for i in st.session_state.area_ids
            ],
            "service_notes": st.session_state.service_notes,
            "services": [
                {"name": st.session_state.get(f"svc_name_{i}", ""),
                 "description": st.session_state.get(f"svc_desc_{i}", ""),
                 "schedule": st.session_state.get(f"svc_sched_{i}", ""),
                 "price_ex": st.session_state.get(f"svc_ex_{i}", ""),
                 "price_period": st.session_state.get(f"svc_period_{i}", ""),
                 "price_inc": st.session_state.get(f"svc_inc_{i}", "")}
                for i in st.session_state.service_ids
            ],
            "inclusions": st.session_state.inclusions,
            "terms": [
                {"label": str(r.get("Term", "")), "value": str(r.get("Detail", ""))}
                for r in terms_edited.to_dict("records")
            ],
        }
        with st.spinner("Building your proposal…"):
            docx_bytes = proposal.render_docx_bytes(data)
            pdf_bytes = proposal.docx_bytes_to_pdf_bytes(docx_bytes) if make_pdf else None
        fname = "ACS Proposal - " + proposal.safe_filename(data["client_name"])
        st.session_state.result = {
            "docx": docx_bytes, "pdf": pdf_bytes, "fname": fname,
            "pdf_wanted": make_pdf,
        }

# show download buttons (persist across reruns)
res = st.session_state.get("result")
if res:
    st.success("Proposal ready.")
    st.download_button(
        "⬇ Download Word (.docx)", res["docx"], res["fname"] + ".docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    if res["pdf"]:
        st.download_button("⬇ Download PDF", res["pdf"], res["fname"] + ".pdf",
                           mime="application/pdf")
    elif res["pdf_wanted"]:
        st.info("PDF couldn't be produced here — the Word document is ready. "
                "On the deployed app (with LibreOffice) the PDF is created automatically.")
