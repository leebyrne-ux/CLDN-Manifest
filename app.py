import email
import email.utils
import html
import imaplib
import re
import socket
import zoneinfo
from bs4 import BeautifulSoup
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CLdN Manifest Reconciliation", page_icon="🚢", layout="wide"
)

st.title("🚢 CLdN Manifest Reconciliation")
st.write(
    "Upload your Qargo export file (.csv or .xlsx) to reconcile trailers, instructions, GMR, PBN, and Booking Refs against CLdN sailing emails."
)

GMAIL_USER = st.secrets.get("GMAIL_USER", "lee.byrne@gogginstransport.ie")
GMAIL_APP_PASS = st.secrets.get("GMAIL_APP_PASS", "")
GMAIL_LABEL = "AA Shipping/CLDN"

socket.setdefaulttimeout(6.0)

# Precise pattern for GTC (3 digits) and GTUK (2 digits)
TRAILER_REGEX = r"\b(?:GTC\d{3}|GTUK\d{2})\b"


def format_to_gmt(raw_date_str):
    if not raw_date_str:
        return "Unknown Date"
    try:
        dt_utc = email.utils.parsedate_to_datetime(raw_date_str)
        local_tz = zoneinfo.ZoneInfo("Europe/Dublin")
        dt_local = dt_utc.astimezone(local_tz)
        return dt_local.strftime("%a, %d %b %Y %H:%M:%S %Z")
    except Exception:
        return raw_date_str


def fetch_sailing_emails_fast():
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", timeout=6)
        mail.login(GMAIL_USER, GMAIL_APP_PASS)

        status, _ = mail.select(f'"{GMAIL_LABEL}"')
        if status != "OK":
            status, _ = mail.select("INBOX")
            if status != "OK":
                st.error(f"❌ Could not select folder '{GMAIL_LABEL}' or 'INBOX'")
                return {}

        status, data = mail.search(None, "ALL")
        mail_ids = data[0].split()

        if not mail_ids:
            mail.logout()
            return {}

        # Scan up to the last 50 emails
        recent_ids = list(reversed(mail_ids[-50:]))
        email_options = {}

        for eid in recent_ids:
            status, data = mail.fetch(
                eid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])"
            )
            if not data or not data[0]:
                continue
            msg = email.message_from_bytes(data[0][1])
            subject = msg.get("Subject", "No Subject")
            raw_date = msg.get("Date", "")

            # Flexible subject match for sailing/manifest emails
            subj_lower = subject.lower()
            if any(term in subj_lower for term in ["sailing", "manifest", "confirmation", "cldn"]):
                formatted_date = format_to_gmt(raw_date)
                label_text = f"{subject} — ({formatted_date})"
                email_options[label_text] = eid

        mail.logout()
        return email_options
    except Exception as e:
        st.error(
            f"⚠️ Could not connect to Gmail IMAP automatically: {e}. Use the 'Paste Raw Email Text/HTML' tab instead!"
        )
        return {}


def fetch_email_body_fast(email_id):
    """Fetches HTML body specifically to preserve email table structures."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", timeout=6)
        mail.login(GMAIL_USER, GMAIL_APP_PASS)
        mail.select(f'"{GMAIL_LABEL}"')

        status, data = mail.fetch(email_id, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        html_body = ""
        plain_body = ""

        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                payload = part.get_payload(decode=True)
                if payload:
                    if ctype == "text/html":
                        html_body += payload.decode(errors="ignore") + "\n"
                    elif ctype == "text/plain":
                        plain_body += payload.decode(errors="ignore") + "\n"
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                if msg.get_content_type() == "text/html":
                    html_body = payload.decode(errors="ignore")
                else:
                    plain_body = payload.decode(errors="ignore")

        mail.logout()
        return html_body if html_body.strip() else plain_body
    except Exception as e:
        st.error(f"❌ Error fetching email content: {e}")
        return ""


def find_column(df, search_terms):
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for term in search_terms:
            if term.lower() in col_clean:
                return col
    return None


def clean_val(val):
    if not val or pd.isna(val) or str(val).lower() == "nan":
        return ""
    text = html.unescape(str(val)).replace("\xa0", " ")
    return re.sub(r"[^\w]", "", text).upper()


# 1. File Upload Section
uploaded_file = st.file_uploader(
    "1️⃣ Drop Qargo Export File (.csv, .xlsx, or .xls)",
    type=["csv", "xlsx", "xls"],
)

st.divider()

# 2. Email Input Modes
st.subheader("2️⃣ Select or Paste CLdN Manifest Email")

input_tab1, input_tab2 = st.tabs(
    ["📩 Fetch via Gmail IMAP", "📋 Paste Raw Email Text/HTML"]
)

selected_email_id = None
pasted_email_content = ""

with input_tab1:
    if "email_options" not in st.session_state:
        st.session_state.email_options = None

    c1, c2 = st.columns([3, 1])
    with c1:
        if st.session_state.email_options is None:
            if st.button("🔌 Connect & Load Recent Emails"):
                with st.spinner("Connecting to Gmail..."):
                    st.session_state.email_options = (
                        fetch_sailing_emails_fast()
                    )
                st.rerun()
        elif st.session_state.email_options:
            selected_label = st.selectbox(
                "Select CLdN Sailing Email:",
                options=list(st.session_state.email_options.keys()),
            )
            selected_email_id = st.session_state.email_options[selected_label]
        else:
            st.warning("No matching emails found in the folder.")

    with c2:
        if st.session_state.email_options is not None:
            if st.button("🔄 Refresh Emails"):
                st.session_state.email_options = fetch_sailing_emails_fast()
                st.rerun()

with input_tab2:
    pasted_email_content = st.text_area(
        "Paste email body or HTML table here:",
        height=180,
        placeholder="Paste the content of the CLdN email here if IMAP connection times out...",
    )

st.divider()

# 3. Reconciliation Execution
if uploaded_file and (selected_email_id or pasted_email_content):
    if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
        with st.spinner("Processing reconciliation..."):

            if pasted_email_content.strip():
                email_body = pasted_email_content
            else:
                email_body = fetch_email_body_fast(selected_email_id)

            try:
                if uploaded_file.name.endswith((".xlsx", ".xls")):
                    df = pd.read_excel(uploaded_file)
                else:
                    try:
                        df = pd.read_csv(uploaded_file)
                    except Exception:
                        uploaded_file.seek(0)
                        df = pd.read_csv(
                            uploaded_file, encoding_errors="ignore"
                        )

                df.columns = [str(c).strip() for c in df.columns]

            except Exception as e:
                st.error(f"❌ Error reading uploaded file: {e}")
                df = pd.DataFrame()

            if not df.empty and email_body:
                trailer_col = find_column(
                    df, ["trailers & instructions", "trailer"]
                )
                instr_col = find_column(
                    df, ["instructions", "instruction", "notes"]
                )
                gmr_col = find_column(df, ["gmr"])
                pbn_col = find_column(df, ["pbn"])
                booking_col = find_column(
                    df, ["booking ref", "booking reference", "booking"]
                )

                # Parse HTML structure with BeautifulSoup
                soup = BeautifulSoup(email_body, "html.parser")

                cldn_units = set()
                trailer_row_data = {}  # CLEAN_TRAILER -> Set of clean row tokens

                rows = soup.find_all("tr")
