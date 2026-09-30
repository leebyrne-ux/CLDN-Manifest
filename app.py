import email
import email.utils
import imaplib
import re
import zoneinfo
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CLdN Manifest Reconciliation", page_icon="🚢", layout="centered"
)

st.title("🚢 CLdN Manifest Reconciliation")
st.write(
    "Upload your Qargo export file (.csv or .xlsx) below to reconcile against a CLdN sailing email."
)

# Fetch Credentials from Secrets or defaults
GMAIL_USER = st.secrets.get("GMAIL_USER", "lee.byrne@gogginstransport.ie")
GMAIL_APP_PASS = st.secrets.get("GMAIL_APP_PASS", "xqlu xnex rfik sgzd")
GMAIL_LABEL = "AA Shipping/CLDN"


def format_to_gmt(raw_date_str):
    """Converts raw email date headers (UTC) to formatted local GMT/IST time."""
    if not raw_date_str:
        return "Unknown Date"
    try:
        dt_utc = email.utils.parsedate_to_datetime(raw_date_str)
        local_tz = zoneinfo.ZoneInfo("Europe/Dublin")
        dt_local = dt_utc.astimezone(local_tz)
        return dt_local.strftime("%a, %d %b %Y %H:%M:%S %Z")
    except Exception:
        return raw_date_str


def get_recent_emails_list(user, app_pass, label, limit=10):
    """Fetches subjects and GMT dates for the last N emails in the folder."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(user, app_pass)

        status, _ = mail.select(f'"{label}"')
        if status != "OK":
            status, _ = mail.select("INBOX")
            if status != "OK":
                return {}

        status, data = mail.search(None, "ALL")
        mail_ids = data[0].split()

        if not mail_ids:
            return {}

        # Grab the last N emails and reverse them (newest first)
        recent_ids = mail_ids[-limit:]
        recent_ids.reverse()

        email_options = {}
        for eid in recent_ids:
            status, data = mail.fetch(eid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])")
            msg = email.message_from_bytes(data[0][1])
            subject = msg.get("Subject", "No Subject")
            raw_date = msg.get("Date", "")
            
            # Convert UTC date to GMT/IST
            formatted_date = format_to_gmt(raw_date)
            
            label_text = f"{subject} — ({formatted_date})"
            email_options[label_text] = eid

        mail.logout()
        return email_options
    except Exception as e:
        st.error(f"❌ Error fetching email list: {e}")
        return {}


def fetch_email_body_by_id(user, app_pass, label, email_id):
    """Fetches the full body text for a selected email ID."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(user, app_pass)
        mail.select(f'"{label}"')

        status, data = mail.fetch(email_id, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() in ["text/plain", "text/html"]:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body += payload.decode(errors="ignore") + "\n"
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode(errors="ignore")

        mail.logout()
        return body
    except Exception as e:
        st.error(f"❌ Error fetching email body: {e}")
        return ""


# 1. Email Dropdown Selector
selected_email_id = None
if GMAIL_APP_PASS:
    email_options = get_recent_emails_list(GMAIL_USER, GMAIL_APP_PASS, GMAIL_LABEL)

    if email_options:
        selected_email_label = st.selectbox(
            "📩 Select the CLdN Manifest Email to Reconcile:",
            options=list(email_options.keys()),
        )
        selected_email_id = email_options[selected_email_label]
    else:
        st.warning("⚠️ No emails found in folder or connection failed.")
else:
    st.warning("⚠️ Gmail App Password missing in Streamlit Secrets.")

# 2. File Upload Section
uploaded_file = st.file_uploader(
    "Drop Qargo File (.csv, .xlsx, or .xls)", type=["csv", "xlsx", "xls"]
)

# 3. Execution & Results
if uploaded_file and GMAIL_APP_PASS and selected_email_id:
    if st.button("Run Reconciliation", type="primary", use_container_width=True):
        with st.spinner("Processing reconciliation..."):

            # Fetch selected email
            email_body = fetch_email_body_by_id(
                GMAIL_USER, GMAIL_APP_PASS, GMAIL_LABEL, selected_email_id
            )

            # Extract units from email
            cldn_raw = re.findall(r"GTC[\s-]?\d+", email_body, re.IGNORECASE)
            cldn_units = set(re.sub(r"[\s-]", "", t).upper() for t in cldn_raw)

            # Extract units from Qargo file
            qargo_units = set()
            try:
                if uploaded_file.name.endswith((".xlsx", ".xls")):
                    df = pd.read_excel(uploaded_file)
                else:
                    try:
                        df = pd.read_csv(uploaded_file)
                    except Exception:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, encoding_errors="ignore")

                content = df.to_string()
                qargo_raw = re.findall(r"GTC[\s-]?\d+", content, re.IGNORECASE)
                qargo_units = set(re.sub(r"[\s-]", "", t).upper() for t in qargo_raw)
            except Exception as e:
                st.error(f"❌ Error processing uploaded file: {e}")

            # Display Results
            all_trailers = sorted(list(qargo_units.union(cldn_units)))

            st.divider()
            st.subheader("📋 Itemized Trailer Breakdown")

            matched_count = 0
            missing_count = 0
            forward_count = 0

            for trailer in all_trailers:
                in_qargo = trailer in qargo_units
                in_cldn = trailer in cldn_units

                if in_qargo and in_cldn:
                    st.success(f"• **{trailer}** — Matched")
                    matched_count += 1
                elif in_qargo and not in_cldn:
                    st.error(f"• **{trailer}** — Not Present (Left Behind)")
                    missing_count += 1
                elif not in_qargo and in_cldn:
                    st.info(f"• **{trailer}** — Forward Shipped")
                    forward_count += 1

            st.divider()
            col1, col2, col3 = st.columns(3)
            col1.metric("Matched", matched_count)
            col2.metric("Not Present", missing_count)
            col3.metric("Forward Shipped", forward_count)
