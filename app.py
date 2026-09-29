import email
import imaplib
import re
import pandas as pd
import streamlit as st

# Config & Page Setup
st.set_page_config(
    page_title="CLdN Manifest Reconciliation", page_icon="🚢", layout="centered"
)

st.title("🚢 CLdN Manifest Reconciliation")
st.write(
    "Upload your Qargo export file (.csv or .xlsx) below to reconcile against the latest CLdN sailing email."
)

# Fetch Credentials directly from Streamlit Secrets or defaults
GMAIL_USER = st.secrets.get("GMAIL_USER", "lee.byrne@gogginstransport.ie")
GMAIL_APP_PASS = st.secrets.get("GMAIL_APP_PASS", "")
GMAIL_LABEL = "AA Shipping/CLDN"


def fetch_latest_cldn_email(user, app_pass, label):
    """Connects to Gmail and fetches the latest Sailing Confirmation body text & metadata."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(user, app_pass)

        # Select label (handling spaces/slashes safely)
        status, _ = mail.select(f'"{label}"')
        if status != "OK":
            status, _ = mail.select("INBOX")
            if status != "OK":
                st.error("❌ Failed to access Gmail folder.")
                return "", "", ""

        # Search for Sailing Confirmation emails first, fall back to ALL if subject varies
        status, data = mail.search(None, '(SUBJECT "Sailing Confirmation")')
        mail_ids = data[0].split()

        if not mail_ids:
            status, data = mail.search(None, "ALL")
            mail_ids = data[0].split()

        if not mail_ids:
            st.error("❌ No emails found in the specified Gmail folder.")
            return "", "", ""

        latest_email_id = mail_ids[-1]
        status, data = mail.fetch(latest_email_id, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = msg.get("Subject", "Unknown Subject")
        date_sent = msg.get("Date", "Unknown Date")

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type in ["text/plain", "text/html"]:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body += payload.decode(errors="ignore") + "\n"
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                body = payload.decode(errors="ignore")

        mail.logout()
        return body, subject, date_sent

    except Exception as e:
        st.error(f"❌ Error connecting to Gmail: {e}")
        return "", "", ""


# File Upload Section
uploaded_file = st.file_uploader(
    "Drop Qargo File (.csv, .xlsx, or .xls)", type=["csv", "xlsx", "xls"]
)

if uploaded_file:
    if not GMAIL_APP_PASS:
        st.warning("⚠️ Gmail App Password missing in Streamlit Secrets.")
    else:
        if st.button("Run Reconciliation", type="primary", use_container_width=True):
            with st.spinner("Fetching email and checking records..."):

                # 1. Fetch Email Data
                email_body, subject, date_sent = fetch_latest_cldn_email(
                    GMAIL_USER, GMAIL_APP_PASS, GMAIL_LABEL
                )

                if email_body:
                    # Capture GTC numbers with optional spaces/dashes (e.g. GTC143 or GTC 143)
                    cldn_raw = re.findall(r"GTC[\s-]?\d+", email_body, re.IGNORECASE)
                    cldn_units = set(re.sub(r"[\s-]", "", t).upper() for t in cldn_raw)

                    # 2. Extract Qargo Data
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

                    # 3. Diagnostic Expander (Checks if wrong email was fetched)
                    with st.expander("🔍 Diagnostic Inspector (Click to verify source data)"):
                        st.write(f"**Fetched Email Subject:** `{subject}`")
                        st.write(f"**Email Date:** `{date_sent}`")
                        st.write("---")
                        col_a, col_b = st.columns(2)
                        with col_a:
                            st.write(f"**Units found on Email ({len(cldn_units)}):**")
                            st.json(sorted(list(cldn_units)))
                        with col_b:
                            st.write(f"**Units found in Qargo File ({len(qargo_units)}):**")
                            st.json(sorted(list(qargo_units)))

                    # 4. Process & Display Results
                    all_trailers = sorted(list(qargo_units.union(cldn_units)))

                    if not all_trailers:
                        st.warning("⚠️ No trailer IDs (GTCxxx) found in email or uploaded file.")
                    else:
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

                        # Summary Metrics
                        st.divider()
                        st.subheader("📊 Totals")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Matched", matched_count)
                        col2.metric("Not Present", missing_count)
                        col3.metric("Forward Shipped", forward_count)
