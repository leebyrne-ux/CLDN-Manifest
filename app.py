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

# Fetch Credentials directly without displaying sidebar input boxes
GMAIL_USER = st.secrets.get("GMAIL_USER", "lee.byrne@gogginstransport.ie")
GMAIL_APP_PASS = st.secrets.get("GMAIL_APP_PASS", "onru ktez ivct utnp")
GMAIL_LABEL = "AA Shipping/CLDN"


def fetch_latest_cldn_email(user, app_pass):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(user, app_pass)

        status, _ = mail.select(f'"{GMAIL_LABEL}"')
        if status != "OK":
            mail.select("inbox")

        status, data = mail.search(None, '(SUBJECT "Sailing Confirmation")')
        mail_ids = data[0].split()

        if not mail_ids:
            return ""

        latest_email_id = mail_ids[-1]
        status, data = mail.fetch(latest_email_id, "(RFC822)")
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
        st.error(f"Gmail Error: {e}")
        return ""


# File Uploader component
uploaded_file = st.file_uploader(
    "Drop Qargo File (.csv or .xlsx)", type=["csv", "xlsx", "xls"]
)

if uploaded_file and GMAIL_APP_PASS:
    if st.button("Run Reconciliation", type="primary"):
        with st.spinner("Fetching email & analyzing file..."):
            # 1. Fetch Email Units
            email_body = fetch_latest_cldn_email(GMAIL_USER, GMAIL_APP_PASS)
            cldn_raw = re.findall(r"GTC[\s-]?\d+", email_body, re.IGNORECASE)
            cldn_units = set(
                re.sub(r"[\s-]", "", t).upper() for t in cldn_raw
            )

            # 2. Process Uploaded File (Pandas handles corrupt CSVs and Excel automatically)
            try:
                if uploaded_file.name.endswith((".xlsx", ".xls")):
                    df = pd.read_excel(uploaded_file)
                else:
                    df = pd.read_csv(uploaded_file, encoding_errors="ignore")

                content = df.to_string()
                qargo_raw = re.findall(r"GTC[\s-]?\d+", content, re.IGNORECASE)
                qargo_units = set(
                    re.sub(r"[\s-]", "", t).upper() for t in qargo_raw
                )
            except Exception as e:
                st.error(f"Error reading file: {e}")
                qargo_units = set()
# --- ADD THIS DEBUG BLOCK RIGHT BEFORE CATEGORIZING UNITS ---
st.divider()
st.subheader("🔍 Diagnostics & Debug Info")

# Show which email subject was fetched
st.write(f"**Extracted {len(cldn_units)} units from CLdN Email:**")
st.code(sorted(list(cldn_units)))

# Show what Qargo extracted
st.write(f"**Extracted {len(qargo_units)} units from Qargo Upload:**")
st.code(sorted(list(qargo_units)))
            # 3. Categorize
            all_trailers = sorted(list(qargo_units.union(cldn_units)))

            matched_count = 0
            missing_count = 0
            forward_count = 0

            st.divider()
            st.subheader("Reconciliation Results")

            for trailer in all_trailers:
                in_qargo = trailer in qargo_units
                in_cldn = trailer in cldn_units

                if in_qargo and in_cldn:
                    st.success(f"• **{trailer}** - Matched")
                    matched_count += 1
                elif in_qargo and not in_cldn:
                    st.error(f"• **{trailer}** - Not Present (Left Behind)")
                    missing_count += 1
                elif not in_qargo and in_cldn:
                    st.info(f"• **{trailer}** - Forward Shipped")
                    forward_count += 1

            st.divider()
            col1, col2, col3 = st.columns(3)
            col1.metric("Matched", matched_count)
            col2.metric("Not Present", missing_count)
            col3.metric("Forward Shipped", forward_count)
