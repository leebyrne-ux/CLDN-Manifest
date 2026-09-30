import email
import email.utils
import imaplib
import re
import zoneinfo
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CLdN Manifest Reconciliation", page_icon="🚢", layout="wide"
)

st.title("🚢 CLdN Manifest Reconciliation")
st.write(
    "Upload your Qargo export file (.csv or .xlsx) to reconcile trailers, instructions, GMR, PBN, and Booking Refs against CLdN sailing emails."
)

# Fetch Credentials from Secrets or defaults
GMAIL_USER = st.secrets.get("GMAIL_USER", "lee.byrne@gogginstransport.ie")
GMAIL_APP_PASS = st.secrets.get("GMAIL_APP_PASS", "bshg mcwd kahp xpqa")
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

        recent_ids = mail_ids[-limit:]
        recent_ids.reverse()

        email_options = {}
        for eid in recent_ids:
            status, data = mail.fetch(eid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])")
            msg = email.message_from_bytes(data[0][1])
            subject = msg.get("Subject", "No Subject")
            raw_date = msg.get("Date", "")

            formatted_date = format_to_gmt(raw_date)
            label_text = f"{subject} — ({formatted_date})"
            email_options[label_text] = eid

        mail.logout()
        return email_options
    except Exception as e:
        st.error(f"❌ Error fetching email list: {e}")
        return {}


def fetch_email_body_by_id(user, app_pass, label, email_id):
    """Fetches full body text for a selected email ID."""
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


def find_column(df, search_terms):
    """Finds a column name in dataframe matching any search terms (case-insensitive)."""
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for term in search_terms:
            if term.lower() in col_clean:
                return col
    return None


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

            # Read Qargo File into DataFrame
            try:
                if uploaded_file.name.endswith((".xlsx", ".xls")):
                    df = pd.read_excel(uploaded_file)
                else:
                    try:
                        df = pd.read_csv(uploaded_file)
                    except Exception:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, encoding_errors="ignore")

                df.columns = [str(c).strip() for c in df.columns]

            except Exception as e:
                st.error(f"❌ Error reading uploaded file: {e}")
                df = pd.DataFrame()

            if not df.empty:
                # Locate specific columns dynamically
                trailer_col = find_column(df, ["trailers & instructions", "trailer"])
                instr_col = find_column(df, ["instructions", "instruction", "notes"])
                gmr_col = find_column(df, ["gmr"])
                pbn_col = find_column(df, ["pbn"])
                booking_col = find_column(df, ["booking ref", "booking reference", "booking"])

                # Extract units directly from CLdN Email
                cldn_raw = re.findall(r"GTC[\s-]?\d+", email_body, re.IGNORECASE)
                cldn_units = set(re.sub(r"[\s-]", "", t).upper() for t in cldn_raw)

                st.divider()
                st.subheader("📋 Reconciliation Results")

                results_data = []

                for idx, row in df.iterrows():
                    # Parse Trailer ID
                    raw_trailer = str(row[trailer_col]) if trailer_col and pd.notna(row[trailer_col]) else ""
                    trailer_match = re.search(r"GTC[\s-]?\d+", raw_trailer, re.IGNORECASE)
                    clean_trailer = re.sub(r"[\s-]", "", trailer_match.group(0)).upper() if trailer_match else raw_trailer.strip()

                    if not clean_trailer or clean_trailer.lower() == "nan":
                        continue

                    # 1. Sailing Status
                    in_email = clean_trailer in cldn_units
                    status = "Matched 🟢" if in_email else "Not Present (Left Behind) 🔴"

                    # 2. Smart Instructions Check
                    raw_instr = str(row[instr_col]).strip() if instr_col and pd.notna(row[instr_col]) else ""
                    instr_trailer_match = re.search(r"GTC[\s-]?\d+", raw_instr, re.IGNORECASE)
                    instr_trailer_clean = re.sub(r"[\s-]", "", instr_trailer_match.group(0)).upper() if instr_trailer_match else ""

                    clean_raw_text = raw_instr if raw_instr.lower() != "nan" else ""
                    extra_notes = re.sub(r"GTC[\s-]?\d+", "", clean_raw_text, flags=re.IGNORECASE).strip() if clean_raw_text else ""

                    if not clean_raw_text:
                        instr_status = "Clean (Blank) 🟢"
                    elif instr_trailer_clean and instr_trailer_clean != clean_trailer:
                        instr_status = f"⚠️️ Mismatch: {clean_raw_text} (Expected {clean_trailer})"
                    elif extra_notes:
                        instr_status = f"⚠️ Flagged Note: {clean_raw_text}"
                    else:
                        instr_status = "Clean 🟢"

                    # 3. Cross-reference GMR
                    raw_gmr = str(row[gmr_col]).strip() if gmr_col and pd.notna(row[gmr_col]) else ""
                    gmr_status = "-"
                    if raw_gmr and raw_gmr.lower() != "nan":
                        gmr_in_email = raw_gmr.lower() in email_body.lower()
                        gmr_status = f"{raw_gmr} " + ("🟢 Found" if gmr_in_email else "🔴 Missing in Email")

                    # 4. Cross-reference PBN
                    raw_pbn = str(row[pbn_col]).strip() if pbn_col and pd.notna(row[pbn_col]) else ""
                    pbn_status = "-"
                    if raw_pbn and raw_pbn.lower() != "nan":
                        pbn_in_email = raw_pbn.lower() in email_body.lower()
                        pbn_status = f"{raw_pbn} " + ("🟢 Found" if pbn_in_email else "🔴 Missing in Email")

                    # 5. Cross-reference Booking Ref
                    raw_booking = str(row[booking_col]).strip() if booking_col and pd.notna(row[booking_col]) else ""
                    booking_status = "-"
                    if raw_booking and raw_booking.lower() != "nan":
                        booking_in_email = raw_booking.lower() in email_body.lower()
                        booking_status = f"{raw_booking} " + ("🟢 Found" if booking_in_email else "🔴 Missing in Email")

                    results_data.append({
                        "Trailer ID": clean_trailer,
                        "Sailing Status": status,
                        "Instructions Check": instr_status,
                        "GMR Ref": gmr_status,
                        "PBN Ref": pbn_status,
                        "Booking Ref": booking_status
                    })

                # Check for Forward Shipped units
                qargo_found_units = set(r["Trailer ID"] for r in results_data)
                forward_shipped = cldn_units - qargo_found_units

                for f_unit in sorted(list(forward_shipped)):
                    results_data.append({
                        "Trailer ID": f_unit,
                        "Sailing Status": "Forward Shipped 🔵",
                        "Instructions Check": "-",
                        "GMR Ref": "-",
                        "PBN Ref": "-",
                        "Booking Ref": "-"
                    })

                # Output Results as an Interactive Table
                results_df = pd.DataFrame(results_data)
                st.dataframe(results_df, use_container_width=True, hide_index=True)

                # Summary Totals
                st.divider()
                st.subheader("📊 Summary Metrics")
                c1, c2, c3, c4 = st.columns(4)

                matched_cnt = sum(1 for r in results_data if "Matched" in r["Sailing Status"])
                left_cnt = sum(1 for r in results_data if "Left Behind" in r["Sailing Status"])
                forward_cnt = sum(1 for r in results_data if "Forward Shipped" in r["Sailing Status"])
                flagged_cnt = sum(1 for r in results_data if "⚠️️" in r["Instructions Check"])

                c1.metric("Matched Trailers", matched_cnt)
                c2.metric("Left Behind", left_cnt)
                c3.metric("Forward Shipped", forward_cnt)
                c4.metric("Flagged Instructions", flagged_cnt)
