import email
import email.utils
import imaplib
import re
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
GMAIL_APP_PASS = st.secrets.get("GMAIL_APP_PASS", "bshg mcwd kahp xpqa")
GMAIL_LABEL = "AA Shipping/CLDN"


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


def fetch_email_options():
    """Connects to Gmail safely with timeout."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", timeout=10)
        mail.login(GMAIL_USER, GMAIL_APP_PASS)

        status, _ = mail.select(f'"{GMAIL_LABEL}"')
        if status != "OK":
            status, _ = mail.select("INBOX")
            if status != "OK":
                return {}

        status, data = mail.search(None, "ALL")
        mail_ids = data[0].split()

        if not mail_ids:
            return {}

        recent_ids = mail_ids[-10:]
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


def fetch_email_body(email_id):
    """Fetches body content for selected message ID."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com", timeout=10)
        mail.login(GMAIL_USER, GMAIL_APP_PASS)
        mail.select(f'"{GMAIL_LABEL}"')

        status, data = mail.fetch(email_id, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() in ["text/html", "text/plain"]:
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
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for term in search_terms:
            if term.lower() in col_clean:
                return col
    return None


def clean_val(val):
    if not val or pd.isna(val) or str(val).lower() == "nan":
        return ""
    return re.sub(r"[^\w]", "", str(val)).upper()


# Render Uploader
uploaded_file = st.file_uploader(
    "Drop Qargo File (.csv, .xlsx, or .xls)", type=["csv", "xlsx", "xls"]
)

# Fetch Email List
if GMAIL_APP_PASS:
    if "email_options" not in st.session_state:
        with st.spinner("Connecting to Gmail..."):
            st.session_state.email_options = fetch_email_options()

    if st.session_state.email_options:
        selected_label = st.selectbox(
            "📩 Select the CLdN Manifest Email to Reconcile:",
            options=list(st.session_state.email_options.keys()),
        )
        selected_email_id = st.session_state.email_options[selected_label]
        
        if st.button("🔄 Refresh Email List"):
            st.session_state.email_options = fetch_email_options()
            st.rerun()
    else:
        st.warning("⚠️ No emails found in folder or connection failed.")
else:
    st.warning("⚠️ Gmail App Password missing in Streamlit Secrets.")

# Execution
if uploaded_file and GMAIL_APP_PASS and selected_email_id:
    if st.button("Run Reconciliation", type="primary", use_container_width=True):
        with st.spinner("Processing reconciliation..."):

            email_body = fetch_email_body(selected_email_id)

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
                trailer_col = find_column(df, ["trailers & instructions", "trailer"])
                instr_col = find_column(df, ["instructions", "instruction", "notes"])
                gmr_col = find_column(df, ["gmr"])
                pbn_col = find_column(df, ["pbn"])
                booking_col = find_column(df, ["booking ref", "booking reference", "booking"])

                # Parse HTML structure
                soup = BeautifulSoup(email_body, "html.parser")
                trailer_blocks = {}
                cldn_units = set()

                rows = soup.find_all("tr")
                for tr in rows:
                    row_text = clean_val(tr.get_text(separator=" "))
                    matches = re.findall(r"GTC[\s-]?\d+", row_text, re.IGNORECASE)
                    for m in matches:
                        clean_m = clean_val(m)
                        cldn_units.add(clean_m)
                        trailer_blocks[clean_m] = row_text

                full_email_clean = clean_val(soup.get_text())

                st.divider()
                st.subheader("📋 Reconciliation Results")

                results_data = []

                for idx, row in df.iterrows():
                    raw_trailer = str(row[trailer_col]) if trailer_col and pd.notna(row[trailer_col]) else ""
                    trailer_match = re.search(r"GTC[\s-]?\d+", raw_trailer, re.IGNORECASE)
                    clean_trailer = clean_val(trailer_match.group(0)) if trailer_match else clean_val(raw_trailer)

                    if not clean_trailer:
                        continue

                    # 1. Sailing Status
                    in_email = clean_trailer in cldn_units
                    status = "Matched 🟢" if in_email else "Not Present (Left Behind) 🔴"

                    # 2. Check Instructions Column in Qargo
                    raw_instr = str(row[instr_col]).strip() if instr_col and pd.notna(row[instr_col]) else ""
                    instr_trailer_match = re.search(r"GTC[\s-]?\d+", raw_instr, re.IGNORECASE)
                    instr_trailer_clean = clean_val(instr_trailer_match.group(0)) if instr_trailer_match else ""

                    clean_raw_text = raw_instr if raw_instr.lower() != "nan" else ""
                    extra_notes = re.sub(r"GTC[\s-]?\d+", "", clean_raw_text, flags=re.IGNORECASE).strip() if clean_raw_text else ""

                    if not clean_raw_text:
                        instr_status = "Clean 🟢"
                    elif instr_trailer_clean and instr_trailer_clean != clean_trailer:
                        instr_status = f"⚠️ Mismatch: {clean_raw_text}"
                    elif extra_notes:
                        instr_status = f"⚠️ Flagged Note: {clean_raw_text}"
                    else:
                        instr_status = "Clean 🟢"

                    # 3. Reference Row Validator
                    t_block = trailer_blocks.get(clean_trailer, "")

                    def verify_ref(ref_val):
                        c_ref = clean_val(ref_val)
                        if not c_ref:
                            return "-"

                        if t_block and c_ref in t_block:
                            return f"{ref_val} 🟢 Matched"
                        elif c_ref in full_email_clean:
                            return f"{ref_val} ⚠️ Mismatch"
                        else:
                            return f"{ref_val} 🔴 Missing"

                    raw_gmr = str(row[gmr_col]) if gmr_col and pd.notna(row[gmr_col]) else ""
                    raw_pbn = str(row[pbn_col]) if pbn_col and pd.notna(row[pbn_col]) else ""
                    raw_booking = str(row[booking_col]) if booking_col and pd.notna(row[booking_col]) else ""

                    results_data.append({
                        "Trailer ID": clean_trailer,
                        "Sailing Status": status,
                        "Instructions Check": instr_status,
                        "GMR Ref": verify_ref(raw_gmr),
                        "PBN Ref": verify_ref(raw_pbn),
                        "Booking Ref": verify_ref(raw_booking)
                    })

                # Forward Shipped Units
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

                # Render Table
                results_df = pd.DataFrame(results_data)
                st.dataframe(results_df, use_container_width=True, hide_index=True)

                # Summary Metrics
                st.divider()
                st.subheader("📊 Summary Metrics")
                c1, c2, c3, c4 = st.columns(4)

                matched_cnt = sum(1 for r in results_data if "Matched" in r["Sailing Status"])
                left_cnt = sum(1 for r in results_data if "Left Behind" in r["Sailing Status"])
                forward_cnt = sum(1 for r in results_data if "Forward Shipped" in r["Sailing Status"])
                flagged_cnt = sum(1 for r in results_data if "⚠️" in r["Instructions Check"] or "⚠️" in r["PBN Ref"] or "⚠️" in r["GMR Ref"] or "⚠️️" in r["Booking Ref"])

                c1.metric("Matched Trailers", matched_cnt)
                c2.metric("Left Behind", left_cnt)
                c3.metric("Forward Shipped", forward_cnt)
                c4.metric("Flagged / Swapped", flagged_cnt)
