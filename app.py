import email
import email.header
import email.utils
import html
import imaplib
import re
import socket
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup


# ============================================================
# PAGE SETUP AND SETTINGS
# ============================================================

st.set_page_config(
    page_title="CLdN Manifest Reconciliation",
    page_icon="🚢",
    layout="wide",
)

st.title("🚢 CLdN Manifest Reconciliation")
st.write(
    "Upload a Qargo export. The app will identify the CLdN Sailing "
    "Confirmation using the intermodal departure date and route, then "
    "check trailers, instructions, GMR, PBN and booking references."
)

GMAIL_USER = st.secrets.get(
    "GMAIL_USER", "lee.byrne@gogginstransport.ie"
)
GMAIL_APP_PASS = st.secrets.get("GMAIL_APP_PASS", "")
GMAIL_LABEL = "AA Shipping/CLDN"
IMAP_HOST = "imap.gmail.com"
IMAP_TIMEOUT = 12
socket.setdefaulttimeout(IMAP_TIMEOUT)

TRAILER_REGEX = r"\b(?:GTC\d{3}|GTUK\d{2})\b"
DIRECTION_MAP = {
    "DUB": ("DUBLIN", "LIVERPOOL"),
    "LIV": ("LIVERPOOL", "DUBLIN"),
}
DUBLIN_TZ = ZoneInfo("Europe/Dublin")


# ============================================================
# GENERAL HELPERS
# ============================================================

def decode_header_value(value):
    """Decode MIME-encoded header text into readable text."""
    if not value:
        return ""
    chunks = email.header.decode_header(value)
    result = []
    for chunk, charset in chunks:
        if isinstance(chunk, bytes):
            try:
                result.append(chunk.decode(charset or "utf-8", errors="replace"))
            except LookupError:
                result.append(chunk.decode("utf-8", errors="replace"))
        else:
            result.append(str(chunk))
    return "".join(result)


def clean_val(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = html.unescape(str(value)).replace("\xa0", " ").strip()
    if text.lower() in {"", "nan", "none", "nat"}:
        return ""
    return re.sub(r"[^\w]", "", text).upper()


def find_column(df, search_terms):
    lowered = [(col, str(col).strip().lower()) for col in df.columns]
    # Prefer exact or more-specific matches first.
    for term in search_terms:
        for col, name in lowered:
            if name == term.lower():
                return col
    for term in search_terms:
        for col, name in lowered:
            if term.lower() in name:
                return col
    return None


def as_datetime_local(value):
    """Parse a Qargo Intermodal Departure into Europe/Dublin date and time.

    ISO values with an explicit offset (e.g. 2026-10-14T12:00:00+01:00)
    are converted to Europe/Dublin. Naive values are treated as local Irish time.
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()
    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime(value.year, value.month, value.day)
    else:
        text = str(value).strip()
        if not text or text.lower() in {"nan", "nat", "none"}:
            return None
        try:
            parsed = pd.to_datetime(text, dayfirst=False, errors="coerce")
            if pd.isna(parsed):
                # Qargo exports may also use a day-first display format.
                parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
            if pd.isna(parsed):
                return None
            if isinstance(parsed, pd.Timestamp):
                parsed = parsed.to_pydatetime()
        except (TypeError, ValueError, OverflowError):
            return None

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=DUBLIN_TZ)
    else:
        parsed = parsed.astimezone(DUBLIN_TZ)
    return parsed.replace(second=0, microsecond=0)


def direction_from_booking_refs(values):
    prefixes = set()
    examples = []
    for value in values:
        raw = str(value).strip().upper()
        if not raw or raw in {"NAN", "NONE"}:
            continue
        match = re.search(r"\b(DUB|LIV)[A-Z0-9-]*", raw)
        if match:
            prefix = match.group(1)
            prefixes.add(prefix)
            examples.append(raw)
    if not prefixes:
        return None, examples
    if len(prefixes) > 1:
        return "CONFLICT", examples
    return next(iter(prefixes)), examples


def direction_text(prefix):
    if prefix not in DIRECTION_MAP:
        return "Unknown route"
    origin, destination = DIRECTION_MAP[prefix]
    return f"{origin.title()} → {destination.title()}"


# ============================================================
# GMAIL / MIME HELPERS
# ============================================================

def open_gmail_folder(mail):
    """Select the configured label, trying common Gmail IMAP spellings."""
    candidates = [
        f'"{GMAIL_LABEL}"',
        GMAIL_LABEL,
        f'"[Gmail]/{GMAIL_LABEL}"',
    ]
    last_error = None
    for folder in candidates:
        try:
            status, data = mail.select(folder, readonly=True)
            if status == "OK":
                return folder
            last_error = data
        except Exception as exc:
            last_error = str(exc)
    raise RuntimeError(
        f"Could not open Gmail label '{GMAIL_LABEL}'. "
        f"Check the label name and IMAP access. Details: {last_error}"
    )


def extract_email_content(msg):
    """Return best readable body text plus attachment filenames."""
    html_parts = []
    plain_parts = []
    attachment_names = []

    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
        filename = decode_header_value(part.get_filename())
        disposition = (part.get("Content-Disposition") or "").lower()
        if filename:
            attachment_names.append(filename)

        payload = part.get_payload(decode=True)
        if not payload:
            continue

        charset = part.get_content_charset() or "utf-8"
        try:
            content = payload.decode(charset, errors="replace")
        except LookupError:
            content = payload.decode("utf-8", errors="replace")

        is_attachment = "attachment" in disposition or bool(filename)
        content_type = part.get_content_type().lower()

        if is_attachment:
            lower_name = filename.lower()
            if lower_name.endswith((".txt", ".csv", ".html", ".htm")):
                if content_type == "text/html" or lower_name.endswith((".html", ".htm")):
                    html_parts.append(content)
                else:
                    plain_parts.append(content)
            continue

        if content_type == "text/html":
            html_parts.append(content)
        elif content_type == "text/plain":
            plain_parts.append(content)

    if html_parts:
        return "\n".join(html_parts), attachment_names
    return "\n".join(plain_parts), attachment_names


def body_to_text(body):
    if not body:
        return ""
    soup = BeautifulSoup(body, "html.parser")
    return re.sub(r"[ \t]+", " ", soup.get_text("\n", strip=True)).replace("\xa0", " ")


def fetch_sailing_emails_for_date(sailing_datetime):
    """Search Gmail for confirmations mentioning the Qargo sailing date.

    Gmail IMAP cannot search on a sailing date as a structured field because
    that date lives inside the email body, not the email's received-date header.
    Search the subject and the exact CLdN body date first, then download full
    messages only for those candidates. If this targeted search returns none,
    do not broaden into downloading hundreds of unrelated emails.
    """
    if not GMAIL_APP_PASS:
        raise RuntimeError(
            "GMAIL_APP_PASS is missing from Streamlit secrets. "
            "Add your Gmail app password in the app's Secrets settings."
        )

    date_text = sailing_datetime.strftime("%d/%m/%Y")
    mail = None
    results = []
    standard_candidates = []
    bookings_candidates = []
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, timeout=IMAP_TIMEOUT)
        mail.login(GMAIL_USER, GMAIL_APP_PASS)
        folder = open_gmail_folder(mail)

        # CLdN confirmations show the scheduled sailing as e.g. 14/10/2026.
        # The body search narrows results before any full email is downloaded.
        status, data = mail.uid(
            "search", None,
            'SUBJECT "Sailing Confirmation" BODY "' + date_text + '"'
        )
        if status != "OK":
            raise RuntimeError("Gmail search for date-specific Sailing Confirmations failed.")

        ids = data[0].split() if data and data[0] else []
        for message_id in reversed(ids):
            # Read only lightweight headers first. Ignore booking-notification
            # copies before downloading their full message bodies.
            status, header_data = mail.uid(
                "fetch",
                message_id,
                "(BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])",
            )
            if status != "OK" or not header_data:
                continue
            raw_header = next(
                (
                    item[1] for item in header_data
                    if isinstance(item, tuple)
                    and len(item) > 1
                    and isinstance(item[1], bytes)
                ),
                None,
            )
            if not raw_header:
                continue
            header_msg = email.message_from_bytes(raw_header)
            subject = decode_header_value(header_msg.get("Subject", ""))
            subject_lower = subject.lower()
            if "sailing confirmation" not in subject_lower:
                continue
            # Prefer the standard confirmation when it exists. Some sailings
            # arrive only as a [Bookings] message, so retain those as a fallback
            # rather than discarding them outright.
            candidate = (message_id, subject, "[bookings]" in subject_lower)
            if candidate[2]:
                bookings_candidates.append(candidate)
            else:
                standard_candidates.append(candidate)

        # Check both standard and [Bookings] confirmations. A standard email
        # for a different sailing on the same date must not hide the correct
        # [Bookings] confirmation. Preference is applied only after exact
        # scheduled date/time and route matching in the main app.
        chosen_candidates = standard_candidates + bookings_candidates

        for message_id, header_subject, is_bookings in chosen_candidates:
            status, full_data = mail.uid("fetch", message_id, "(RFC822)")
            if status != "OK" or not full_data:
                continue
            raw_message = next(
                (
                    item[1] for item in full_data
                    if isinstance(item, tuple)
                    and len(item) > 1
                    and isinstance(item[1], bytes)
                ),
                None,
            )
            if not raw_message:
                continue
            msg = email.message_from_bytes(raw_message)
            # Use the full message's decoded subject in the result.
            subject = decode_header_value(msg.get("Subject", ""))
            body, attachments = extract_email_content(msg)
            text = body_to_text(body)
            results.append(
                {
                    "id": message_id.decode(errors="ignore"),
                    "folder": folder,
                    "subject": subject,
                    "date_header": msg.get("Date", ""),
                    "body": body,
                    "text": text,
                    "attachments": attachments,
                    "is_bookings": "[bookings]" in subject.lower(),
                }
            )
        return results
    finally:
        if mail is not None:
            try:
                mail.logout()
            except Exception:
                pass


def parse_sailing_details(email_item):
    """Read route and scheduled sailing date/time, deliberately ignoring Departed."""
    text = email_item.get("text") or body_to_text(email_item.get("body", ""))
    upper = text.upper()

    route_match = re.search(
        r"\b(DUBLIN|LIVERPOOL)\s*(?:TO|→|->)\s*(DUBLIN|LIVERPOOL)\b",
        upper,
    )
    route = None
    if route_match:
        origin, destination = route_match.group(1), route_match.group(2)
        if origin != destination:
            route = (origin, destination)

    # CLdN confirmation format: Wednesday 14/10/2026 12:00.
    # Only parse the first date/time in the body before the "Departed:" line.
    before_departed = re.split(r"\bDEPARTED\s*:", upper, maxsplit=1)[0]
    date_match = re.search(
        r"\b(?:MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)?"
        r"\s*(\d{1,2})[/-](\d{1,2})[/-](\d{4})\s+(\d{1,2}):(\d{2})\b",
        before_departed,
    )
    scheduled = None
    if date_match:
        day, month, year, hour, minute = map(int, date_match.groups())
        try:
            scheduled = datetime(year, month, day, hour, minute, tzinfo=DUBLIN_TZ)
        except ValueError:
            scheduled = None

    # In the CLdN layout the vessel is usually the non-empty line immediately
    # before the route line (e.g. MAZARINE before DUBLIN to LIVERPOOL).
    vessel = ""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for index, line in enumerate(lines):
        if re.search(r"\b(DUBLIN|LIVERPOOL)\s+(?:TO|→|->)\s+(DUBLIN|LIVERPOOL)\b", line.upper()):
            for previous in reversed(lines[:index]):
                if previous.upper() not in {"CLDN", "SAILING CONFIRMATION", "GOGGINS TRANSPORT COMPANY"}:
                    vessel = previous
                    break
            break
    return {
        "route": route,
        "scheduled": scheduled,
        "vessel": vessel,
        "text": text,
    }


# ============================================================
# QARGO DATA
# ============================================================

def read_qargo_file(uploaded_file):
    filename = uploaded_file.name.lower()
    if filename.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding_errors="replace")
    if filename.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    raise ValueError("Please upload a CSV, XLSX or XLS file.")


def identify_qargo_fields(df):
    departure_col = find_column(
        df,
        ["intermodal departure", "intermodal departure date", "intermodal depart"],
    )
    booking_col = find_column(df, ["booking ref", "booking reference", "booking"])
    trailer_col = find_column(df, ["trailers & instructions", "trailer"])
    instr_col = find_column(df, ["instructions", "instruction", "notes"])
    gmr_col = find_column(df, ["gmr"])
    pbn_col = find_column(df, ["pbn"])

    missing = []
    if departure_col is None:
        missing.append("Intermodal Departure")
    if booking_col is None:
        missing.append("Booking Ref")
    if trailer_col is None:
        missing.append("Trailer (or Trailers & Instructions)")
    if missing:
        raise ValueError(
            "Could not identify these required Qargo columns: "
            + ", ".join(missing)
            + ". Please check the column headings in your export."
        )
    return {
        "departure": departure_col,
        "booking": booking_col,
        "trailer": trailer_col,
        "instructions": instr_col,
        "gmr": gmr_col,
        "pbn": pbn_col,
    }


def infer_qargo_sailing(df, columns):
    departure_datetimes = set()
    bad_departures = []
    for value in df[columns["departure"]].tolist():
        if str(value).strip() and str(value).strip().lower() not in {"nan", "nat", "none"}:
            parsed = as_datetime_local(value)
            if parsed:
                departure_datetimes.add(parsed)
            else:
                bad_departures.append(str(value))

    if not departure_datetimes:
        raise ValueError(
            "No usable date/time values were found in the Qargo 'Intermodal Departure' column."
        )
    if len(departure_datetimes) > 1:
        shown = ", ".join(
            dt.strftime("%d/%m/%Y %H:%M") for dt in sorted(departure_datetimes)
        )
        raise ValueError(
            "The Qargo export contains more than one Intermodal Departure date/time "
            f"({shown}). Please upload an export for one sailing only; the app "
            "will not guess which sailing you intend to reconcile."
        )

    prefix, examples = direction_from_booking_refs(df[columns["booking"]].tolist())
    if prefix == "CONFLICT":
        raise ValueError(
            "Booking references contain both DUB and LIV prefixes. "
            "Please check that the export contains one sailing direction only."
        )
    if prefix is None:
        raise ValueError(
            "Could not infer the sailing direction from the Booking Ref column. "
            "The app expects DUB-prefixed references for Dublin → Liverpool or "
            "LIV-prefixed references for Liverpool → Dublin."
        )

    return next(iter(departure_datetimes)), prefix


# ============================================================
# RECONCILIATION LOGIC (preserves original checks)
# ============================================================

def reconcile(df, email_text, columns):
    soup = BeautifulSoup(email_text, "html.parser")
    visible_text = soup.get_text("\n", strip=True) if soup.find() else email_text
    visible_text = html.unescape(visible_text).replace("\xa0", " ")
    cldn_units = set()
    trailer_row_data = {}

    for tr in soup.find_all("tr"):
        cells = tr.find_all(["td", "th"])
        if not cells:
            continue
        cell_texts = [
            cell.get_text(separator=" ", strip=True) for cell in cells
        ]
        cleaned_tokens = [clean_val(cell) for cell in cell_texts if clean_val(cell)]
        combined = " ".join(cleaned_tokens)
        for match in re.findall(TRAILER_REGEX, combined, re.IGNORECASE):
            trailer = clean_val(match)
            cldn_units.add(trailer)
            trailer_row_data.setdefault(trailer, set()).update(cleaned_tokens)

    if not cldn_units:
        for line in visible_text.splitlines():
            cleaned_line = clean_val(line)
            for match in re.findall(TRAILER_REGEX, cleaned_line, re.IGNORECASE):
                trailer = clean_val(match)
                cldn_units.add(trailer)
                trailer_row_data.setdefault(trailer, set()).add(cleaned_line)

    full_email_clean = clean_val(visible_text)
    results = []

    def value_from(row, column):
        if column is None:
            return ""
        value = row[column]
        return "" if pd.isna(value) else str(value).strip()

    def verify_ref(ref_value, row_tokens):
        clean_ref = clean_val(ref_value)
        if not clean_ref:
            return "-"
        if any(clean_ref in token or token in clean_ref for token in row_tokens):
            return f"{ref_value} 🟢 Matched"
        if clean_ref in full_email_clean:
            return f"{ref_value} ⚠️ Mismatch"
        return f"{ref_value} 🔴 Missing"

    for _, row in df.iterrows():
        raw_trailer = value_from(row, columns["trailer"])
        match = re.search(TRAILER_REGEX, raw_trailer, re.IGNORECASE)
        trailer = clean_val(match.group(0)) if match else clean_val(raw_trailer)
        if not trailer:
            continue

        in_email = trailer in cldn_units
        sailing_status = (
            "Matched 🟢" if in_email else "Not Present (Left Behind) 🔴"
        )

        raw_instructions = value_from(row, columns["instructions"])
        instr_match = re.search(TRAILER_REGEX, raw_instructions, re.IGNORECASE)
        instr_trailer = clean_val(instr_match.group(0)) if instr_match else ""
        if instr_trailer and instr_trailer != trailer:
            instructions_status = (
                f"⚠️ Mismatch: {instr_trailer} (Expected {trailer})"
            )
        else:
            instructions_status = "Clean 🟢"

        row_tokens = trailer_row_data.get(trailer, set())
        results.append(
            {
                "Trailer ID": trailer,
                "Sailing Status": sailing_status,
                "Instructions Check": instructions_status,
                "GMR Ref": verify_ref(value_from(row, columns["gmr"]), row_tokens),
                "PBN Ref": verify_ref(value_from(row, columns["pbn"]), row_tokens),
                "Booking Ref": verify_ref(value_from(row, columns["booking"]), row_tokens),
            }
        )

    qargo_units = {item["Trailer ID"] for item in results}
    for trailer in sorted(cldn_units - qargo_units):
        results.append(
            {
                "Trailer ID": trailer,
                "Sailing Status": "Forward Shipped 🔵",
                "Instructions Check": "-",
                "GMR Ref": "-",
                "PBN Ref": "-",
                "Booking Ref": "-",
            }
        )

    return pd.DataFrame(results), len(cldn_units)


def render_results(results_df, cldn_count):
    st.subheader("📋 Reconciliation Results")
    st.dataframe(results_df, use_container_width=True, hide_index=True)

    matched = int(results_df["Sailing Status"].str.contains("Matched", na=False).sum()) if not results_df.empty else 0
    left = int(results_df["Sailing Status"].str.contains("Left Behind", na=False).sum()) if not results_df.empty else 0
    forward = int(results_df["Sailing Status"].str.contains("Forward Shipped", na=False).sum()) if not results_df.empty else 0
    flagged = 0
    if not results_df.empty:
        for column in ["Instructions Check", "GMR Ref", "PBN Ref", "Booking Ref"]:
            flagged += int(results_df[column].astype(str).str.contains("⚠️|🔴", regex=True, na=False).sum())

    st.divider()
    st.subheader("📊 Summary Metrics")
    st.metric("🚢 Total Trailers on Sailing Manifest", cldn_count)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Matched Trailers", matched)
    c2.metric("Forward Shipped", forward)
    c3.metric("Left Behind", left)
    c4.metric("Flagged / Swapped Checks", flagged)


# ============================================================
# MAIN APP — UPLOAD, AUTO-MATCH, RECONCILE
# ============================================================

uploaded_file = st.file_uploader(
    "Upload Qargo Export (.csv, .xlsx or .xls)",
    type=["csv", "xlsx", "xls"],
)

if uploaded_file:
    if st.button("🚀 Go", type="primary", use_container_width=True):
        with st.spinner("Reading Qargo export and finding the matching sailing..."):
            try:
                qargo_df = read_qargo_file(uploaded_file)
                if qargo_df.empty:
                    st.error("The uploaded Qargo file contains no rows.")
                    st.stop()

                qargo_df.columns = [str(col).strip() for col in qargo_df.columns]
                columns = identify_qargo_fields(qargo_df)
                sailing_datetime, direction_prefix = infer_qargo_sailing(qargo_df, columns)
                wanted_route = DIRECTION_MAP[direction_prefix]

                st.write(f"**Qargo Intermodal Departure:** {sailing_datetime.strftime('%A %d %B %Y at %H:%M %Z')}")
                st.write(f"**Direction inferred from booking refs:** {direction_text(direction_prefix)}")

                emails = fetch_sailing_emails_for_date(sailing_datetime)
                if not emails:
                    st.error(
                        "No readable emails with subject 'Sailing Confirmation' "
                        f"were found in Gmail label '{GMAIL_LABEL}'."
                    )
                    st.stop()

                matches = []
                parsed_emails = []
                for item in emails:
                    details = parse_sailing_details(item)
                    parsed_emails.append((item, details))
                    if details["scheduled"] is None or details["route"] is None:
                        continue
                    if details["scheduled"].astimezone(DUBLIN_TZ) != sailing_datetime:
                        continue
                    if details["route"] != wanted_route:
                        continue
                    matches.append((item, details))

                # If both email types confirm this exact sailing, use the
                # standard confirmation in preference to its [Bookings] copy.
                # Do not let an unrelated standard confirmation suppress a
                # matching [Bookings] email.
                standard_matches = [
                    pair for pair in matches if not pair[0].get("is_bookings", False)
                ]
                if standard_matches:
                    matches = standard_matches

                if not matches:
                    st.error(
                        "No Sailing Confirmation matched the exact Qargo departure "
                        f"scheduled date/time ({sailing_datetime.strftime('%d/%m/%Y %H:%M')}) and route "
                        f"({direction_text(direction_prefix)}). No reconciliation "
                        "was run. Check the Qargo date, booking-reference prefix, "
                        "Gmail label and confirmation emails."
                    )
                    with st.expander("Show recent confirmations checked"):
                        rows = []
                        for item, details in parsed_emails:
                            scheduled = details["scheduled"]
                            rows.append({
                                "Subject": item["subject"],
                                "Email date": item["date_header"],
                                "Route": (
                                    f"{details['route'][0].title()} → {details['route'][1].title()}"
                                    if details["route"] else "Route not detected"
                                ),
                                "Scheduled sailing": scheduled.strftime("%d/%m/%Y %H:%M") if scheduled else "Date/time not detected",
                            })
                        if rows:
                            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                    st.stop()

                if len(matches) > 1:
                    st.error(
                        f"{len(matches)} Sailing Confirmations matched the exact scheduled date/time and "
                        "route. To avoid choosing the wrong sailing, no reconciliation "
                        "was run."
                    )
                    st.dataframe(
                        pd.DataFrame([
                            {
                                "Subject": item["subject"],
                                "Scheduled sailing": details["scheduled"].strftime("%A %d %B %Y at %H:%M"),
                                "Vessel": details["vessel"],
                                "Email date": item["date_header"],
                            }
                            for item, details in matches
                        ]),
                        use_container_width=True,
                        hide_index=True,
                    )
                    st.stop()

                selected_email, details = matches[0]
                st.success("Matching sailing identified.")
                st.markdown("### 🚢 Selected Sailing Confirmation")
                st.write(f"**Subject:** {selected_email['subject']}")
                st.write(f"**Route:** {direction_text(direction_prefix)}")
                st.write(f"**Scheduled sailing:** {details['scheduled'].strftime('%A %d %B %Y at %H:%M')}")
                if details["vessel"]:
                    st.write(f"**Vessel:** {details['vessel']}")
                st.caption(
                    "The scheduled date/time in the confirmation was used for matching. "
                    "The separate 'Departed' time is intentionally ignored."
                )

                results_df, cldn_count = reconcile(
                    qargo_df, selected_email["body"], columns
                )
                if results_df.empty:
                    st.error(
                        "The matching email was found, but no trailer IDs could be "
                        "read from the Qargo file. Please check the trailer column."
                    )
                    st.stop()
                render_results(results_df, cldn_count)

                st.download_button(
                    "⬇️ Download Reconciliation Results (CSV)",
                    data=results_df.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"CLDN_Reconciliation_{sailing_datetime:%Y%m%d_%H%M}_{direction_prefix}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            except Exception as exc:
                st.error(f"Reconciliation stopped: {type(exc).__name__}: {exc}")
                st.info(
                    "No sailing was guessed. Check the Qargo column headings, "
                    "Gmail credentials/secrets, and the matching Sailing Confirmation."
                )
else:
    st.info("Upload a Qargo export to begin. The app will identify the sailing automatically.")
