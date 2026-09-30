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

                # Extract trailer IDs present anywhere in the email
                cldn_raw = re.findall(r"GTC[\s-]?\d+", email_body, re.IGNORECASE)
                cldn_units = set(re.sub(r"[\s-]", "", t).upper() for t in cldn_raw)

                # Map each trailer to its specific row/line in the CLdN email
                email_lines = email_body.splitlines()
                trailer_email_row_map = {}
                for line in email_lines:
                    matches = re.findall(r"GTC[\s-]?\d+", line, re.IGNORECASE)
                    for m in matches:
                        clean_m = re.sub(r"[\s-]", "", m).upper()
                        trailer_email_row_map[clean_m] = line.lower()

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

                    # Get specific email row text for this trailer
                    specific_email_row = trailer_email_row_map.get(clean_trailer, "")

                    # 2. Smart Instructions Check
                    raw_instr = str(row[instr_col]).strip() if instr_col and pd.notna(row[instr_col]) else ""
                    instr_trailer_match = re.search(r"GTC[\s-]?\d+", raw_instr, re.IGNORECASE)
                    instr_trailer_clean = re.sub(r"[\s-]", "", instr_trailer_match.group(0)).upper() if instr_trailer_match else ""

                    clean_raw_text = raw_instr if raw_instr.lower() != "nan" else ""
                    extra_notes = re.sub(r"GTC[\s-]?\d+", "", clean_raw_text, flags=re.IGNORECASE).strip() if clean_raw_text else ""

                    if not clean_raw_text:
                        instr_status = "Clean (Blank) 🟢"
                    elif instr_trailer_clean and instr_trailer_clean != clean_trailer:
                        instr_status = f"⚠️ Mismatch: {clean_raw_text} (Expected {clean_trailer})"
                    elif extra_notes:
                        instr_status = f"⚠️ Flagged Note: {clean_raw_text}"
                    else:
                        instr_status = "Clean 🟢"

                    # 3. Cross-reference GMR (Row Specific)
                    raw_gmr = str(row[gmr_col]).strip() if gmr_col and pd.notna(row[gmr_col]) else ""
                    gmr_status = "-"
                    if raw_gmr and raw_gmr.lower() != "nan":
                        if specific_email_row:
                            gmr_in_row = raw_gmr.lower() in specific_email_row
                            gmr_status = f"{raw_gmr} " + ("🟢 Matched" if gmr_in_row else "⚠️ Swapped / Mismatch")
                        else:
                            gmr_status = f"{raw_gmr} 🔴 Line Missing"

                    # 4. Cross-reference PBN (Row Specific)
                    raw_pbn = str(row[pbn_col]).strip() if pbn_col and pd.notna(row[pbn_col]) else ""
                    pbn_status = "-"
                    if raw_pbn and raw_pbn.lower() != "nan":
                        if specific_email_row:
                            pbn_in_row = raw_pbn.lower() in specific_email_row
                            pbn_status = f"{raw_pbn} " + ("🟢 Matched" if pbn_in_row else "⚠️ Swapped / Mismatch")
                        else:
                            pbn_status = f"{raw_pbn} 🔴 Line Missing"

                    # 5. Cross-reference Booking Ref (Row Specific)
                    raw_booking = str(row[booking_col]).strip() if booking_col and pd.notna(row[booking_col]) else ""
                    booking_status = "-"
                    if raw_booking and raw_booking.lower() != "nan":
                        if specific_email_row:
                            booking_in_row = raw_booking.lower() in specific_email_row
                            booking_status = f"{raw_booking} " + ("🟢 Matched" if booking_in_row else "⚠️ Swapped / Mismatch")
                        else:
                            booking_status = f"{raw_booking} 🔴 Line Missing"

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
                flagged_cnt = sum(1 for r in results_data if "⚠️" in r["Instructions Check"] or "⚠️" in r["PBN Ref"] or "⚠️" in r["GMR Ref"])

                c1.metric("Matched Trailers", matched_cnt)
                c2.metric("Left Behind", left_cnt)
                c3.metric("Forward Shipped", forward_cnt)
                c4.metric("Flagged / Swapped", flagged_cnt)
