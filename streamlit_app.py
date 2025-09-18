"""
Aiclex — PDF <> Excel Matcher with robust email-column detection and emailing
Save as: streamlit_aiclex_full.py
Run: streamlit run streamlit_aiclex_full.py
"""

import os
import io
import re
import zipfile
import tempfile
from datetime import datetime
from typing import List, Dict, Optional

import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
import smtplib
from email.message import EmailMessage

# ---------------- Config / Branding ----------------
APP_TITLE = "Aiclex — PDF ≠ Excel Matcher (Email Ready)"
BRAND_COLOR = "#0b74de"
FOOTER_TEXT = "App by Aiclex Technologies — Faridabad | Contact: info@aiclex.in"

# Patterns
HALLTICKET_PATTERN = re.compile(r"\b[0-9]{4,}\b")
ABSENT_PATTERN = re.compile(r"\b(absent|a b s e n t|not present)\b", re.IGNORECASE)

st.set_page_config(page_title=APP_TITLE, layout="wide")

# Sidebar: SMTP & Template
st.sidebar.header("SMTP / Email settings (optional)")
smtp_host = st.sidebar.text_input("SMTP host", value=os.environ.get("SMTP_HOST", "smtp.gmail.com"))
smtp_port = st.sidebar.number_input("SMTP port", min_value=1, max_value=65535, value=int(os.environ.get("SMTP_PORT", 587)))
smtp_use_tls = st.sidebar.checkbox("Use STARTTLS", value=True)
smtp_user = st.sidebar.text_input("SMTP username (email)", value=os.environ.get("SMTP_USER", ""))
smtp_pass = st.sidebar.text_input("SMTP password (or app password)", value=os.environ.get("SMTP_PASS", ""), type="password")
from_email = st.sidebar.text_input("From email (optional)", value=smtp_user or os.environ.get("FROM_EMAIL", ""))
test_target = st.sidebar.text_input("Test email to (try before sending)", value="")

st.sidebar.markdown("---")
st.sidebar.header("Email Template")
default_subject = "Exam Result: {Employee Name} — {status}"
default_body = (
    "Dear {Employee Name},\n\n"
    "Your Hallticket: {Hallticket}\n"
    "Marks: {marks}\n"
    "Status: {status}\n\n"
    "Location: {Location}\n\n"
    "Regards,\nAiclex Technologies"
)
email_subject_tpl = st.sidebar.text_input("Email subject template", value=default_subject)
email_body_tpl = st.sidebar.text_area("Email body template", value=default_body, height=200)
st.sidebar.markdown("Placeholders: {Employee Name}, {Hallticket}, {marks}, {status}, {Location}")

# Header
st.markdown(
    f"<div style='display:flex; align-items:center; gap:16px;'>"
    f"<div><h1 style='color:{BRAND_COLOR}; margin:0;'>{APP_TITLE}</h1>"
    f"<div style='color:gray'>Developed by Aiclex Technologies</div></div>"
    f"</div>",
    unsafe_allow_html=True,
)
st.markdown("---")

# ----------------- PDF / marks extraction helpers -----------------
def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    text_parts: List[str] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                if txt.strip():
                    text_parts.append(txt)
    except Exception:
        pass
    combined = "\n".join(text_parts).strip()
    if combined:
        return combined
    # OCR fallback
    ocr_texts: List[str] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                try:
                    pil_image = page.to_image(resolution=150).original
                    ocr_text = pytesseract.image_to_string(pil_image)
                    ocr_texts.append(ocr_text)
                except Exception:
                    continue
    except Exception:
        pass
    return "\n".join(ocr_texts)

def better_find_marks_and_status(text: Optional[str]) -> (Optional[int], Optional[str]):
    if not text:
        return None, None
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    # 1) lines with PASS/FAIL/ABSENT
    for ln in lines:
        mstatus = re.search(r"\b(PASS|FAIL|ABSENT)\b", ln, re.IGNORECASE)
        if mstatus:
            nums = re.findall(r"\b(\d{1,3})\b", ln)
            if nums:
                status_word = mstatus.group(1).upper()
                status_pos = mstatus.start()
                # choose number before status if exists
                best = None
                best_dist = None
                for n in nums:
                    pos = ln.find(n)
                    if pos != -1 and pos < status_pos:
                        dist = status_pos - pos
                        if best is None or dist < best_dist:
                            best = int(n)
                            best_dist = dist
                if best is not None:
                    return best, status_word
                return int(nums[-1]), status_word
    # 2) look for 'marks' header and following lines
    for i, ln in enumerate(lines):
        if "marks obtained" in ln.lower() or re.search(r"\bmarks\b", ln, re.IGNORECASE):
            for j in range(i+1, min(i+4, len(lines))):
                nums = re.findall(r"\b(\d{1,3})\b", lines[j])
                if nums:
                    return int(nums[-1]), None
    # 3) fallback: filter out 'minimum' occurrences and return last good number
    all_nums = re.findall(r"\b(\d{1,3})\b", text)
    all_nums = [int(n) for n in all_nums if 0 <= int(n) <= 100]
    filtered = []
    for n in all_nums:
        idx = text.find(str(n))
        if idx == -1:
            continue
        window = text[max(0, idx-40): idx+40].lower()
        if "minimum" in window or "passing" in window:
            continue
        filtered.append(n)
    if filtered:
        return filtered[-1], None
    return None, None

def find_hallticket_in_text(text: Optional[str]) -> Optional[str]:
    if not text:
        return None
    m = re.search(r"(?:Hall ?Ticket|Hallticket|Registration No|Reg No|Roll No|RollNumber|Roll)[:\s]*([0-9]{4,})", text, re.IGNORECASE)
    if m:
        return m.group(1)
    candidates = HALLTICKET_PATTERN.findall(text or "")
    return max(candidates, key=len) if candidates else None

# ----------------- ZIP processing -----------------
def process_uploaded_zip_bytes(zip_bytes: bytes) -> List[Dict]:
    results: List[Dict] = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as outer:
            for name in outer.namelist():
                if name.endswith("/"):
                    continue
                data = outer.read(name)
                if name.lower().endswith(".zip"):
                    nested = process_uploaded_zip_bytes(data)
                    results.extend(nested)
                elif name.lower().endswith(".pdf"):
                    text = extract_text_from_pdf_bytes(data)
                    hall = find_hallticket_in_text(text) or ""
                    marks, status = better_find_marks_and_status(text)
                    results.append({"hallticket": str(hall).strip(), "marks": marks, "status_extracted": status, "source_pdf": name})
                else:
                    continue
    except zipfile.BadZipFile:
        pass
    return results

# ----------------- Excel fill logic -----------------
def fill_excel_with_pdf_data(df: pd.DataFrame, pdf_data: List[Dict]):
    # detect hallticket column
    hall_col = None
    for c in df.columns:
        lc = c.lower()
        if "hall" in lc or "ticket" in lc or "roll" in lc:
            hall_col = c
            break
    if not hall_col:
        st.warning("No Hallticket column found in Excel — please ensure column name contains 'hall' or 'ticket'.")
        return df, 0, []

    # marks/status columns
    marks_col = None
    status_col = None
    for c in df.columns:
        lc = c.lower()
        if lc == "marks" or "mark" in lc:
            marks_col = c
        if lc == "status":
            status_col = c
    if not marks_col:
        df["marks"] = None
        marks_col = "marks"
    if not status_col:
        df["status"] = None
        status_col = "status"

    lookup = { str(item.get("hallticket") or "").strip(): item for item in pdf_data }

    filled = 0
    unmatched = []
    for idx, row in df.iterrows():
        cell = str(row.get(hall_col, "")).strip()
        if not cell:
            unmatched.append({"index": idx, "reason": "missing hallticket"})
            continue
        item = None
        if cell in lookup:
            item = lookup[cell]
        else:
            digits = re.sub(r"\D", "", cell)
            if digits and digits in lookup:
                item = lookup[digits]
            else:
                found = None
                for k in lookup.keys():
                    if digits and (k.endswith(digits) or digits.endswith(k)):
                        found = k
                        break
                if found:
                    item = lookup[found]
        if item:
            marks = item.get("marks")
            status_ex = item.get("status_extracted")
            if isinstance(marks, str) and str(marks).lower().startswith("abs"):
                final_status = "Absent"
            else:
                try:
                    mm = int(marks) if marks is not None else None
                    final_status = "Pass" if (mm is not None and mm > 49) else ("Fail" if mm is not None else (status_ex.title() if status_ex else "Unknown"))
                except Exception:
                    final_status = status_ex.title() if status_ex else "Unknown"
            df.at[idx, marks_col] = marks
            df.at[idx, status_col] = final_status
            filled += 1
        else:
            unmatched.append({"index": idx, "hallticket": cell})
    return df, filled, unmatched

# ----------------- Email helpers -----------------
def render_template(tpl: str, row: pd.Series) -> str:
    mapping = {
        "Employee Name": str(row.get("Employee Name") or row.get("Name") or ""),
        "Hallticket": str(row.get("Hallticket") or row.get("hallticket") or ""),
        "marks": str(row.get("marks") or row.get("Marks") or ""),
        "status": str(row.get("status") or ""),
        "Location": str(row.get("Location") or "")
    }
    result = tpl
    for k, v in mapping.items():
        result = result.replace("{" + k + "}", v)
    return result

def send_emails_for_dataframe(df: pd.DataFrame, email_col_name: str, subject_tpl: str, body_tpl: str,
                              smtp_host: str, smtp_port: int, use_tls: bool, user: str, pwd: str, from_addr: str):
    sent = 0
    fails = []
    total = len(df)
    progress = st.progress(0)
    status = st.empty()

    if not from_addr:
        from_addr = user

    try:
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=20)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port, timeout=20)
        if use_tls and smtp_port != 465:
            server.starttls()
        if user and pwd:
            server.login(user, pwd)
    except Exception as e:
        st.error(f"Failed to connect to SMTP server: {e}")
        return 0, [{"error": str(e)}]

    try:
        for i, (idx, row) in enumerate(df.iterrows(), start=1):
            try:
                to_email = str(row.get(email_col_name) or "").strip()
                if not to_email:
                    fails.append({"index": idx, "reason": "no recipient email", "row_index": idx})
                    continue
                subj = render_template(subject_tpl, row)
                body = render_template(body_tpl, row)
                msg = EmailMessage()
                msg["From"] = from_addr
                msg["To"] = to_email
                msg["Subject"] = subj
                msg.set_content(body)
                server.send_message(msg)
                sent += 1
            except Exception as ex:
                fails.append({"index": idx, "error": str(ex), "to": to_email})
            progress.progress(i / total)
            status.text(f"Sent: {sent} | Failed: {len(fails)} | Processing row {i}/{total}")
    finally:
        try:
            server.quit()
        except Exception:
            pass
    return sent, fails

# ----------------- UI: Upload / Process -----------------
st.info("Upload Excel/CSV + ZIP (nested ZIPs with PDFs). After processing you can send results by email.")

col1, col2 = st.columns([2, 3])
with col1:
    uploaded_excel = st.file_uploader("Excel/CSV file", type=["xlsx", "xls", "csv"], key="excel_up")
with col2:
    uploaded_zip = st.file_uploader("ZIP file (nested zips with PDFs)", type=["zip"], key="zip_up")

process_btn = st.button("Process and Match")

processed_df = None
pdf_data = None
filled_count = 0
unmatched_rows = []

if process_btn:
    if not uploaded_excel or not uploaded_zip:
        st.error("Please upload both Excel/CSV and ZIP files.")
    else:
        try:
            if uploaded_excel.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_excel, dtype=str)
            else:
                df = pd.read_excel(uploaded_excel, dtype=str, engine="openpyxl")
        except Exception as e:
            st.error(f"Failed to read Excel/CSV: {e}")
            st.stop()
        st.success(f"Excel loaded — {len(df)} rows")

        with st.spinner("Processing ZIP and PDFs (may use OCR)..."):
            zip_bytes = uploaded_zip.read()
            pdf_data = process_uploaded_zip_bytes(zip_bytes)

        st.info(f"PDFs processed: {len(pdf_data)}")
        updated_df, filled_count, unmatched_rows = fill_excel_with_pdf_data(df, pdf_data)
        processed_df = updated_df

        st.success(f"Filled {filled_count} rows from PDFs")
        st.markdown("### Preview (first 30 rows)")
        st.dataframe(updated_df.head(30))

        out_buf = io.BytesIO()
        try:
            updated_df.to_excel(out_buf, index=False, engine="openpyxl")
            out_buf.seek(0)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button("Download updated Excel", data=out_buf.read(), file_name=f"updated_aiclex_{ts}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Failed to prepare download: {e}")

        if pdf_data:
            st.markdown("### Sample extracted PDF data")
            st.dataframe(pd.DataFrame(pdf_data).head(30))
        if unmatched_rows:
            st.markdown("### Sample unmatched rows")
            st.dataframe(pd.DataFrame(unmatched_rows).head(30))

# ----------------- Email UI -----------------
st.markdown("---")
st.header("Email: send results to addresses in Excel")

if processed_df is None:
    st.info("Process an Excel + ZIP first to enable emailing.")
else:
    # detect email-like columns
    def detect_email_columns(columns: List[str]) -> List[str]:
        email_candidates = []
        patterns = ["email", "e-mail", "email id", "email_id", "emailaddress", "mail", "contact email", "contact_email"]
        for c in columns:
            lc = c.lower().strip()
            if any(p in lc for p in patterns) or re.search(r"^\S+@\S+\.\S+$", lc):
                email_candidates.append(c)
        # also include exact matches
        for c in columns:
            if c not in email_candidates and "@" in str(processed_df.get(c).astype(str).head(20).to_list()):
                email_candidates.append(c)
        return list(dict.fromkeys(email_candidates))  # unique preserve order

    cols = list(processed_df.columns)
    email_cols = detect_email_columns(cols)

    st.write("Detected columns:", cols)
    if email_cols:
        st.success(f"Auto-detected email-like columns: {email_cols}")
    else:
        st.warning("No email-like column auto-detected. You can pick manually or type the column name below.")

    # allow user to choose from detected or any column or type manually
    choice_method = st.radio("Choose email column:", ("Pick from detected", "Pick from all columns", "Type column name manually"))
    chosen_email_col = None
    if choice_method == "Pick from detected":
        if email_cols:
            chosen_email_col = st.selectbox("Select email column", email_cols)
        else:
            st.info("No detected columns available.")
    elif choice_method == "Pick from all columns":
        chosen_email_col = st.selectbox("Select email column", cols)
    else:
        manually = st.text_input("Type exact column name (case-sensitive)", value="")
        chosen_email_col = manually.strip() if manually.strip() else None

    st.markdown("**Preview recipients (first 10 rows)**")
    if chosen_email_col and chosen_email_col in processed_df.columns:
        st.dataframe(processed_df[[chosen_email_col, *([c for c in ['Employee Name','Hallticket','marks','status','Location'] if c in processed_df.columns])]].head(10))
    else:
        st.info("No valid email column selected yet.")

    # Test email
    st.markdown("### Test SMTP / Test Email")
    if st.button("Send test email (to sidebar Test email address)"):
        if not smtp_host or not smtp_user or not smtp_pass:
            st.error("Please configure SMTP host, username and password in the sidebar.")
        elif not test_target:
            st.error("Please provide a test target address in the sidebar 'Test email to'.")
        else:
            try:
                test_df = pd.DataFrame([{chosen_email_col: test_target, "Employee Name": "Test User", "Hallticket": "0000", "marks": "NA", "status": "Test"}])
                sent_count, fails = send_emails_for_dataframe(
                    test_df,
                    email_col_name=chosen_email_col,
                    subject_tpl=email_subject_tpl,
                    body_tpl=email_body_tpl,
                    smtp_host=smtp_host,
                    smtp_port=smtp_port,
                    use_tls=smtp_use_tls,
                    user=smtp_user,
                    pwd=smtp_pass,
                    from_addr=from_email
                )
                if sent_count > 0:
                    st.success("Test email sent successfully.")
                else:
                    st.error(f"Test failed: {fails}")
            except Exception as e:
                st.error(f"Test send failed: {e}")

    # Send real emails
    if st.button("Send results emails now"):
        if not chosen_email_col or chosen_email_col not in processed_df.columns:
            st.error("Please select or type a valid email column first.")
        elif not smtp_host or not smtp_user or not smtp_pass:
            st.error("Please configure SMTP host, username and password in the sidebar.")
        else:
            # choose scope
            scope = st.selectbox("Which recipients?", ("All rows", "Only rows with filled marks/status", "Only unmatched rows (report)"))
            if scope == "All rows":
                target_df = processed_df.copy()
            elif scope == "Only rows with filled marks/status":
                if "marks" in processed_df.columns:
                    target_df = processed_df[processed_df["marks"].notnull()].copy()
                else:
                    target_df = processed_df.copy()
            else:
                idxs = [u.get("index") for u in unmatched_rows] if unmatched_rows else []
                target_df = processed_df.loc[idxs].copy() if idxs else pd.DataFrame()
            if target_df.empty:
                st.warning("No recipients in selected scope.")
            else:
                st.info(f"Sending to {len(target_df)} recipients...")
                sent_count, fails = send_emails_for_dataframe(
                    target_df,
                    email_col_name=chosen_email_col,
                    subject_tpl=email_subject_tpl,
                    body_tpl=email_body_tpl,
                    smtp_host=smtp_host,
                    smtp_port=smtp_port,
                    use_tls=smtp_use_tls,
                    user=smtp_user,
                    pwd=smtp_pass,
                    from_addr=from_email
                )
                st.success(f"Email sending finished. Sent: {sent_count} | Failed: {len(fails)}")
                if fails:
                    st.markdown("### Failures (sample)")
                    st.dataframe(pd.DataFrame(fails).head(50))

st.markdown("---")
st.markdown(f"<div style='color:gray; font-size:12px'>{FOOTER_TEXT}</div>", unsafe_allow_html=True)
