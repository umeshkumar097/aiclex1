# streamlit_app.py
"""
Aiclex — Result Showing (final)
Features:
- Upload Excel/CSV and nested ZIP of PDFs
- OCR each PDF and extract Marks following the label "Marks Obtained"
- If Marks numeric -> Pass if >49 else Fail. If label present but no numeric -> Absent.
- Fill Excel marks/status columns (prefer existing Excel marks if present)
- Produce Summary sheet (overall + per-location)
- Group PDFs per recipient email, then group by location, create ZIPs, split ZIPs > 3MB
- Preview ZIP parts and send via Gmail SMTP (smtp.gmail.com TLS) with test-mode option
- Progress bar & logs in Streamlit UI
"""

import os
import io
import re
import time
import zipfile
import tempfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from email.message import EmailMessage

import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import pytesseract

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE = True
except Exception:
    PDF2IMAGE = False

# ========== Configuration ==========
DEFAULT_ATTACHMENT_MB = 3.0
DEFAULT_DPI = 200
# streamlit page
st.set_page_config(page_title="Aiclex — Result Showing", layout="wide")
st.title("📊 Aiclex — Result Showing")

# ========== Sidebar: settings ==========
st.sidebar.header("OCR / Sending Settings")
tesseract_path = st.sidebar.text_input("Tesseract path (optional)", value=os.environ.get("TESSERACT_CMD", ""))
ocr_lang = st.sidebar.text_input("OCR language for Tesseract (e.g. 'eng' or 'eng+hin')", value="eng")
ocr_dpi = st.sidebar.number_input("OCR DPI (pdf2image)", min_value=100, max_value=400, value=DEFAULT_DPI)
show_ocr_debug = st.sidebar.checkbox("Show OCR debug snippets", value=False)

st.sidebar.markdown("---")
st.sidebar.header("Email / Attachment Settings")
attachment_limit_mb = st.sidebar.number_input("Attachment limit (MB)", value=DEFAULT_ATTACHMENT_MB, step=0.5)
send_delay = st.sidebar.number_input("Delay between sends (seconds)", value=1.0, step=0.5)
st.sidebar.markdown("SMTP: smtp.gmail.com (TLS port 587) recommended for Gmail (use App Password).")

if tesseract_path:
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

# ========== Utility helpers ==========
def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ("B","KB","MB","GB"):
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

LABEL_RE = re.compile(r"Marks\s*Obtained\s*[:\-]?\s*(?:\n\s*)?([0-9]{1,3})", re.IGNORECASE)
ABSENT_RE = re.compile(r"\b(absent|not present|a\s*b\s*s\s*e\s*n\s*t)\b", re.IGNORECASE)
HALL_RE = re.compile(r"\b[0-9]{3,}\b")  # hallticket numeric candidates
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# ========== OCR / PDF extraction ==========
def extract_text_from_pdf_bytes(pdf_bytes, dpi=200, lang="eng"):
    """Try pdfplumber text extraction first; if empty, use pdf2image -> pytesseract or page.to_image() fallback."""
    # 1) pdfplumber text
    texts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for p in pdf.pages:
                try:
                    t = p.extract_text() or ""
                except Exception:
                    t = ""
                if t and t.strip():
                    texts.append(t)
    except Exception:
        pass
    combined = "\n".join(texts).strip()
    if combined:
        return combined

    # 2) pdf2image -> pytesseract
    if PDF2IMAGE:
        try:
            images = convert_from_bytes(pdf_bytes, dpi=dpi)
            ocr_texts = []
            for im in images:
                try:
                    txt = pytesseract.image_to_string(im, lang=lang)
                except Exception:
                    txt = pytesseract.image_to_string(im)
                ocr_texts.append(txt or "")
            final = "\n".join(ocr_texts).strip()
            if final:
                return final
        except Exception:
            pass

    # 3) pdfplumber page.to_image() -> pytesseract
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            ocr_texts = []
            for p in pdf.pages:
                try:
                    pil = p.to_image(resolution=dpi).original
                    try:
                        txt = pytesseract.image_to_string(pil, lang=lang)
                    except Exception:
                        txt = pytesseract.image_to_string(pil)
                    ocr_texts.append(txt or "")
                except Exception:
                    continue
            final = "\n".join(ocr_texts).strip()
            if final:
                return final
    except Exception:
        pass

    # 4) PIL fallback
    try:
        im = Image.open(io.BytesIO(pdf_bytes))
        try:
            txt = pytesseract.image_to_string(im, lang=lang)
        except Exception:
            txt = pytesseract.image_to_string(im)
        return (txt or "").strip()
    except Exception:
        return ""

def extract_from_zip_recursive(zip_bytes, ocr_dpi, ocr_lang, show_debug=False):
    """Return list of dicts: {pdf_name, pdf_bytes, hallticket (if found), marks (int|'Absent'|None), text_snippet}"""
    out = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith("/"):
                    continue
                data = zf.read(name)
                lname = name.lower()
                if lname.endswith(".zip"):
                    out.extend(extract_from_zip_recursive(data, ocr_dpi, ocr_lang, show_debug))
                elif lname.endswith(".pdf"):
                    text = extract_text_from_pdf_bytes(data, dpi=ocr_dpi, lang=ocr_lang)
                    # find "Marks Obtained" label number (first attempt)
                    m = LABEL_RE.search(text or "")
                    if m:
                        try:
                            val = int(m.group(1))
                        except:
                            val = None
                    else:
                        val = None
                    # if label present but no numeric -> treat as Absent
                    if m and val is None:
                        marks = "Absent"
                    else:
                        # if no label found, check absent words
                        if val is None and ABSENT_RE.search((text or "").lower()):
                            marks = "Absent"
                        elif val is None:
                            # per your request: if Marks Obtained label not present / no number -> Absent
                            marks = None
                        else:
                            marks = val
                    # find possible hallticket in PDF text or fallback to filename digits
                    hall = None
                    h_cands = HALL_RE.findall(text or "")
                    if h_cands:
                        hall = max(h_cands, key=len)
                    else:
                        fn_digits = re.findall(r"\d+", os.path.basename(name))
                        hall = fn_digits[-1] if fn_digits else ""
                    out.append({
                        "pdf_name": os.path.basename(name),
                        "pdf_bytes": data,
                        "hallticket": str(hall) if hall else "",
                        "marks": marks,   # int | "Absent" | None
                        "text_snippet": (text or "")[:2000] if show_debug else ""
                    })
                else:
                    # ignore other file types
                    continue
    except zipfile.BadZipFile:
        pass
    return out

# ========== Excel fill & summary ==========
def fill_excel_with_pdf_data(df, pdf_entries, hall_col):
    """
    df: pandas DataFrame (string columns)
    pdf_entries: list of dicts from extract_from_zip_recursive
    hall_col: column name string in df that has hallticket
    Rules:
      - Prefer existing Excel marks (if non-empty)
      - Else use PDF 'marks' if int
      - If PDF 'marks' is "Absent" or Excel marks blank and PDF has no numeric -> Absent
      - Pass if numeric >49, else Fail (49 => Fail)
    Returns: updated_df, filled_count, unmatched_rows
    """
    # map pdf hallticket -> list of marks (choose first numeric if multiple)
    pdf_map = {}
    for p in pdf_entries:
        key = str(p.get("hallticket") or "").strip()
        if not key:
            continue
        # prefer numeric marks; if multiple PDFs for same hallticket keep first with numeric else first Absent
        current = pdf_map.get(key)
        if current is None:
            pdf_map[key] = p.get("marks")
        else:
            # if current None or 'Absent' and p has numeric, prefer numeric
            if (current in (None, "Absent")) and isinstance(p.get("marks"), int):
                pdf_map[key] = p.get("marks")

    # ensure columns
    marks_col = "marks"
    status_col = "status"
    if marks_col not in df.columns:
        df[marks_col] = ""
    if status_col not in df.columns:
        df[status_col] = ""

    filled = 0
    unmatched = []

    for idx, row in df.iterrows():
        ht_val = str(row.get(hall_col, "")).strip()
        if not ht_val:
            unmatched.append({"index": idx, "reason": "no_hallticket"})
            continue

        # existing excel mark?
        excel_raw = row.get(marks_col, "")
        excel_mark = None
        if pd.notna(excel_raw) and str(excel_raw).strip() != "":
            # try parse int; if 'Absent' string treat accordingly
            txt = str(excel_raw).strip()
            if re.search(r'abs', txt, re.IGNORECASE):
                excel_mark = "Absent"
            else:
                digits = re.sub(r"\D", "", txt)
                if digits:
                    try:
                        excel_mark = int(digits)
                    except:
                        excel_mark = None

        # pdf_map match attempts
        pdf_mark = None
        if ht_val in pdf_map:
            pdf_mark = pdf_map[ht_val]
        else:
            digits = re.sub(r"\D", "", ht_val)
            if digits and digits in pdf_map:
                pdf_mark = pdf_map[digits]
            else:
                # try endswith / contains
                for k in pdf_map.keys():
                    kd = re.sub(r"\D","", str(k))
                    if kd and (k.endswith(digits) or digits.endswith(kd) or kd.endswith(digits)):
                        pdf_mark = pdf_map[k]
                        break

        # decide final
        final = None
        if excel_mark is not None:
            final = excel_mark
        elif pdf_mark is not None:
            final = pdf_mark
        else:
            final = None

        # set marks & status
        if final is None:
            # treat as Absent per your instruction when no numeric found in 'Marks Obtained'
            df.at[idx, marks_col] = ""
            df.at[idx, status_col] = "Absent"
        else:
            if isinstance(final, str) and re.search(r'abs', final, re.IGNORECASE):
                df.at[idx, marks_col] = ""
                df.at[idx, status_col] = "Absent"
            else:
                try:
                    mm = int(final)
                    df.at[idx, marks_col] = mm
                    df.at[idx, status_col] = "Pass" if mm > 49 else "Fail"
                except:
                    df.at[idx, marks_col] = ""
                    df.at[idx, status_col] = "Absent"
        filled += 1

    return df, filled, unmatched

def make_summary_sheet(df, location_col):
    """
    Create summary dataframe(s) and return a dict of DataFrames to write as separate sheets:
    - 'results' : updated df
    - 'summary_overall' : overall counts
    - 'summary_by_location' : per-location pass/fail/absent counts
    """
    total = len(df)
    pass_count = sum(df['status'] == 'Pass')
    fail_count = sum(df['status'] == 'Fail')
    absent_count = sum(df['status'] == 'Absent')

    summary_overall = pd.DataFrame([{
        "Total": total,
        "Pass": int(pass_count),
        "Fail": int(fail_count),
        "Absent": int(absent_count)
    }])

    by_loc = []
    if location_col in df.columns:
        for loc, g in df.groupby(location_col):
            by_loc.append({
                "Location": loc,
                "Total": len(g),
                "Pass": int((g['status'] == 'Pass').sum()),
                "Fail": int((g['status'] == 'Fail').sum()),
                "Absent": int((g['status'] == 'Absent').sum())
            })
    summary_by_location = pd.DataFrame(by_loc)

    return {"results": df, "summary_overall": summary_overall, "summary_by_location": summary_by_location}

# ========== ZIP packaging & splitting ==========
def make_zip_bytes(file_entries):
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, b in file_entries:
            zf.writestr(fname, b)
    bio.seek(0)
    return bio.read()

def split_files_into_zip_parts(file_entries, max_bytes, zip_name_prefix="results"):
    parts = []
    current = []
    for fname, b in file_entries:
        test = current + [(fname, b)]
        test_zip = make_zip_bytes(test)
        if len(test_zip) <= max_bytes:
            current = test
        else:
            if current:
                parts.append((f"{zip_name_prefix}_part{len(parts)+1}.zip", make_zip_bytes(current)))
            # put current file into its own part (may exceed max_bytes)
            single = [(fname, b)]
            parts.append((f"{zip_name_prefix}_part{len(parts)+1}.zip", make_zip_bytes(single)))
            current = []
    if current:
        parts.append((f"{zip_name_prefix}_part{len(parts)+1}.zip", make_zip_bytes(current)))
    return parts

# ========== Email sending (Gmail TLS) ==========
def send_email_gmail(smtp_user, smtp_pass, to_email, subject, body, attachments):
    """
    attachments: list of tuples (filename, bytes)
    returns True/False, error_message(if any)
    """
    msg = EmailMessage()
    msg["From"] = smtp_user
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)
    for fname, b in attachments:
        msg.add_attachment(b, maintype="application", subtype="zip", filename=fname)
    try:
        import smtplib
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(smtp_user, smtp_pass)
            s.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

# ========== Streamlit UI: main flow ==========
st.header("1) Upload Excel/CSV and ZIP (nested ZIPs supported)")
uploaded_excel = st.file_uploader("Upload Excel or CSV", type=["xlsx", "csv"])
uploaded_zip = st.file_uploader("Upload ZIP (containing PDFs, nested zips ok)", type=["zip"])

if uploaded_excel and uploaded_zip:
    # read excel
    try:
        if uploaded_excel.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded_excel, dtype=str, engine="openpyxl").fillna("")
    except Exception as e:
        st.error(f"Failed to read Excel/CSV: {e}")
        st.stop()

    st.success(f"Excel loaded: {len(df)} rows")
    cols = df.columns.tolist()
    hall_col = st.selectbox("Select Hallticket column", cols)
    email_col = st.selectbox("Select Email column", cols)
    location_col = st.selectbox("Select Location column", cols)

    st.info("Processing ZIP and extracting PDFs (OCR may take time)...")
    with st.spinner("Running OCR and extracting PDF fields..."):
        pdf_entries = extract_from_zip_recursive(uploaded_zip.read(), ocr_dpi, ocr_lang, show_debug=show_ocr_debug)

    st.success(f"Extracted {len(pdf_entries)} PDF records from ZIP(s)")

    if show_ocr_debug:
        st.subheader("OCR debug (sample)")
        debug_rows = []
        for p in pdf_entries:
            debug_rows.append({
                "pdf_name": p["pdf_name"],
                "hallticket": p["hallticket"],
                "marks": p["marks"],
                "text_snippet": p["text_snippet"][:400]
            })
        st.dataframe(pd.DataFrame(debug_rows).head(200))

    # fill excel
    updated_df, filled_count, unmatched = fill_excel_with_pdf_data(df.copy(), pdf_entries, hall_col)
    st.success(f"Filled {filled_count} rows (marks/status).")
    if unmatched:
        st.warning(f"{len(unmatched)} rows had no hallticket value.")

    st.subheader("Preview updated results (first 100 rows)")
    st.dataframe(updated_df.head(100))

    # produce summary workbook in memory and offer download
    sheets = make_summary_sheet(updated_df, location_col)
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    excel_buf.seek(0)
    st.download_button("Download results + summary (Excel)", data=excel_buf,
                       file_name=f"aiclex_results_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ========== Build grouping: recipients by email and collect their PDFs location-wise ==========
    st.markdown("---")
    st.header("2) Preview grouping & prepare ZIPs")
    # Build mapping hallticket -> list of pdf entries
    pdf_map = defaultdict(list)
    for p in pdf_entries:
        key = str(p.get("hallticket") or "").strip()
        if key:
            pdf_map[key].append(p)

    # Map recipients to their halltickets and locations
    recipients = defaultdict(lambda: {"locations": defaultdict(list)})
    # track missing pdfs
    missing_pdf_log = []
    for _, row in updated_df.iterrows():
        ht = str(row.get(hall_col, "")).strip()
        emails_raw = str(row.get(email_col, "")).strip()
        loc = str(row.get(location_col, "")).strip() or "Unknown"
        if not emails_raw:
            continue
        # split on common separators
        emails = [e.strip() for e in re.split(r"[;, \n]+", emails_raw) if e.strip()]
        for em in emails:
            if not EMAIL_RE.match(em):
                continue
            # add pdfs for this hallticket
            if ht and ht in pdf_map:
                for p in pdf_map[ht]:
                    recipients[em]["locations"][loc].append(p)
            else:
                # try matching by digits
                digits = re.sub(r"\D","", ht)
                found = False
                if digits:
                    for k, lst in pdf_map.items():
                        kd = re.sub(r"\D","", str(k))
                        if kd and (kd == digits or kd.endswith(digits) or digits.endswith(kd)):
                            for p in lst: recipients[em]["locations"][loc].append(p)
                            found = True
                            break
                if not found:
                    missing_pdf_log.append({"email": em, "hallticket": ht, "location": loc})

    st.subheader("Recipient summary (sample)")
    rec_preview = []
    for em, info in list(recipients.items())[:200]:
        total_files = sum(len(v) for v in info["locations"].values())
        rec_preview.append({"email": em, "locations": ", ".join(info["locations"].keys()), "files": total_files})
    if rec_preview:
        st.dataframe(pd.DataFrame(rec_preview))
    else:
        st.info("No recipients with matched PDFs found yet. Check email column / hallticket mapping.")

    if missing_pdf_log:
        st.warning(f"Some rows had no matching PDF found: {len(missing_pdf_log)} (sample shown)")
        st.dataframe(pd.DataFrame(missing_pdf_log).head(50))

    # ========== Prepare ZIP parts in memory ==========
    if st.button("Prepare ZIPs (grouped by recipient -> location)"):
        st.info("Preparing ZIP parts in memory (may use RAM).")
        prepared = {}  # recipient -> list of (loc, [(partname, bytes), ...])
        max_bytes = int(attachment_limit_mb * 1024 * 1024)
        for em, info in recipients.items():
            prepared[em] = []
            for loc, plist in info["locations"].items():
                # create file entries as (filename_inside_zip, bytes)
                file_entries = []
                for p in plist:
                    fname = f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}"
                    file_entries.append((fname, p["pdf_bytes"]))
                parts = split_files_into_zip_parts(file_entries, max_bytes, zip_name_prefix=re.sub(r"[^A-Za-z0-9]+","_",loc)[:40])
                prepared[em].append((loc, parts))
        st.session_state["prepared"] = prepared
        st.success("Prepared ZIP parts stored in session (in-memory).")

    # Show prepared preview
    if "prepared" in st.session_state:
        preview_rows = []
        for em, locs in st.session_state["prepared"].items():
            for loc, parts in locs:
                for i, (pname, pbytes) in enumerate(parts, start=1):
                    preview_rows.append({"email": em, "location": loc, "part": i, "zip_name": pname, "size": human_bytes(len(pbytes))})
        st.subheader("Prepared ZIPs preview")
        st.dataframe(pd.DataFrame(preview_rows).head(200))

        # ========== Sending UI ==========
        st.markdown("---")
        st.header("3) Send emails")
        smtp_user = st.text_input("Gmail address (SMTP user)", value=os.environ.get("SMTP_USER",""))
        smtp_pass = st.text_input("Gmail App Password (SMTP pass)", type="password", value=os.environ.get("SMTP_PASS",""))
        test_mode = st.checkbox("Test mode (send everything to single test email)", value=True)
        test_email = st.text_input("Test email address (if test mode on)", value=os.environ.get("TEST_EMAIL",""))
        body_template = st.text_area("Email body (use {location}, {part}, {total_parts})",
                                    value="Hello,\n\nPlease find attached results for {location} (Part {part}/{total_parts}).\n\nRegards,\nAiclex")
        subject_template = st.text_input("Email subject (use {location})", value="Results for {location}")
        if st.button("Start sending prepared ZIPs"):
            if not smtp_user or not smtp_pass:
                st.error("Provide Gmail address and app password for SMTP.")
            else:
                # compute total sends for progress
                total_sends = 0
                for em, locs in st.session_state["prepared"].items():
                    recips = [test_email] if test_mode else [em]
                    for loc, parts in locs:
                        total_sends += len(recips) * max(1, len(parts))
                if total_sends == 0:
                    st.warning("No prepared zips to send.")
                else:
                    progress = st.progress(0)
                    status = st.empty()
                    sent = 0
                    failed = []
                    count = 0
                    for em, locs in st.session_state["prepared"].items():
                        recipients_list = [test_email] if test_mode else [em]
                        for loc, parts in locs:
                            total_parts = max(1, len(parts))
                            for part_index, (zipname, zipbytes) in enumerate(parts, start=1):
                                subj = subject_template.format(location=loc)
                                body = body_template.format(location=loc, part=part_index, total_parts=total_parts)
                                for r in recipients_list:
                                    ok, err = send_email_gmail(smtp_user, smtp_pass, r, subj, body, [(zipname, zipbytes)])
                                    count += 1
                                    progress.progress(min(1.0, count / total_sends))
                                    status.text(f"Sent {count}/{total_sends} → {r} ({loc} part {part_index}/{total_parts})")
                                    if ok:
                                        sent += 1
                                    else:
                                        failed.append({"recipient": r, "loc": loc, "zip": zipname, "error": err})
                                    time.sleep(send_delay)
                    st.success(f"Sending complete. Sent: {sent}. Failed: {len(failed)}")
                    if failed:
                        st.error("Some failures (sample):")
                        st.dataframe(pd.DataFrame(failed).head(100))

# End of file
