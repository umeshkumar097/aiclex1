# streamlit_app.py
"""
Aiclex — Result Showing (Polished Streamlit App)
Features:
- Upload Excel/CSV (Hallticket, Email, Location)
- Upload nested ZIP(s) -> extract PDFs -> OCR (pdfplumber -> pdf2image+pytesseract fallback)
- Detect Marks / Pass/Fail / Absent per PDF using rules:
    * If "ABSENT" present -> Absent
    * If "NN PASS"/"NN FAIL" present -> NN marks, status
    * If "Marks Obtained" label present but no number -> Absent
    * Else fallback to last 0-100 number in PDF
    * Pass if marks > 49 else Fail (49 => Fail)
- Match PDFs to Excel by Hallticket (from PDF text or filename)
- Fill Excel columns 'marks' and 'status'
- Produce summary sheets (overall + by-location) and downloadable Excel
- Group PDFs location-wise per recipient, ZIP them, split ZIPs > 3MB, preview
- Send emails via Gmail (smtp.gmail.com TLS) with Test Mode and progress bar
- Robust nested ZIP handling (safe against corrupt or mislabeled entries)
- Polished UI with branding and logs
"""

import os, io, re, time, zipfile, logging
from collections import defaultdict
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import pytesseract

# try pdf2image fallback for better OCR
try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE = True
except Exception:
    PDF2IMAGE = False

# -------------------- CONFIG --------------------
APP_TITLE = "Aiclex — Result Showing"
BRAND = "Aiclex Technologies"
MAX_ATTACHMENT_MB = 3.0
DEFAULT_OCR_DPI = 200
DEFAULT_OCR_LANG = "eng"

# logging
logger = logging.getLogger("aiclex")
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(ch)
logger.setLevel(logging.INFO)

# -------------------- UI setup --------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.markdown(f"<h1 style='color:#0b74de'>{APP_TITLE}</h1><div style='color:gray'>Built by {BRAND}</div>", unsafe_allow_html=True)
st.write("---")
st.info("Follow steps: 1) Upload Excel/CSV & ZIP  2) Process & Preview  3) Prepare ZIPs  4) Send emails (Test Mode available)")

# -------------------- Regex / patterns --------------------
LABEL_RE = re.compile(r"Marks\s*Obtained", re.IGNORECASE)
MARKS_NUM_RE = re.compile(r"\b([0-9]{1,3})\b")
ABSENT_RE = re.compile(r"\b(absent|not present)\b", re.IGNORECASE)
PASSFAIL_RE = re.compile(r"([0-9]{1,3})\s*(PASS|FAIL)", re.IGNORECASE)
HALL_RE = re.compile(r"\b[0-9]{3,}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# -------------------- Sidebar config --------------------
st.sidebar.header("OCR & Email Settings")
tesseract_path = st.sidebar.text_input("Tesseract path (optional)", value=os.environ.get("TESSERACT_CMD",""))
ocr_lang = st.sidebar.text_input("OCR language (e.g. eng or eng+hin)", value=DEFAULT_OCR_LANG)
ocr_dpi = st.sidebar.number_input("OCR DPI (pdf2image)", value=DEFAULT_OCR_DPI, min_value=100, max_value=400, step=10)
attachment_limit_mb = st.sidebar.number_input("Attachment limit (MB)", value=MAX_ATTACHMENT_MB, step=0.5)
send_delay = st.sidebar.number_input("Delay between sends (s)", value=1.0, step=0.5)
show_ocr_debug = st.sidebar.checkbox("Show OCR debug snippets", value=False)
st.sidebar.markdown("---")
st.sidebar.write("Make sure system packages installed: poppler-utils, tesseract-ocr")

if tesseract_path:
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

# -------------------- Utility helpers --------------------
def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ["B","KB","MB","GB"]:
        if n < 1024:
            return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

def is_pdf_bytes(b: bytes) -> bool:
    try:
        return bool(b) and b.lstrip().startswith(b"%PDF")
    except Exception:
        return False

# -------------------- OCR / PDF text extraction --------------------
def extract_text_from_pdf_bytes(pdf_bytes: bytes, dpi: int = 200, lang: str = "eng") -> str:
    """
    Robust extraction:
      1) pdfplumber extract_text
      2) pdf2image -> pytesseract (if available)
      3) pdfplumber page.to_image -> pytesseract
      4) PIL fallback
    """
    texts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                try:
                    t = page.extract_text() or ""
                except Exception:
                    t = ""
                if t and t.strip():
                    texts.append(t)
    except Exception as e:
        logger.debug("pdfplumber read failed: %s", e)

    combined = "\n".join(texts).strip()
    if combined:
        return combined

    # Try pdf2image + pytesseract (preferred OCR fallback)
    if PDF2IMAGE:
        try:
            pages = convert_from_bytes(pdf_bytes, dpi=dpi)
            ocr_texts = []
            for im in pages:
                try:
                    ocr_texts.append(pytesseract.image_to_string(im, lang=lang))
                except Exception:
                    ocr_texts.append(pytesseract.image_to_string(im))
            final = "\n".join(ocr_texts).strip()
            if final:
                return final
        except Exception as e:
            logger.debug("pdf2image fallback failed: %s", e)

    # pdfplumber page.to_image -> pytesseract fallback
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            ocr_texts = []
            for page in pdf.pages:
                try:
                    pil = page.to_image(resolution=dpi).original
                    try:
                        ocr_texts.append(pytesseract.image_to_string(pil, lang=lang))
                    except Exception:
                        ocr_texts.append(pytesseract.image_to_string(pil))
                except Exception:
                    continue
            final = "\n".join(ocr_texts).strip()
            if final:
                return final
    except Exception:
        pass

    # PIL image fallback (rare)
    try:
        im = Image.open(io.BytesIO(pdf_bytes))
        try:
            return pytesseract.image_to_string(im, lang=lang)
        except Exception:
            return pytesseract.image_to_string(im)
    except Exception:
        return ""

# -------------------- PDF parsing (business rules) --------------------
def parse_pdf_bytes(pdf_bytes: bytes, fname: str = "", ocr_dpi: int = DEFAULT_OCR_DPI, ocr_lang_s: str = DEFAULT_OCR_LANG):
    """
    Returns dict:
      {
        pdf_name, pdf_bytes, hallticket (string), marks (int or '' or 'Absent' or None), status (Pass/Fail/Absent),
        text_snippet
      }
    """
    text = extract_text_from_pdf_bytes(pdf_bytes, dpi=ocr_dpi, lang=ocr_lang_s) or ""
    text_norm = text.replace('\xa0', ' ')
    # hallticket: prefer numeric candidate inside text, else digits from filename
    h_cands = HALL_RE.findall(text_norm)
    if h_cands:
        hall = max(h_cands, key=len)
    else:
        fn_digits = re.findall(r"\d+", os.path.basename(fname))
        hall = fn_digits[-1] if fn_digits else ""

    # determine marks/status per rules
    marks = None
    status = "Absent"

    # 1) If 'ABSENT' present anywhere -> Absent
    if ABSENT_RE.search(text_norm):
        marks = ""
        status = "Absent"
    else:
        # 2) PASS/FAIL lines with number -> NN PASS/FAIL
        pf = PASSFAIL_RE.search(text_norm)
        if pf:
            try:
                val = int(pf.group(1))
                marks = val
                status = "Pass" if val > 49 else "Fail"
            except:
                marks = ""
                status = "Absent"
        else:
            # 3) If "Marks Obtained" label present -> look for number after label else Absent
            lbl = LABEL_RE.search(text_norm)
            if lbl:
                snippet = text_norm[lbl.end():lbl.end() + 200]  # check next chars for number
                mnum = re.search(r"([0-9]{1,3})", snippet)
                if mnum:
                    val = int(mnum.group(1))
                    marks = val
                    status = "Pass" if val > 49 else "Fail"
                else:
                    marks = ""
                    status = "Absent"
            else:
                # 4) fallback: last 0-100 number in document
                nums = MARKS_NUM_RE.findall(text_norm)
                nums = [int(n) for n in nums if 0 <= int(n) <= 100]
                if nums:
                    val = nums[-1]
                    marks = val
                    status = "Pass" if val > 49 else "Fail"
                else:
                    marks = ""
                    status = "Absent"

    return {
        "pdf_name": os.path.basename(fname),
        "pdf_bytes": pdf_bytes,
        "hallticket": str(hall).strip(),
        "marks": marks,
        "status": status,
        "text_snippet": (text_norm[:2000] if show_ocr_debug else "")
    }

# -------------------- Robust nested ZIP extraction --------------------
def extract_from_zip_recursive(zip_bytes: bytes, ocr_dpi: int, ocr_lang_s: str):
    """
    Recursively extract PDFs from possibly nested zips.
    Handles mislabeled entries gracefully: if .zip entry is not a valid zip,
    tries treating it as PDF if header matches; otherwise skip.
    Returns list of parse result dicts.
    """
    results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                name = info.filename
                try:
                    data = zf.read(info)
                except Exception as e:
                    logger.warning("Cannot read entry %s: %s", name, e)
                    continue
                lname = name.lower()
                # nested zip
                if lname.endswith(".zip"):
                    # attempt open nested zip; guard BadZipFile
                    try:
                        nested = extract_from_zip_recursive(data, ocr_dpi, ocr_lang_s)
                        results.extend(nested)
                    except zipfile.BadZipFile:
                        # not a real zip - maybe PDF bytes inside mislabeled .zip file
                        if is_pdf_bytes(data):
                            try:
                                results.append(parse_pdf_bytes(data, fname=name, ocr_dpi=ocr_dpi, ocr_lang_s=ocr_lang_s))
                            except Exception as e:
                                logger.warning("Failed parse mislabeled PDF %s: %s", name, e)
                        else:
                            logger.info("Skipping non-zip, non-pdf entry: %s", name)
                elif lname.endswith(".pdf"):
                    try:
                        results.append(parse_pdf_bytes(data, fname=name, ocr_dpi=ocr_dpi, ocr_lang_s=ocr_lang_s))
                    except Exception as e:
                        logger.warning("Failed parse PDF %s: %s", name, e)
                else:
                    # unknown extension -> maybe pdf content without .pdf extension
                    if is_pdf_bytes(data):
                        try:
                            results.append(parse_pdf_bytes(data, fname=name, ocr_dpi=ocr_dpi, ocr_lang_s=ocr_lang_s))
                        except Exception as e:
                            logger.warning("Failed parse raw-PDF %s: %s", name, e)
                    else:
                        # ignore other files
                        continue
    except zipfile.BadZipFile:
        # Top-level not a zip file
        raise
    return results

# -------------------- Excel fill & summary --------------------
def fill_excel_using_pdf_data(df: pd.DataFrame, pdf_data: list, hall_col: str):
    # Build map hall -> best pdf marks (prefer numeric over Absent)
    pdf_map = {}
    for p in pdf_data:
        k = str(p.get("hallticket", "")).strip()
        if not k:
            continue
        existing = pdf_map.get(k)
        # prefer numeric marks over blank/''
        if existing is None:
            pdf_map[k] = p
        else:
            # if existing has no numeric and new has numeric, replace
            if (not isinstance(existing.get("marks"), int)) and isinstance(p.get("marks"), int):
                pdf_map[k] = p

    # ensure columns
    marks_col = "marks"
    status_col = "status"
    if marks_col not in df.columns: df[marks_col] = ""
    if status_col not in df.columns: df[status_col] = ""

    filled = 0
    unmatched = []
    for idx, row in df.iterrows():
        ht = str(row.get(hall_col, "")).strip()
        if not ht:
            unmatched.append({"index": idx, "reason": "no_hallticket"})
            continue
        val = None
        if ht in pdf_map:
            val = pdf_map[ht]
        else:
            # try digits-only matching
            digits = re.sub(r"\D", "", ht)
            if digits and digits in pdf_map:
                val = pdf_map[digits]
            else:
                # try endswith cases
                for k in pdf_map.keys():
                    kd = re.sub(r"\D", "", str(k))
                    if kd and (k.endswith(digits) or digits.endswith(kd) or kd.endswith(digits)):
                        val = pdf_map[k]
                        break
        if val is None:
            # per rule: if no mark found -> treat Absent
            df.at[idx, marks_col] = ""
            df.at[idx, status_col] = "Absent"
        else:
            m = val.get("marks")
            if isinstance(m, int):
                df.at[idx, marks_col] = int(m)
                df.at[idx, status_col] = "Pass" if int(m) > 49 else "Fail"
            else:
                # '' or 'Absent'
                df.at[idx, marks_col] = ""
                df.at[idx, status_col] = "Absent"
        filled += 1

    return df, filled, unmatched, pdf_map

def make_summary_sheets(df: pd.DataFrame, location_col: str):
    total = len(df)
    pass_count = int((df['status'] == 'Pass').sum())
    fail_count = int((df['status'] == 'Fail').sum())
    absent_count = int((df['status'] == 'Absent').sum())
    summary_overall = pd.DataFrame([{"Total": total, "Pass": pass_count, "Fail": fail_count, "Absent": absent_count}])
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

# -------------------- ZIP packing & splitting --------------------
def make_zip_bytes(file_entries):
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, content in file_entries:
            zf.writestr(fname, content)
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
            # put current file by itself (may exceed limit)
            parts.append((f"{zip_name_prefix}_part{len(parts)+1}.zip", make_zip_bytes([(fname, b)])))
            current = []
    if current:
        parts.append((f"{zip_name_prefix}_part{len(parts)+1}.zip", make_zip_bytes(current)))
    return parts

# -------------------- Email send (Gmail TLS) --------------------
def send_email_with_attachments_gmail(smtp_user, smtp_pass, to_email, subject, body, attachments):
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
            s.ehlo(); s.starttls(); s.ehlo()
            s.login(smtp_user, smtp_pass)
            s.send_message(msg)
        return True, None
    except Exception as e:
        logger.error("SMTP send failed: %s", e)
        return False, str(e)

# -------------------- Streamlit main flow --------------------
st.header("Step 1 — Upload Excel/CSV and ZIP (nested allowed)")
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("Upload Excel or CSV", type=["xlsx","csv"])
with col2:
    uploaded_zip = st.file_uploader("Upload ZIP (can contain nested zips with PDFs)", type=["zip"])

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

    st.success(f"Excel loaded — {len(df)} rows")
    cols = df.columns.tolist()
    hall_col = st.selectbox("Select Hallticket column", cols)
    email_col = st.selectbox("Select Email column", cols)
    location_col = st.selectbox("Select Location column", cols)

    # process nested zip -> extract pdf_info list
    with st.spinner("Processing ZIP(s) and running OCR (may take time)..."):
        try:
            pdf_data = extract_from_zip_recursive(uploaded_zip.read(), ocr_dpi=ocr_dpi, ocr_lang_s=ocr_lang)
        except zipfile.BadZipFile:
            st.error("Uploaded file is not a valid ZIP archive. Please upload a valid zip.")
            pdf_data = []
    st.info(f"PDF records extracted: {len(pdf_data)}")

    if show_ocr_debug and pdf_data:
        st.subheader("OCR debug (first 500 chars)")
        debug_rows = [{"pdf_name": p["pdf_name"], "hallticket": p["hallticket"], "marks": p["marks"], "status": p["status"], "text_snippet": p.get("text_snippet","")[:500]} for p in pdf_data]
        st.dataframe(pd.DataFrame(debug_rows).head(200))

    # fill excel
    updated_df, filled_count, unmatched, pdf_map = fill_excel_using_pdf_data(df.copy(), pdf_data, hall_col)
    st.success(f"Filled {filled_count} rows with marks/status")
    if unmatched:
        st.warning(f"{len(unmatched)} rows had no hallticket value.")

    st.subheader("Preview updated results (first 100 rows)")
    st.dataframe(updated_df.head(100))

    # prepare summary excel for download
    sheets = make_summary_sheets(updated_df, location_col)
    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    out_buf.seek(0)
    st.download_button("Download results + summary (Excel)", data=out_buf,
                       file_name=f"aiclex_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # grouping: build pdf map hall->list and recipients grouping
    st.markdown("---")
    st.header("Step 2 — Group PDFs by Location & Prepare ZIPs")
    # map hallticket->pdf list
    pdf_map_multi = defaultdict(list)
    for p in pdf_data:
        k = str(p.get("hallticket","")).strip()
        if k:
            pdf_map_multi[k].append(p)

    recipients = defaultdict(lambda: defaultdict(list))  # email -> location -> list of (fname, bytes)
    missing_log = []
    for _, row in updated_df.iterrows():
        ht = str(row.get(hall_col,"")).strip()
        emails_raw = str(row.get(email_col,"")).strip()
        loc = str(row.get(location_col,"")).strip() or "Unknown"
        if not emails_raw:
            continue
        emails = [e.strip() for e in re.split(r"[;, \n]+", emails_raw) if e.strip()]
        for e in emails:
            if not EMAIL_RE.match(e):
                continue
            # attach all pdfs matching this hallticket
            found = False
            if ht in pdf_map_multi:
                for p in pdf_map_multi[ht]:
                    recipients[e][loc].append((f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}", p["pdf_bytes"]))
                found = True
            else:
                # try digits fallback
                digits = re.sub(r"\D","", ht)
                for k, lst in pdf_map_multi.items():
                    kd = re.sub(r"\D","", str(k))
                    if kd and (kd == digits or kd.endswith(digits) or digits.endswith(kd)):
                        for p in lst:
                            recipients[e][loc].append((f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}", p["pdf_bytes"]))
                        found = True
                        break
            if not found:
                missing_log.append({"email": e, "hallticket": ht, "location": loc})

    # show recipient summary
    rec_preview = []
    for em, info in list(recipients.items())[:200]:
        files_count = sum(len(lst) for lst in info.values())
        rec_preview.append({"email": em, "locations": ", ".join(info.keys()), "files": files_count})
    if rec_preview:
        st.dataframe(pd.DataFrame(rec_preview))
    else:
        st.info("No recipients with matched PDFs found. Check hallticket or email columns.")

    if missing_log:
        st.warning(f"{len(missing_log)} rows had no matching PDFs (sample shown).")
        st.dataframe(pd.DataFrame(missing_log).head(50))

    # prepare zip parts in-memory
    if st.button("Prepare ZIPs (grouped by recipient->location)"):
        st.info("Preparing ZIP parts (kept in memory).")
        prepared = {}
        max_bytes = int(attachment_limit_mb * 1024 * 1024)
        for em, locs in recipients.items():
            prepared[em] = []
            for loc, files in locs.items():
                parts = split_files_into_zip_parts(files, max_bytes, zip_name_prefix=re.sub(r"[^A-Za-z0-9]+","_", loc)[:40])
                prepared[em].append((loc, parts))
        st.session_state["prepared"] = prepared
        st.success("Prepared ZIP parts stored in session memory")

    if "prepared" in st.session_state:
        st.subheader("ZIP preview (sample)")
        preview_rows = []
        for em, locs in st.session_state["prepared"].items():
            for loc, parts in locs:
                for i, (pname, pbytes) in enumerate(parts, start=1):
                    preview_rows.append({"email": em, "location": loc, "part": i, "zip_name": pname, "size": human_bytes(len(pbytes))})
        st.dataframe(pd.DataFrame(preview_rows).head(200))

        # send emails
        st.markdown("---")
        st.header("Step 3 — Send Emails")
        smtp_user = st.text_input("Gmail address (SMTP user)", value=os.environ.get("SMTP_USER",""))
        smtp_pass = st.text_input("Gmail App Password (SMTP pass)", type="password", value=os.environ.get("SMTP_PASS",""))
        test_mode = st.checkbox("Test mode (send all to test email)", value=True)
        test_email = st.text_input("Test email (if test mode ON)", value=os.environ.get("TEST_EMAIL",""))
        subj_template = st.text_input("Subject template", value="Results for {location}")
        body_template = st.text_area("Body template (use {location} {part}/{total_parts})", value="Hello,\n\nPlease find attached results for {location} (Part {part}/{total_parts}).\n\nRegards,\nAiclex")

        if st.button("Start sending prepared ZIPs"):
            if not smtp_user or not smtp_pass:
                st.error("Provide Gmail + App Password for SMTP.")
            else:
                # count total sends
                total_sends = 0
                for em, locs in st.session_state["prepared"].items():
                    recipients_list = [test_email] if test_mode else [em]
                    for loc, parts in locs:
                        total_sends += len(recipients_list) * max(1, len(parts))
                if total_sends == 0:
                    st.warning("No ZIP parts prepared to send.")
                else:
                    progress = st.progress(0)
                    status = st.empty()
                    sent_count = 0
                    failed = []
                    cur = 0
                    for em, locs in st.session_state["prepared"].items():
                        recipients_list = [test_email] if test_mode else [em]
                        for loc, parts in locs:
                            total_parts = max(1, len(parts))
                            for part_idx, (zipname, zipbytes) in enumerate(parts, start=1):
                                subj = subj_template.format(location=loc, part=part_idx, total_parts=total_parts)
                                body = body_template.format(location=loc, part=part_idx, total_parts=total_parts)
                                for r in recipients_list:
                                    ok, err = send_email_with_attachments_gmail(smtp_user, smtp_pass, r, subj, body, [(zipname, zipbytes)])
                                    cur += 1
                                    progress.progress(min(1.0, cur / total_sends))
                                    status.text(f"Sending {cur}/{total_sends} → {r} ({loc} part {part_idx}/{total_parts})")
                                    if ok:
                                        sent_count += 1
                                    else:
                                        failed.append({"recipient": r, "loc": loc, "zip": zipname, "error": err})
                                    time.sleep(send_delay)
                    st.success(f"Sending finished. Sent: {sent_count}, Failed: {len(failed)}")
                    if failed:
                        st.error("Some sends failed (sample):")
                        st.dataframe(pd.DataFrame(failed).head(50))

st.write("---")
st.markdown(f"<div style='color:gray; font-size:12px'>App by {BRAND} — Aiclex</div>", unsafe_allow_html=True)
