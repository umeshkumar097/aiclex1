# streamlit_app.py
"""
Aiclex — Result Showing (final combined)
"""

import os
import io
import re
import time
import zipfile
import logging
from collections import defaultdict
from datetime import datetime
from email.message import EmailMessage

import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import pytesseract

# optional faster OCR pipeline
try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE = True
except Exception:
    PDF2IMAGE = False

# ---------------- Config / Branding ----------------
APP_TITLE = "Aiclex — Result Showing"
BRAND = "Aiclex Technologies"
DEFAULT_OCR_DPI = 200
DEFAULT_OCR_LANG = "eng"
DEFAULT_ATTACHMENT_MB = 3.0

# logging
logger = logging.getLogger("aiclex")
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(ch)
logger.setLevel(logging.INFO)

# ---------------- Streamlit UI setup ----------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.markdown(f"<h1 style='color:#0b74de'>{APP_TITLE}</h1><div style='color:gray'>Built by {BRAND}</div>", unsafe_allow_html=True)
st.write("---")
st.info("Steps: 1) Upload Excel & ZIP, 2) Process & Preview, 3) Prepare ZIPs, 4) Send (Test Mode available).")

# ---------------- Patterns ----------------
LABEL_RE = re.compile(r"Marks\s*Obtained", re.IGNORECASE)
MARKS_NUM_RE = re.compile(r"\b([0-9]{1,3})\b")
ABSENT_RE = re.compile(r"\b(absent|not present)\b", re.IGNORECASE)
PASSFAIL_RE = re.compile(r"([0-9]{1,3})\s*(PASS|FAIL)", re.IGNORECASE)
HALL_RE = re.compile(r"\b[0-9]{3,}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# ---------------- Sidebar config ----------------
st.sidebar.header("OCR & Email Settings")
tesseract_path = st.sidebar.text_input("Tesseract path (optional)", value=os.environ.get("TESSERACT_CMD",""))
ocr_lang = st.sidebar.text_input("OCR language (e.g. eng or eng+hin)", value=DEFAULT_OCR_LANG)
ocr_dpi = st.sidebar.number_input("OCR DPI (pdf2image)", value=int(DEFAULT_OCR_DPI), min_value=100, max_value=400, step=10)
attachment_limit_mb = st.sidebar.number_input("Attachment limit (MB)", value=float(DEFAULT_ATTACHMENT_MB), step=0.5)
send_delay = st.sidebar.number_input("Delay between sends (s)", value=1.0, step=0.5)
show_ocr_debug = st.sidebar.checkbox("Show OCR debug snippet", value=False)
st.sidebar.markdown("Install system packages if needed: poppler-utils, tesseract-ocr")

if tesseract_path:
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

# ---------------- Helpers ----------------
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

def is_pdf_bytes(b: bytes) -> bool:
    try:
        return bool(b) and b.lstrip().startswith(b"%PDF")
    except Exception:
        return False

# OCR / text extraction
def extract_text_from_pdf_bytes(pdf_bytes: bytes, dpi: int = DEFAULT_OCR_DPI, lang: str = DEFAULT_OCR_LANG) -> str:
    # 1) pdfplumber text
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
    except Exception:
        pass
    combined = "\n".join(texts).strip()
    if combined:
        return combined

    # 2) pdf2image -> pytesseract
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
        except Exception:
            pass

    # 3) pdfplumber page.to_image -> pytesseract
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

    # 4) PIL fallback
    try:
        im = Image.open(io.BytesIO(pdf_bytes))
        try:
            return pytesseract.image_to_string(im, lang=lang)
        except Exception:
            return pytesseract.image_to_string(im)
    except Exception:
        return ""

# parse PDF according to given rules
def parse_pdf_bytes(pdf_bytes: bytes, fname: str = "", ocr_dpi: int = DEFAULT_OCR_DPI, ocr_lang_s: str = DEFAULT_OCR_LANG):
    text = extract_text_from_pdf_bytes(pdf_bytes, dpi=ocr_dpi, lang=ocr_lang_s) or ""
    text_norm = text.replace('\xa0', ' ')
    # hallticket candidate
    h_cands = HALL_RE.findall(text_norm)
    if h_cands:
        hall = max(h_cands, key=len)
    else:
        fn_digits = re.findall(r"\d+", os.path.basename(fname))
        hall = fn_digits[-1] if fn_digits else ""
    # compute marks/status
    marks = None
    status = "Absent"
    if ABSENT_RE.search(text_norm):
        marks = ""
        status = "Absent"
    else:
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
            lbl = LABEL_RE.search(text_norm)
            if lbl:
                snippet = text_norm[lbl.end():lbl.end()+200]
                mnum = re.search(r"([0-9]{1,3})", snippet)
                if mnum:
                    val = int(mnum.group(1))
                    marks = val
                    status = "Pass" if val > 49 else "Fail"
                else:
                    marks = ""
                    status = "Absent"
            else:
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

# Robust recursive ZIP extraction
def extract_from_zip_recursive(zip_bytes: bytes, ocr_dpi: int, ocr_lang_s: str):
    results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = zf.namelist()
            total = len(names)
            prog_place = st.empty()
            prog = st.progress(0)
            for i, name in enumerate(names, start=1):
                prog_place.text(f"Scanning archive entry {i}/{total}: {name}")
                try:
                    data = zf.read(name)
                except Exception as e:
                    logger.warning("Cannot read entry %s: %s", name, e)
                    prog.progress(i/total)
                    continue
                lname = name.lower()
                if lname.endswith(".zip"):
                    try:
                        nested = extract_from_zip_recursive(data, ocr_dpi, ocr_lang_s)
                        results.extend(nested)
                    except zipfile.BadZipFile:
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
                    if is_pdf_bytes(data):
                        try:
                            results.append(parse_pdf_bytes(data, fname=name, ocr_dpi=ocr_dpi, ocr_lang_s=ocr_lang_s))
                        except Exception as e:
                            logger.warning("Failed parse raw-PDF %s: %s", name, e)
                    else:
                        # ignore other files
                        pass
                prog.progress(i/total)
            prog_place.empty()
    except zipfile.BadZipFile:
        # top-level not a zip -> caller will handle
        raise
    return results

# Fill excel logic
def fill_excel_using_pdf_data(df: pd.DataFrame, pdf_data: list, hall_col: str):
    # build pdf map prefer numeric
    pdf_map = {}
    for p in pdf_data:
        k = str(p.get("hallticket","")).strip()
        if not k:
            continue
        existing = pdf_map.get(k)
        if existing is None:
            pdf_map[k] = p
        else:
            if (not isinstance(existing.get("marks"), int)) and isinstance(p.get("marks"), int):
                pdf_map[k] = p

    marks_col = "marks"
    status_col = "status"
    if marks_col not in df.columns:
        df[marks_col] = ""
    if status_col not in df.columns:
        df[status_col] = ""

    filled = 0
    unmatched = []
    # progress
    place = st.empty()
    prog = st.progress(0)
    total = len(df)
    for i, (idx, row) in enumerate(df.iterrows(), start=1):
        place.text(f"Filling row {i}/{total}")
        ht = str(row.get(hall_col,"")).strip()
        if not ht:
            unmatched.append({"index": idx, "reason": "no_hallticket"})
            prog.progress(i/total)
            continue
        val = None
        if ht in pdf_map:
            val = pdf_map[ht]
        else:
            digits = re.sub(r"\D","", ht)
            if digits and digits in pdf_map:
                val = pdf_map[digits]
            else:
                for k in pdf_map.keys():
                    kd = re.sub(r"\D","", str(k))
                    if kd and (k.endswith(digits) or digits.endswith(kd) or kd.endswith(digits)):
                        val = pdf_map[k]
                        break
        if val is None:
            df.at[idx, marks_col] = ""
            df.at[idx, status_col] = "Absent"
        else:
            m = val.get("marks")
            if isinstance(m,int):
                df.at[idx, marks_col] = int(m)
                df.at[idx, status_col] = "Pass" if int(m) > 49 else "Fail"
            else:
                df.at[idx, marks_col] = ""
                df.at[idx, status_col] = "Absent"
        filled += 1
        prog.progress(i/total)
    place.empty()
    return df, filled, unmatched, pdf_map

# ZIP helpers
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
    part_no = 1
    def flush_current():
        nonlocal part_no, parts, current
        if not current:
            return
        zip_name = f"{zip_name_prefix}_part{part_no}.zip"
        parts.append((zip_name, make_zip_bytes(current)))
        part_no += 1
        current = []
    for fname, b in file_entries:
        test = current + [(fname, b)]
        test_zip = make_zip_bytes(test)
        if len(test_zip) <= max_bytes:
            current = test
            continue
        # can't add -> flush current if any
        if current:
            flush_current()
        # try single file
        single_zip = make_zip_bytes([(fname, b)])
        if len(single_zip) <= max_bytes:
            current = [(fname, b)]
        else:
            # file alone exceeds limit -> put as its own part (may be > max_bytes)
            zip_name = f"{zip_name_prefix}_part{part_no}.zip"
            parts.append((zip_name, single_zip))
            part_no += 1
            current = []
    if current:
        flush_current()
    return parts

# Email send
def send_email_with_attachments_gmail(smtp_user, smtp_pass, to_emails, subject, body, attachments):
    msg = EmailMessage()
    msg["From"] = smtp_user
    if isinstance(to_emails, list):
        msg["To"] = ", ".join(to_emails)
    else:
        msg["To"] = to_emails
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

# ---------------- Main UI Flow ----------------
st.header("Step 1 — Upload Excel/CSV and ZIP")
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("Upload Excel or CSV", type=["xlsx","csv"])
with col2:
    uploaded_zip = st.file_uploader("Upload ZIP (nested zips with PDFs)", type=["zip"])

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

    # process zip -> pdf_data
    with st.spinner("Processing ZIP(s) and running OCR (may take time)..."):
        try:
            pdf_data = extract_from_zip_recursive(uploaded_zip.read(), ocr_dpi=ocr_dpi, ocr_lang_s=ocr_lang)
        except zipfile.BadZipFile:
            st.error("Uploaded file is not a valid ZIP archive.")
            pdf_data = []
    st.info(f"PDF records extracted: {len(pdf_data)}")

    if show_ocr_debug and pdf_data:
        st.subheader("OCR debug (sample snippets)")
        debug_rows = [{"pdf_name": p["pdf_name"], "hallticket": p["hallticket"], "marks": p["marks"], "status": p["status"], "text_snippet": p.get("text_snippet","")[:500]} for p in pdf_data]
        st.dataframe(pd.DataFrame(debug_rows).head(200))

    # fill excel with progress
    updated_df, filled_count, unmatched, pdf_map = fill_excel_using_pdf_data(df.copy(), pdf_data, hall_col)
    st.success(f"Filled {filled_count} rows (marks/status).")
    if unmatched:
        st.warning(f"{len(unmatched)} rows had missing hallticket.")

    st.subheader("Preview updated results (first 100 rows)")
    st.dataframe(updated_df.head(100))

    # download results + summary
    sheets = {}
    total = len(updated_df)
    pass_count = int((updated_df['status'] == 'Pass').sum())
    fail_count = int((updated_df['status'] == 'Fail').sum())
    absent_count = int((updated_df['status'] == 'Absent').sum())
    summary_overall = pd.DataFrame([{"Total": total, "Pass": pass_count, "Fail": fail_count, "Absent": absent_count}])
    sheets["results"] = updated_df
    sheets["summary_overall"] = summary_overall
    # optional per-location summary
    if location_col in updated_df.columns:
        by_loc = []
        for loc, g in updated_df.groupby(location_col):
            by_loc.append({"Location": loc, "Total": len(g), "Pass": int((g['status']=="Pass").sum()), "Fail": int((g['status']=="Fail").sum()), "Absent": int((g['status']=="Absent").sum())})
        sheets["summary_by_location"] = pd.DataFrame(by_loc)

    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    out_buf.seek(0)
    st.download_button("Download results + summary (Excel)", data=out_buf, file_name=f"aiclex_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Group PDFs per recipient (email) and per location
    st.markdown("---")
    st.header("Step 2 — Prepare ZIPs grouped by recipient & location")
    # build hallticket->list map
    pdf_map_multi = defaultdict(list)
    for p in pdf_data:
        k = str(p.get("hallticket","")).strip()
        if k:
            pdf_map_multi[k].append(p)

    recipients = defaultdict(lambda: defaultdict(list))  # email -> location -> list of (fname, bytes)
    missing_log = []
    # iterate rows safely
    for idx, row in updated_df.iterrows():
        ht = str(row.get(hall_col,"")).strip()
        emails_raw = str(row.get(email_col,"")).strip()
        loc = str(row.get(location_col,"")).strip() or "Unknown"
        if not emails_raw:
            continue
        emails = [e.strip() for e in re.split(r"[;, \n]+", emails_raw) if e.strip()]
        found_any = False
        for e in emails:
            if ht and ht in pdf_map_multi:
                for p in pdf_map_multi[ht]:
                    recipients[e][loc].append((f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}", p["pdf_bytes"]))
                found_any = True
            else:
                digits = re.sub(r"\D","", ht)
                matched = False
                if digits:
                    for k, lst in pdf_map_multi.items():
                        kd = re.sub(r"\D","", str(k))
                        if kd and (kd == digits or kd.endswith(digits) or digits.endswith(kd)):
                            for p in lst:
                                recipients[e][loc].append((f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}", p["pdf_bytes"]))
                            matched = True
                            found_any = True
                            break
        if not found_any:
            missing_log.append({"index": idx, "hallticket": ht, "emails": emails, "location": loc})

    st.info(f"Recipients prepared: {len(recipients)} (sample below)")
    rec_preview = []
    for em, locs in list(recipients.items())[:200]:
        files_count = sum(len(lst) for lst in locs.values())
        rec_preview.append({"email": em, "locations": ", ".join(locs.keys()), "files": files_count})
    if rec_preview:
        st.dataframe(pd.DataFrame(rec_preview))
    if missing_log:
        st.warning(f"{len(missing_log)} rows had no matching PDFs (sample):")
        st.dataframe(pd.DataFrame(missing_log).head(50))

    # Prepare ZIP parts in memory with progress
    if st.button("Prepare ZIPs (grouped by recipient->location)"):
        st.info("Preparing ZIP parts in memory (may use RAM).")
        max_bytes = int(attachment_limit_mb * 1024 * 1024)
        prepared = {}
        total_recipients = len(recipients)
        prog_place = st.empty()
        prog = st.progress(0)
        for i, (em, locs) in enumerate(recipients.items(), start=1):
            prog_place.text(f"Preparing for recipient {i}/{total_recipients}: {em}")
            prepared[em] = []
            for loc, files in locs.items():
                safe_prefix = re.sub(r"[^A-Za-z0-9]+","_", loc)[:40] or "loc"
                parts = split_files_into_zip_parts(files, max_bytes, zip_name_prefix=safe_prefix)
                prepared[em].append((loc, parts))
            prog.progress(i/total_recipients)
        prog_place.empty()
        st.session_state["prepared"] = prepared
        st.success("Prepared ZIP parts stored in session memory.")

# ---------------- Preview & Send ----------------
if "prepared" in st.session_state:
    st.header("Step 3 — ZIP Preview (by recipient & location)")
    preview_rows = []
    # Build a location-wise summary as user requested
    location_summary = defaultdict(list)  # location -> list of part names
    for em, locs in st.session_state["prepared"].items():
        for loc, parts in locs:
            for pname, pbytes in parts:
                preview_rows.append({"email": em, "location": loc, "zip_name": pname, "size": human_bytes(len(pbytes))})
                location_summary[loc].append(pname)
    if preview_rows:
        st.dataframe(pd.DataFrame(preview_rows).head(500))
    # Location parts summary
    loc_summary_rows = []
    for loc, partnames in location_summary.items():
        loc_summary_rows.append({"Location": loc, "PartsCount": len(partnames), "Parts": ", ".join(partnames)})
    if loc_summary_rows:
        st.subheader("Location-wise parts summary")
        st.dataframe(pd.DataFrame(loc_summary_rows))

    st.markdown("---")
    st.header("Step 4 — Send Emails")
    smtp_user = st.text_input("Gmail address (SMTP user)", value="info@cruxmanagement.com")
    smtp_pass = st.text_input("Gmail App Password (SMTP pass)", type="password", value="norx wxop hvsm bvfu")
    test_mode = st.checkbox("Test mode (send all to test email)", value=True)
    test_email = st.text_input("Test email (if test mode ON)")
    subj_template = st.text_input("Subject template", value="Results for {location} (Part {part}/{total_parts})")
    body_template = st.text_area("Body template", value="Hello,\n\nPlease find attached results for {location} (Part {part}/{total_parts}).\n\nRegards,\nAiclex")

    if st.button("Start sending prepared ZIPs"):
        if not smtp_user or not smtp_pass:
            st.error("Provide Gmail address and App Password (2FA app password).")
        else:
            # count sends
            total_sends = 0
            for em, locs in st.session_state["prepared"].items():
                recipient_list = [e.strip() for e in re.split(r"[;, \n]+", em) if e.strip()]
                if test_mode:
                    recipient_list = [test_email] if test_email else []
                for loc, parts in locs:
                    total_sends += len(recipient_list) * max(1, len(parts))
            if total_sends == 0:
                st.warning("No prepared zips to send.")
            else:
                progress = st.progress(0)
                status = st.empty()
                sent_count = 0
                failed = []
                cur = 0
                items = list(st.session_state["prepared"].items())
                total_items = len(items)
                for ri, (em, locs) in enumerate(items, start=1):
                    status.text(f"Processing recipient {ri}/{total_items}: {em}")
                    recipient_list = [e.strip() for e in re.split(r"[;, \n]+", em) if e.strip()]
                    if test_mode:
                        recipient_list = [test_email] if test_email else []
                    for loc, parts in locs:
                        total_parts = max(1, len(parts))
                        for part_idx, (zipname, zipbytes) in enumerate(parts, start=1):
                            subj = subj_template.format(location=loc, part=part_idx, total_parts=total_parts)
                            body = body_template.format(location=loc, part=part_idx, total_parts=total_parts)
                            ok, err = send_email_with_attachments_gmail(smtp_user, smtp_pass, recipient_list, subj, body, [(zipname, zipbytes)])
                            cur += 1
                            progress.progress(min(1.0, cur / total_sends))
                            status.text(f"Sent {cur}/{total_sends} → {recipient_list} ({loc} part {part_idx}/{total_parts})")
                            if ok:
                                sent_count += 1
                            else:
                                failed.append({"recipients": recipient_list, "loc": loc, "zip": zipname, "error": err})
                            time.sleep(send_delay)
                status.empty()
                st.success(f"Sending finished. Sent: {sent_count}. Failed: {len(failed)}")
                if failed:
                    st.error("Some sends failed (sample):")
                    st.dataframe(pd.DataFrame(failed).head(50))

st.write("---")
st.markdown(f"<div style='color:gray; font-size:12px'>App by {BRAND} — Aiclex</div>", unsafe_allow_html=True)
