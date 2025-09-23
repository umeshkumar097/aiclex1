# streamlit_app.py
"""
Aiclex Hallticket Mailer — Final (with result extraction + email)
Integrated:
- Extract PDFs from nested ZIPs (existing)
- EXTRACT FROM PDF: hallticket, marks, Absent (NEW) using pdfplumber + pytesseract fallback
- Fill uploaded Excel (marks, status) before existing grouping / ZIP / email send flow
- Keep your existing grouping, splitting, and SMTP send UI intact
"""

import os, io, re, time, zipfile, tempfile, shutil, smtplib
from email.message import EmailMessage
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import streamlit as st
import pandas as pd

# PDF libs (added)
import pdfplumber
from PIL import Image
import pytesseract

# ---------------- Streamlit config ----------------
st.set_page_config(page_title="Aiclex Mailer — Final", layout="wide")
st.title("📩 Aiclex Technologies — Final Hall Ticket Mailer")

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
HALL_FILENAME_RE = re.compile(r"\d+")
HALL_IN_PDF_RE = re.compile(r"\b[0-9]{4,}\b")  # adjust if needed
ABSENT_PATTERN = re.compile(r"\b(?:absent|not present|a\s*b\s*s\s*e\s*n\s*t|a\.?b\.?s\b|abs\b)\b", re.IGNORECASE)
MARKS_PATTERN = re.compile(r"(?:marks|mark|score|total|obtained|obtained[:\s\-]*)[:\s\-]*([0-9]{1,3})", re.IGNORECASE)

# ---------- small utils ----------
def human_bytes(n):
    try: n = float(n)
    except: return ""
    for unit in ["B","KB","MB","GB"]:
        if n < 1024: return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

# ---------------- existing helpers from your file ----------------
def extract_zip_bytes_recursively(zip_bytes, out_root):
    """Extract PDFs from nested ZIPs"""
    extracted_pdfs = []
    extraction_root = tempfile.mkdtemp(prefix="aiclex_unzip_", dir=out_root)
    def _process_zip(data_bytes, curdir):
        try:
            with zipfile.ZipFile(io.BytesIO(data_bytes)) as zf:
                for info in zf.infolist():
                    if info.is_dir(): continue
                    lname = info.filename.lower()
                    entry_bytes = zf.read(info)
                    if lname.endswith(".zip"):
                        nested_dir = os.path.join(curdir, os.path.splitext(os.path.basename(info.filename))[0])
                        os.makedirs(nested_dir, exist_ok=True)
                        _process_zip(entry_bytes, nested_dir)
                    elif lname.endswith(".pdf"):
                        os.makedirs(curdir, exist_ok=True)
                        target = os.path.join(curdir, os.path.basename(info.filename))
                        if os.path.exists(target):
                            base, ext = os.path.splitext(os.path.basename(info.filename))
                            target = os.path.join(curdir, f"{base}_{int(time.time()*1000)}{ext}")
                        with open(target, "wb") as wf:
                            wf.write(entry_bytes)
                        extracted_pdfs.append(os.path.abspath(target))
        except Exception:
            return
    _process_zip(zip_bytes, extraction_root)
    return extracted_pdfs, extraction_root

def extract_hallticket_from_filename(path):
    base = os.path.splitext(os.path.basename(path))[0]
    digits = re.findall(r"\d+", base)
    return digits[-1] if digits else None

def create_split_zips(files, out_dir, base_name, max_bytes):
    """Split files into multiple zips if total > max_bytes"""
    os.makedirs(out_dir, exist_ok=True)
    zips = []
    cur, cur_size, part = [], 0, 1
    for f in files:
        size = os.path.getsize(f)
        if cur and (cur_size + size) > max_bytes:
            zpath = os.path.join(out_dir, f"{base_name}_part{part}.zip")
            with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for x in cur: z.write(x, arcname=os.path.basename(x))
            zips.append(zpath)
            cur, cur_size, part = [f], size, part+1
        else:
            cur.append(f); cur_size += size
    if cur:
        zpath = os.path.join(out_dir, f"{base_name}_part{part}.zip")
        with zipfile.ZipFile(zpath, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for x in cur: z.write(x, arcname=os.path.basename(x))
        zips.append(zpath)
    return zips

def send_email_smtp(cfg, recipients, subject, body, attachments):
    msg = EmailMessage()
    msg["From"] = cfg["sender"]
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)
    for ap in attachments:
        with open(ap, "rb") as af: data = af.read()
        msg.add_attachment(data, maintype="application", subtype="zip", filename=os.path.basename(ap))
    if cfg.get("use_ssl", True):
        server = smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=60)
    else:
        server = smtplib.SMTP(cfg["host"], cfg["port"], timeout=60)
        server.starttls()
    if cfg.get("password"): server.login(cfg["sender"], cfg["password"])
    server.send_message(msg); server.quit()

# ---------------- NEW: PDF text extraction & marks extraction helpers ----------------
def extract_text_from_pdf_file(path):
    """Try pdfplumber then OCR fallback via pytesseract."""
    text_parts = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                try:
                    t = page.extract_text() or ""
                except Exception:
                    t = ""
                if t and t.strip():
                    text_parts.append(t)
    except Exception:
        # fall through to OCR later
        pass
    combined = "\n".join(text_parts).strip()
    if combined:
        return combined

    # OCR fallback
    ocr_texts = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                try:
                    img = page.to_image(resolution=150).original
                    ocr = pytesseract.image_to_string(img)
                    ocr_texts.append(ocr)
                except Exception:
                    continue
    except Exception:
        # final fallback: try reading as image via PIL (rare)
        try:
            im = Image.open(path)
            ocr_texts.append(pytesseract.image_to_string(im))
        except Exception:
            pass
    return "\n".join(ocr_texts).strip()

def find_hallticket_in_text(text):
    if not text: return None
    candidates = HALL_IN_PDF_RE.findall(text)
    if not candidates: return None
    # prefer longest numeric
    return max(candidates, key=len)

def find_marks_or_absent_in_text(text):
    if not text: return None
    txt = text.replace('\xa0', ' ')
    lower = txt.lower()
    # absent first
    if ABSENT_PATTERN.search(lower):
        return 'Absent'
    for line in lower.splitlines():
        if 'absent' in line or 'not present' in line:
            return 'Absent'
    # labeled marks
    m = MARKS_PATTERN.search(txt)
    if m:
        try:
            v = int(m.group(1))
            if 0 <= v <= 100:
                return v
        except:
            pass
    # fallback numbers 0-100: prefer lines with mark/score words
    candidates = []
    for line in txt.splitlines():
        nums = re.findall(r"\b(\d{1,3})\b", line)
        for n in nums:
            v = int(n)
            if 0 <= v <= 100:
                weight = 2 if re.search(r'(?:marks|mark|score|total|obtained)', line, re.IGNORECASE) else 1
                candidates.append((weight, v))
    if candidates:
        maxw = max(c[0] for c in candidates)
        last = [c[1] for c in candidates if c[0] == maxw][-1]
        return int(last)
    return None

# ---------------- UI: Sidebar (existing) ----------------
with st.sidebar:
    st.header("SMTP Settings")
    smtp_host = st.text_input("SMTP host", "smtp.hostinger.com")
    smtp_port = st.number_input("SMTP port", value=465)
    smtp_use_ssl = st.checkbox("Use SSL (SMTPS)", value=True)
    smtp_sender = st.text_input("Sender email", "info@aiclex.in")
    smtp_password = st.text_input("Sender password", type="password")
    st.markdown("---")
    subject_template = st.text_input("Subject (use {location}, {part}, {total})", "Hall Tickets — {location} (Part {part}/{total})")
    body_template = st.text_area("Body (use {location}, {part}, {total}, {footer})", 
        "Dear Coordinator,\n\nAttached are the hall tickets for {location} (Part {part} of {total}).\n\n{footer}")
    footer_text = st.text_input("Footer text", "Regards,\nAiclex Technologies")
    delay_seconds = st.number_input("Delay between sends (s)", value=2.0, step=0.5)
    attachment_limit_mb = st.number_input("Attachment limit (MB)", value=3.0, step=0.5)
    test_mode = st.checkbox("Test Mode (redirect to test email)", value=True)
    test_email = st.text_input("Test Email", "info@aiclex.in")

# ---------------- Upload ----------------
st.header("1) Upload Excel & ZIP")
uploaded_excel = st.file_uploader("Upload Excel", type=["xlsx","csv"])
uploaded_zip = st.file_uploader("Upload ZIP", type=["zip"])

if uploaded_excel and uploaded_zip:
    # Load Excel
    if uploaded_excel.name.endswith(".csv"):
        df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
    else:
        df = pd.read_excel(uploaded_excel, dtype=str).fillna("")

    cols = df.columns.tolist()
    # select columns (same as before)
    ht_col = st.selectbox("Hallticket column", cols)
    email_col = st.selectbox("Email column", cols)
    loc_col = st.selectbox("Location column", cols)

    # Extract PDFs to /tmp nested
    pdfs, root = extract_zip_bytes_recursively(uploaded_zip.read(), "/tmp")
    st.success(f"Extracted {len(pdfs)} PDFs from ZIP (root {root})")

    # ------- NEW: extract info from PDFs (hallticket in content, marks/absent) -------
    st.info("Scanning PDFs for hallticket & marks (this may take time if OCR is used)...")
    pdf_info_list = []  # each item: {path, hallticket_from_filename, hallticket_from_pdf, marks, text_snippet}
    for ppath in pdfs:
        try:
            text = extract_text_from_pdf_file(ppath)
        except Exception:
            text = ""
        hall_from_filename = extract_hallticket_from_filename(ppath)
        hall_from_pdf = find_hallticket_in_text(text)
        marks = find_marks_or_absent_in_text(text)
        pdf_info_list.append({
            "path": ppath,
            "hall_filename": hall_from_filename,
            "hall_pdf": hall_from_pdf,
            "marks": marks,
            "text": text[:2000]  # keep small snippet
        })

    # build lookup: prefer hallticket from pdf content, fallback to filename
    pdf_lookup = {}
    pdf_marks_map = {}
    for info in pdf_info_list:
        key = info["hall_pdf"] or info["hall_filename"]
        if not key:
            continue
        pdf_lookup.setdefault(key, []).append(info["path"])
        # store marks if found (prefer pdf content marks)
        if info["marks"] is not None:
            pdf_marks_map[key] = info["marks"]

    st.info(f"Found {len(pdf_lookup)} unique hallticket entries inside PDFs (filename+content combined).")

    # ------- NEW: fill df marks & status columns using pdf_marks_map -------
    # ensure marks and status columns exist
    marks_col_name = "marks"
    status_col_name = "status"
    if marks_col_name not in df.columns:
        df[marks_col_name] = ""
    if status_col_name not in df.columns:
        df[status_col_name] = ""

    filled_count = 0
    unmatched_rows = []
    for idx, row in df.iterrows():
        ht_val = str(row.get(ht_col, "")).strip()
        if not ht_val:
            unmatched_rows.append({"index": idx, "reason": "no_hallticket"})
            continue
        # try exact, then digits-normalized
        mval = None
        if ht_val in pdf_marks_map:
            mval = pdf_marks_map[ht_val]
        else:
            digits = re.sub(r"\D", "", ht_val)
            if digits and digits in pdf_marks_map:
                mval = pdf_marks_map[digits]
            else:
                # try to find any key that endswith the digits
                found = None
                for k in pdf_marks_map.keys():
                    if digits and (k.endswith(digits) or digits.endswith(k)):
                        found = k; break
                if found:
                    mval = pdf_marks_map[found]
        if mval is not None:
            df.at[idx, marks_col_name] = mval
            if isinstance(mval, str) and mval.lower().startswith("abs"):
                df.at[idx, status_col_name] = "Absent"
            else:
                try:
                    mm = int(mval)
                    df.at[idx, status_col_name] = "Fail" if mm < 49 else "Pass"
                except:
                    df.at[idx, status_col_name] = "Unknown"
            filled_count += 1
        else:
            unmatched_rows.append({"index": idx, "hallticket": ht_val})

    st.success(f"Filled marks for {filled_count} rows from PDFs")
    if unmatched_rows:
        st.warning(f"{len(unmatched_rows)} rows not matched to any PDF (see console / preview)")

    # show preview of updated df
    st.markdown("### Preview of updated Excel (first 50 rows)")
    st.dataframe(df.head(50))

    # ------ NOW proceed with your existing grouping & ZIP logic (mostly unchanged) ------
    # Group by location
    grouped = defaultdict(lambda: {"files": [], "recipients": set()})
    for _, r in df.iterrows():
        ht = str(r.get(ht_col,"")).strip()
        loc = str(r.get(loc_col,"")).strip()
        raw_emails = str(r.get(email_col,"")).strip()
        if raw_emails:
            for p in re.split(r"[;, \n]+", raw_emails):
                if p.strip() and EMAIL_RE.match(p.strip()):
                    grouped[loc]["recipients"].add(p.strip())
        # prefer pdf_lookup from inside pdf, else filename
        if ht in pdf_lookup:
            # append all matched file paths
            grouped[loc]["files"].extend(pdf_lookup[ht])
        else:
            # try filename match fallback
            fname_match = extract_hallticket_from_filename
            # if any pdf filename contains ht as digits, include
            for p in pdfs:
                fht = extract_hallticket_from_filename(p)
                if fht and fht == ht:
                    grouped[loc]["files"].append(p)

    st.subheader("Summary")
    rows = []
    for loc, info in grouped.items():
        rows.append({"Location": loc, "Recipients": ", ".join(info["recipients"]),
                     "Files": len(info["files"]),
                     "Total Size": human_bytes(sum(os.path.getsize(f) for f in info["files"]))})
    st.dataframe(pd.DataFrame(rows))

    # Prepare ZIPs with splitting
    if st.button("Prepare ZIPs"):
        zip_dir = tempfile.mkdtemp(prefix="aiclex_zips_", dir="/tmp")
        max_bytes = int(attachment_limit_mb * 1024 * 1024)
        prepared = {}
        for loc, info in grouped.items():
            if info["files"]:
                safe_loc = re.sub(r"[^A-Za-z0-9]+", "_", loc)[:50]
                zips = create_split_zips(info["files"], zip_dir, safe_loc, max_bytes)
                prepared[loc] = zips
        st.session_state["prepared"] = prepared
        st.success("ZIPs prepared with splitting")

    if "prepared" in st.session_state:
        st.subheader("Prepared ZIPs")
        preview = []
        for loc, zips in st.session_state["prepared"].items():
            for i, zp in enumerate(zips, start=1):
                preview.append({"Location": loc, "Part": i, "Zip": os.path.basename(zp), "Size": human_bytes(os.path.getsize(zp))})
        st.dataframe(pd.DataFrame(preview))

        if st.button("📤 Send Emails"):
            smtp_cfg = {"host": smtp_host, "port": int(smtp_port), "use_ssl": smtp_use_ssl, "sender": smtp_sender, "password": smtp_password}
            logs, total = [], sum(len(z) for z in st.session_state["prepared"].values())
            done, progress = 0, st.progress(0)
            for loc, zips in st.session_state["prepared"].items():
                recips = list(grouped[loc]["recipients"])
                if test_mode: recips = [test_email]
                total_parts = len(zips)
                for i, zp in enumerate(zips, start=1):
                    subj = subject_template.format(location=loc, part=i, total=total_parts)
                    body = body_template.format(location=loc, part=i, total=total_parts, footer=footer_text)
                    try:
                        send_email_smtp(smtp_cfg, recips, subj, body, [zp])
                        logs.append({"Location": loc, "Recipients": ", ".join(recips), "Zip": os.path.basename(zp), "Status": f"Sent Part {i}/{total_parts}"})
                        st.success(f"Sent {loc} Part {i}/{total_parts} → {', '.join(recips)}")
                    except Exception as e:
                        logs.append({"Location": loc, "Recipients": ", ".join(recips), "Zip": os.path.basename(zp), "Status": f"Failed: {e}"})
                        st.error(f"Failed {loc} Part {i}/{total_parts}: {e}")
                    done += 1; progress.progress(done/total)
                    time.sleep(delay_seconds)
            st.subheader("Logs")
            st.dataframe(pd.DataFrame(logs))

# End of file
