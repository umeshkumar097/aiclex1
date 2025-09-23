# streamlit_app.py
"""
Aiclex — Result Showing
Improved OCR (pdf2image + pytesseract fallback), marks/status filling, location ZIP + Gmail sending.
"""

import os
import io
import re
import time
import zipfile
import tempfile
import smtplib
from email.message import EmailMessage
from collections import defaultdict
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import pytesseract

# Optional pdf2image import will be used if available
try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE_AVAILABLE = True
except Exception:
    PDF2IMAGE_AVAILABLE = False

# ---------------- Config ----------------
MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024  # 3 MB
st.set_page_config(page_title="Aiclex — Result Showing", layout="wide")
st.title("📊 Aiclex — Result Showing")

# ---------------- Sidebar: OCR & SMTP settings ----------------
st.sidebar.header("OCR & SMTP Settings")
tess_path = st.sidebar.text_input("Tesseract path (leave blank to use system tesseract)", value=os.environ.get("TESSERACT_CMD",""))
ocr_lang = st.sidebar.text_input("OCR language (tesseract), e.g. 'eng' or 'eng+hin'", value="eng")
ocr_dpi = st.sidebar.number_input("OCR DPI (pdf2image)", min_value=100, max_value=400, value=200)
show_ocr_debug = st.sidebar.checkbox("Show OCR debug snippets", value=False)
# SMTP defaults (user may override in UI fields too)
st.sidebar.markdown("---")
st.sidebar.markdown("**SMTP defaults for Gmail** (used in UI below)")
st.sidebar.markdown("Host: smtp.gmail.com  Port: 587 (TLS)")

# apply tesseract path if set
if tess_path:
    pytesseract.pytesseract.tesseract_cmd = tess_path

# ---------------- Regex ----------------
HALL_RE = re.compile(r"\b[0-9]{4,}\b")
ABSENT_RE = re.compile(r"\b(absent|not present|a\s*b\s*s\s*e\s*n\s*t)\b", re.IGNORECASE)
MARKS_RE = re.compile(r"(?:marks|mark|score|total|obtained)[:\s\-]*([0-9]{1,3})", re.IGNORECASE)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# ---------------- Utility functions ----------------
def human_bytes(n):
    try:
        n = float(n)
    except:
        return ""
    for unit in ["B","KB","MB","GB"]:
        if n < 1024: return f"{n:.2f} {unit}"
        n /= 1024
    return f"{n:.2f} TB"

# ---------------- OCR / PDF extraction ----------------
def extract_text_from_pdf_bytes(pdf_bytes, dpi=200, lang="eng"):
    """
    Robust PDF to text:
    1) Try pdfplumber text extraction.
    2) If empty, try pdf2image.convert_from_bytes -> pytesseract on images.
    3) If pdf2image not available, try pdfplumber page.to_image() then OCR.
    Returns combined text.
    """
    # 1) pdfplumber direct text
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

    # 2) Use pdf2image -> pytesseract (preferred fallback)
    if PDF2IMAGE_AVAILABLE:
        try:
            pil_pages = convert_from_bytes(pdf_bytes, dpi=dpi)
            ocr_pages = []
            for im in pil_pages:
                try:
                    txt = pytesseract.image_to_string(im, lang=lang)
                except Exception:
                    txt = pytesseract.image_to_string(im)
                ocr_pages.append(txt or "")
            final = "\n".join(ocr_pages).strip()
            if final:
                return final
        except Exception:
            # fall through to next option
            pass

    # 3) pdfplumber page.to_image() -> OCR
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            ocr_pages = []
            for page in pdf.pages:
                try:
                    pil = page.to_image(resolution=dpi).original
                    try:
                        txt = pytesseract.image_to_string(pil, lang=lang)
                    except Exception:
                        txt = pytesseract.image_to_string(pil)
                    ocr_pages.append(txt or "")
                except Exception:
                    continue
            final = "\n".join(ocr_pages).strip()
            if final:
                return final
    except Exception:
        pass

    # 4) last resort: try opening bytes via PIL (rare)
    try:
        im = Image.open(io.BytesIO(pdf_bytes))
        try:
            txt = pytesseract.image_to_string(im, lang=lang)
        except Exception:
            txt = pytesseract.image_to_string(im)
        return (txt or "").strip()
    except Exception:
        return ""


def find_hallticket_in_text(text):
    if not text:
        return None
    candidates = HALL_RE.findall(text)
    if not candidates:
        return None
    return max(candidates, key=len)

def find_marks_or_absent_in_text(text):
    """Return int mark, or 'Absent' string, or None."""
    if not text:
        return None
    txt = text.replace("\xa0"," ")
    lower = txt.lower()
    # absent detection first
    if ABSENT_RE.search(lower):
        return "Absent"
    # labeled marks
    m = MARKS_RE.search(txt)
    if m:
        try:
            v = int(m.group(1))
            if 0 <= v <= 100:
                return v
        except:
            pass
    # fallback: any number 0-100, prefer lines mentioning mark/score/obtained
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

# ---------------- ZIP processing (nested) ----------------
def process_uploaded_zip_bytes(zip_bytes, ocr_dpi=200, ocr_lang="eng", show_debug=False):
    """
    Recursively process zip bytes, return list of dicts:
    {'hallticket':..., 'marks':..., 'pdf_bytes':..., 'pdf_name':...}
    """
    results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as outer:
            for name in outer.namelist():
                if name.endswith('/'): continue
                data = outer.read(name)
                if name.lower().endswith('.zip'):
                    nested = process_uploaded_zip_bytes(data, ocr_dpi, ocr_lang, show_debug)
                    results.extend(nested)
                elif name.lower().endswith('.pdf'):
                    # extract text with improved OCR
                    text = extract_text_from_pdf_bytes(data, dpi=ocr_dpi, lang=ocr_lang)
                    hall = find_hallticket_in_text(text)
                    marks = find_marks_or_absent_in_text(text)
                    results.append({
                        "hallticket": str(hall).strip() if hall else "",
                        "marks": marks,
                        "pdf_bytes": data,
                        "pdf_name": os.path.basename(name),
                        "text_snippet": text[:2000] if show_debug else ""
                    })
                else:
                    # ignore other files
                    continue
    except zipfile.BadZipFile:
        pass
    return results

# ---------------- Excel fill logic ----------------
def fill_excel_with_pdf_data(df, pdf_data, hall_col_name):
    """
    Fill df['marks'] and df['status'] using pdf_data (list of dicts).
    Rules:
      - prefer existing Excel marks (if present)
      - else use PDF marks if present
      - if mark found: if >49 => Pass else Fail (49 counts as Fail)
      - if no mark anywhere or PDF says 'Absent' => Absent
    """
    # build pdf map: hallticket -> marks (int or 'Absent')
    pdf_map = {}
    for p in pdf_data:
        k = str(p.get("hallticket") or "").strip()
        if not k: continue
        pdf_map[k] = p.get("marks")

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
        ht_val = str(row.get(hall_col_name, "")).strip()
        if not ht_val:
            unmatched.append({"index": idx, "reason": "no_hallticket"})
            continue

        # 1) prefer existing Excel marks (if non-empty)
        excel_raw = row.get(marks_col, "")
        excel_mark = None
        if pd.notna(excel_raw) and str(excel_raw).strip() != "":
            # try parse int, otherwise keep string
            try:
                excel_mark = int(re.sub(r"\D", "", str(excel_raw)))
            except:
                excel_mark = str(excel_raw).strip()

        # 2) pdf_mark
        pdf_mark = None
        if ht_val in pdf_map:
            pdf_mark = pdf_map[ht_val]
        else:
            digits = re.sub(r"\D", "", ht_val)
            if digits and digits in pdf_map:
                pdf_mark = pdf_map[digits]
            else:
                # try endswith/contains heuristic
                for k in pdf_map.keys():
                    kd = re.sub(r"\D", "", str(k))
                    if kd and (k.endswith(digits) or digits.endswith(kd) or kd.endswith(digits)):
                        pdf_mark = pdf_map[k]
                        break

        # 3) decide final_mark: excel preferred
        final_mark = excel_mark if excel_mark is not None else pdf_mark

        # 4) set marks & status based on rules
        if final_mark is None or (isinstance(final_mark, str) and re.search(r'abs', str(final_mark), re.IGNORECASE)):
            df.at[idx, marks_col] = ""
            df.at[idx, status_col] = "Absent"
        else:
            # numeric case
            try:
                mm = int(final_mark)
                df.at[idx, marks_col] = mm
                df.at[idx, status_col] = "Pass" if mm > 49 else "Fail"
            except:
                df.at[idx, marks_col] = ""
                df.at[idx, status_col] = "Absent"

        filled += 1

    return df, filled, unmatched

# ---------------- Zip creation & splitting ----------------
def make_zip_bytes(file_entries):
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, content in file_entries:
            zf.writestr(fname, content)
    bio.seek(0)
    return bio.read()

def split_files_into_zip_parts(file_entries, max_bytes=MAX_ATTACHMENT_BYTES, zip_name_prefix="results"):
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
            # if single file too big, put it alone (may exceed limit)
            current = [(fname, b)]
            test_zip2 = make_zip_bytes(current)
            if len(test_zip2) > max_bytes:
                parts.append((f"{zip_name_prefix}_part{len(parts)+1}.zip", test_zip2))
                current = []
    if current:
        parts.append((f"{zip_name_prefix}_part{len(parts)+1}.zip", make_zip_bytes(current)))
    return parts

# ---------------- Email sending (Gmail) ----------------
def send_email_with_attachments_gmail(smtp_user, smtp_pass, to_email, subject, body, attachments):
    msg = EmailMessage()
    msg["From"] = smtp_user
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)
    for fname, b in attachments:
        msg.add_attachment(b, maintype="application", subtype="zip", filename=fname)
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(smtp_user, smtp_pass)
            s.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

# ---------------- Streamlit UI ----------------
st.write("Upload Excel/CSV and a ZIP (may contain nested ZIPs with PDFs).")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("Excel/CSV file", type=["xlsx", "xls", "csv"])
with col2:
    uploaded_zip = st.file_uploader("ZIP file (nested zips with PDFs)", type=["zip"])

process_button = st.button("Process & Prepare Results")

if process_button:
    if not uploaded_excel or not uploaded_zip:
        st.error("Please upload both Excel/CSV and ZIP file.")
        st.stop()

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

    # pick columns
    cols = df.columns.tolist()
    hall_col = st.selectbox("Hallticket column", cols)
    email_col = st.selectbox("Email column", cols)
    loc_col = st.selectbox("Location column", cols)

    # process zip -> extract pdf data
    with st.spinner("Processing ZIPs & running OCR (may take time)..."):
        pdf_data = process_uploaded_zip_bytes(uploaded_zip.read(), ocr_dpi, ocr_lang, show_ocr_debug)

    st.info(f"PDF records extracted: {len(pdf_data)}")

    if show_ocr_debug:
        st.subheader("OCR Debug (first 200 chars of each PDF text)")
        debug_rows = []
        for p in pdf_data:
            debug_rows.append({"pdf_name": p.get("pdf_name"), "hallticket": p.get("hallticket"), "marks": p.get("marks"), "text_snippet": (p.get("text_snippet") or "")[:200]})
        st.dataframe(pd.DataFrame(debug_rows))

    # fill excel using pdf_data
    updated_df, filled_rows, unmatched = fill_excel_with_pdf_data(df, pdf_data, hall_col)
    st.success(f"Filled {filled_rows} rows from PDFs")
    if unmatched:
        st.warning(f"{len(unmatched)} rows had missing hallticket (see preview)")

    st.markdown("### Preview (first 50 rows)")
    st.dataframe(updated_df.head(50))

    # download updated excel
    out_buf = io.BytesIO()
    try:
        updated_df.to_excel(out_buf, index=False, engine="openpyxl")
        out_buf.seek(0)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button("Download updated Excel", data=out_buf, file_name=f"updated_aiclex_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"Failed to prepare download: {e}")

    # Group PDFs by location -> prepare zip parts per location
    grouped = defaultdict(lambda: {"pdfs": [], "recipients": set()})
    # Build quick map: hallticket -> list of pdf entries
    pdf_map = defaultdict(list)
    for p in pdf_data:
        key = str(p.get("hallticket") or "").strip()
        if key:
            pdf_map[key].append(p)

    for _, row in updated_df.iterrows():
        ht = str(row.get(hall_col, "")).strip()
        loc = str(row.get(loc_col, "")).strip() or "Unknown"
        emails_raw = str(row.get(email_col, "")).strip()
        if emails_raw:
            for e in re.split(r"[;, \n]+", emails_raw):
                if e and EMAIL_RE.match(e.strip()):
                    grouped[loc]["recipients"].add(e.strip())
        # attach pdfs for this hallticket
        if ht and ht in pdf_map:
            for p in pdf_map[ht]:
                grouped[loc]["pdfs"].append(p)

    # Show summary
    st.markdown("### Location summary")
    summary = []
    for loc, info in grouped.items():
        total_size = sum(len(p["pdf_bytes"]) for p in info["pdfs"])
        summary.append({"Location": loc, "Recipients": len(info["recipients"]), "Files": len(info["pdfs"]), "Size": human_bytes(total_size)})
    st.dataframe(pd.DataFrame(summary))

    # Prepare ZIPs (split if needed) and email UI
    if st.button("Prepare ZIPs for sending"):
        prepared = {}
        for loc, info in grouped.items():
            if not info["pdfs"]:
                prepared[loc] = []
                continue
            files_for_zip = []
            for p in info["pdfs"]:
                fname = f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}"
                files_for_zip.append((fname, p["pdf_bytes"]))
            parts = split_files_into_zip_parts(files_for_zip, MAX_ATTACHMENT_BYTES, zip_name_prefix=re.sub(r"[^A-Za-z0-9]+","_",loc)[:40])
            prepared[loc] = parts
        st.session_state["prepared_zips"] = prepared
        st.success("Prepared ZIP parts per location (in memory)")

    if "prepared_zips" in st.session_state:
        st.subheader("Prepared ZIP parts (sample)")
        preview = []
        for loc, parts in st.session_state["prepared_zips"].items():
            for i, (fname, b) in enumerate(parts, start=1):
                preview.append({"Location": loc, "Part": i, "ZipName": fname, "Size": human_bytes(len(b))})
        st.dataframe(pd.DataFrame(preview))

        st.markdown("---")
        st.subheader("Send Emails (Gmail)")

        smtp_user = st.text_input("Gmail address (SMTP user)", value=os.environ.get("SMTP_USER",""))
        smtp_pass = st.text_input("Gmail App Password (SMTP pass)", type="password", value=os.environ.get("SMTP_PASS",""))
        test_mode = st.checkbox("Test mode (send all to test email)", value=True)
        test_email = st.text_input("Test email (if test mode on)", value=os.environ.get("TEST_EMAIL","info@aiclex.in"))
        delay_secs = st.number_input("Delay between sends (seconds)", value=1.0, step=0.5)

        if st.button("Start sending emails"):
            if not smtp_user or not smtp_pass:
                st.error("Provide Gmail address and app password (SMTP_USER & SMTP_PASS).")
            else:
                total_recipients = sum(len(grouped[loc]["recipients"]) for loc in grouped)
                # if test_mode, send to single test address per part
                progress = st.progress(0)
                status = st.empty()
                sent = 0
                failed = []
                total_parts = sum(len(parts) for parts in st.session_state["prepared_zips"].values())
                if total_parts == 0:
                    st.warning("No ZIP parts prepared to send.")
                else:
                    overall_index = 0
                    for loc, parts in st.session_state["prepared_zips"].items():
                        recips = list(grouped[loc]["recipients"])
                        if test_mode:
                            recips = [test_email]
                        if not recips:
                            # still may want to log or skip
                            continue
                        for part_index, (zipname, zipbytes) in enumerate(parts, start=1):
                            subj = f"Results — {loc} (Part {part_index}/{len(parts)})"
                            body = f"Dear Participant,\n\nPlease find attached the results for {loc} (Part {part_index}/{len(parts)}).\n\nRegards,\nAiclex"
                            # send to each recipient (we send same attachments to all in that location)
                            for r in recips:
                                ok, err = send_email_with_attachments_gmail(smtp_user, smtp_pass, r, subj, body, [(zipname, zipbytes)])
                                overall_index += 1
                                progress.progress(min(1.0, overall_index / max(1, total_parts * max(1,len(recips)))))
                                status.text(f"Sent to {r} — {loc} Part {part_index}/{len(parts)} ({overall_index})")
                                if ok:
                                    sent += 1
                                else:
                                    failed.append({"email": r, "loc": loc, "zip": zipname, "error": err})
                                time.sleep(delay_secs)
                st.success(f"Done. Sent: {sent}. Failed: {len(failed)}")
                if failed:
                    st.error("Some sends failed (sample):")
                    st.dataframe(pd.DataFrame(failed).head(50))

# End of file
