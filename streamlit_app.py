# streamlit_app.py
"""
Aiclex — Result Showing (final combined)
v2: SQLite PDF cache + live send UI + persistent send logs
"""

import os
import io
import re
import time
import sqlite3
import hashlib
import zipfile
import logging
import smtplib
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
APP_TITLE   = "CRUX — Result Showing"
BRAND       = "Aiclex Technologies"
DEFAULT_OCR_DPI         = 200
DEFAULT_OCR_LANG        = "eng"
DEFAULT_ATTACHMENT_MB   = 3.0

# SQLite DB path — stored next to the script so it persists across runs
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "aiclex_cache.db")

# logging
logger = logging.getLogger("aiclex")
if not logger.handlers:
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(ch)
logger.setLevel(logging.INFO)

# ---------------- SQLite helpers ----------------
def get_db():
    """Open a new SQLite connection with WAL mode and a generous timeout."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def init_db():
    conn = get_db()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS pdf_cache (
                pdf_hash    TEXT PRIMARY KEY,
                pdf_name    TEXT,
                hallticket  TEXT,
                marks       TEXT,
                status      TEXT,
                processed_at TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS send_log (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp        TEXT,
                recipient_email  TEXT,
                name             TEXT,
                hallticket       TEXT,
                location         TEXT,
                zip_name         TEXT,
                status           TEXT,
                error            TEXT
            )
        """)
        conn.commit()
    finally:
        conn.close()

def pdf_hash(pdf_bytes: bytes) -> str:
    return hashlib.md5(pdf_bytes).hexdigest()

def cache_get(h: str):
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT pdf_name, hallticket, marks, status FROM pdf_cache WHERE pdf_hash=?", (h,)
        ).fetchone()
    finally:
        conn.close()
    if row:
        return {
            "pdf_name":     row[0],
            "hallticket":   row[1],
            "marks":        int(row[2]) if (row[2] and row[2].lstrip("-").isdigit()) else row[2],
            "status":       row[3],
            "pdf_bytes":    None,
            "text_snippet": "",
            "_from_cache":  True
        }
    return None

def cache_put(h: str, result: dict):
    marks_val = str(result.get("marks", "")) if result.get("marks") is not None else ""
    conn = get_db()
    try:
        conn.execute(
            """INSERT OR REPLACE INTO pdf_cache
               (pdf_hash, pdf_name, hallticket, marks, status, processed_at)
               VALUES (?,?,?,?,?,?)""",
            (h, result.get("pdf_name",""), result.get("hallticket",""),
             marks_val, result.get("status",""), datetime.now().isoformat())
        )
        conn.commit()
    finally:
        conn.close()

def log_send(recipient_email, name, hallticket, location, zip_name, status, error=""):
    conn = get_db()
    try:
        conn.execute(
            """INSERT INTO send_log
               (timestamp, recipient_email, name, hallticket, location, zip_name, status, error)
               VALUES (?,?,?,?,?,?,?,?)""",
            (datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
             recipient_email, name, hallticket, location, zip_name, status, error)
        )
        conn.commit()
    finally:
        conn.close()

def load_send_logs() -> pd.DataFrame:
    conn = get_db()
    try:
        df = pd.read_sql_query("SELECT * FROM send_log ORDER BY id DESC LIMIT 5000", conn)
    finally:
        conn.close()
    return df

def cache_stats():
    conn = get_db()
    try:
        total = conn.execute("SELECT COUNT(*) FROM pdf_cache").fetchone()[0]
    finally:
        conn.close()
    return total

# Init DB on startup
init_db()

# ---------------- Patterns ----------------
LABEL_RE     = re.compile(r"Marks\s*Obtained", re.IGNORECASE)
MARKS_NUM_RE = re.compile(r"\b([0-9]{1,3})\b")
ABSENT_RE    = re.compile(r"\b(absent|not present)\b", re.IGNORECASE)
PASSFAIL_RE  = re.compile(r"([0-9]{1,3})\s*(PASS|FAIL)", re.IGNORECASE)
HALL_RE      = re.compile(r"\b[0-9]{3,}\b")
EMAIL_RE     = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# ---------------- Streamlit UI setup ----------------
st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon=None)

st.markdown(
    f"""
    <div style='padding:1.2rem 1.5rem; background:linear-gradient(90deg,#0b74de 0%,#0550a0 100%);
                border-radius:10px; margin-bottom:0.5rem;'>
        <h1 style='color:#ffffff; margin:0; font-size:2rem; letter-spacing:0.5px;'>{APP_TITLE}</h1>
        <p style='color:#cce0ff; margin:0.25rem 0 0; font-size:0.85rem;'>Built by {BRAND}</p>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style='display:flex; gap:0.6rem; margin:0.8rem 0 1rem;'>
        <div style='flex:1; background:#f0f6ff; border-left:4px solid #0b74de;
                    padding:0.5rem 0.75rem; border-radius:4px; font-size:0.82rem; color:#333;'>
            <b>Step 1</b><br>Upload Excel &amp; ZIP
        </div>
        <div style='flex:1; background:#f0f6ff; border-left:4px solid #0b74de;
                    padding:0.5rem 0.75rem; border-radius:4px; font-size:0.82rem; color:#333;'>
            <b>Step 2</b><br>Process &amp; Preview
        </div>
        <div style='flex:1; background:#f0f6ff; border-left:4px solid #0b74de;
                    padding:0.5rem 0.75rem; border-radius:4px; font-size:0.82rem; color:#333;'>
            <b>Step 3</b><br>Prepare ZIPs
        </div>
        <div style='flex:1; background:#f0f6ff; border-left:4px solid #0b74de;
                    padding:0.5rem 0.75rem; border-radius:4px; font-size:0.82rem; color:#333;'>
            <b>Step 4</b><br>Send Emails
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

cached_count = cache_stats()
st.markdown(
    f"<div style='background:#eaf4fb; border:1px solid #b3d9f0; border-radius:6px; "
    f"padding:0.4rem 0.8rem; font-size:0.82rem; color:#1a5276; margin-bottom:0.5rem;'>"
    f"SQLite Cache &nbsp;|&nbsp; <b>{cached_count}</b> PDFs already cached — OCR will be skipped for these."
    f"</div>",
    unsafe_allow_html=True
)

# ---------------- Sidebar config ----------------
st.sidebar.markdown(
    "<div style='font-size:1rem; font-weight:700; color:#0b74de; "
    "padding-bottom:0.3rem; border-bottom:2px solid #0b74de; margin-bottom:0.6rem;'>"
    "Settings</div>",
    unsafe_allow_html=True
)
st.sidebar.markdown("**OCR Configuration**")
tesseract_path      = st.sidebar.text_input("Tesseract path (optional)", value=os.environ.get("TESSERACT_CMD",""))
ocr_lang            = st.sidebar.text_input("OCR language (e.g. eng or eng+hin)", value=DEFAULT_OCR_LANG)
ocr_dpi             = st.sidebar.number_input("OCR DPI (pdf2image)", value=int(DEFAULT_OCR_DPI), min_value=100, max_value=400, step=10)
st.sidebar.markdown("**Email Configuration**")
attachment_limit_mb = st.sidebar.number_input("Attachment limit (MB)", value=float(DEFAULT_ATTACHMENT_MB), step=0.5)
send_delay          = st.sidebar.number_input("Delay between sends (s)", value=1.0, step=0.5)
show_ocr_debug      = st.sidebar.checkbox("Show OCR debug snippet", value=False)
st.sidebar.caption("System packages required: poppler-utils, tesseract-ocr")
st.sidebar.markdown("---")
if st.sidebar.button("Clear PDF Cache"):
    with get_db() as conn:
        conn.execute("DELETE FROM pdf_cache")
        conn.commit()
    st.sidebar.success("PDF cache cleared successfully.")

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

class ProcessTracker:
    def __init__(self, total_steps, description="Processing", show_ui=True):
        self.total_steps  = max(1, total_steps)
        self.current_step = 0
        self.description  = description
        self.start_time   = time.time()
        self.show_ui      = show_ui
        if show_ui:
            self.progress_bar = st.progress(0)
            self.status       = st.empty()
        else:
            self.progress_bar = None
            self.status       = None

    def update(self, step_desc):
        self.current_step += 1
        progress = min(1.0, self.current_step / self.total_steps)
        elapsed  = time.time() - self.start_time
        if self.current_step > 1:
            eta      = (elapsed / self.current_step) * (self.total_steps - self.current_step)
            eta_text = f"ETA: {int(eta)}s"
        else:
            eta_text = "Calculating ETA..."
        if self.show_ui:
            self.progress_bar.progress(progress)
            self.status.write(
                f"{self.description}: {step_desc} "
                f"({min(self.current_step, self.total_steps)}/{self.total_steps}) - {eta_text}"
            )

    def done(self, message="Processing complete!"):
        if self.show_ui:
            self.progress_bar.progress(1.0)
            elapsed = time.time() - self.start_time
            self.status.write(f"{message} (took {int(elapsed)}s)")

def is_pdf_bytes(b: bytes) -> bool:
    try:
        return bool(b) and b.lstrip().startswith(b"%PDF")
    except Exception:
        return False

# OCR / text extraction — LOGIC UNCHANGED
def extract_text_from_pdf_bytes(pdf_bytes: bytes, dpi: int = DEFAULT_OCR_DPI, lang: str = DEFAULT_OCR_LANG) -> str:
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

    if PDF2IMAGE:
        try:
            pages     = convert_from_bytes(pdf_bytes, dpi=dpi)
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

    try:
        im = Image.open(io.BytesIO(pdf_bytes))
        try:
            return pytesseract.image_to_string(im, lang=lang)
        except Exception:
            return pytesseract.image_to_string(im)
    except Exception:
        return ""


def _extract_hallticket_from_filename(fname: str) -> str:
    """
    Extract hallticket from filename like admit-card-1036-29-803038629.pdf
    Rule: take the LAST 9-digit numeric group from the stem.
    Only falls back to other lengths if no 9-digit group exists.
    """
    stem = os.path.splitext(os.path.basename(fname))[0]
    all_groups = re.findall(r"\d+", stem)
    if not all_groups:
        return ""
    # Prefer exactly 9-digit groups (standard hallticket length)
    nine_digit = [g for g in all_groups if len(g) == 9]
    if nine_digit:
        return nine_digit[-1]
    # Fallback: last group with >=6 digits
    long_groups = [g for g in all_groups if len(g) >= 6]
    if long_groups:
        return long_groups[-1]
    return all_groups[-1]


# parse PDF — cache-aware, stronger filename hallticket fallback
def parse_pdf_bytes(pdf_bytes: bytes, fname: str = "",
                    ocr_dpi: int = DEFAULT_OCR_DPI,
                    ocr_lang_s: str = DEFAULT_OCR_LANG):
    h      = pdf_hash(pdf_bytes)
    cached = cache_get(h)
    if cached is not None:
        cached["pdf_bytes"] = pdf_bytes
        cached["pdf_name"]  = os.path.basename(fname) or cached["pdf_name"]
        return cached

    # Cache miss: full OCR
    text      = extract_text_from_pdf_bytes(pdf_bytes, dpi=ocr_dpi, lang=ocr_lang_s) or ""
    text_norm = text.replace('\xa0', ' ')

    # Hallticket: OCR candidates
    h_cands  = HALL_RE.findall(text_norm)
    hall_ocr = max(h_cands, key=len) if h_cands else ""

    # Filename candidate
    hall_fname = _extract_hallticket_from_filename(fname)

    # Prefer filename when it is longer (more specific) OR when OCR gives nothing
    if hall_fname and (not hall_ocr or len(hall_fname) >= len(hall_ocr)):
        hall = hall_fname
    elif hall_ocr:
        hall = hall_ocr
    else:
        hall = ""

    # Marks / status — LOGIC UNCHANGED
    marks  = None
    status = "Absent"
    if ABSENT_RE.search(text_norm):
        marks  = ""
        status = "Absent"
    else:
        pf = PASSFAIL_RE.search(text_norm)
        if pf:
            try:
                val    = int(pf.group(1))
                marks  = val
                status = "Pass" if val > 49 else "Fail"
            except:
                marks  = ""
                status = "Absent"
        else:
            lbl = LABEL_RE.search(text_norm)
            if lbl:
                snippet = text_norm[lbl.end():lbl.end()+200]
                mnum    = re.search(r"([0-9]{1,3})", snippet)
                if mnum:
                    val    = int(mnum.group(1))
                    marks  = val
                    status = "Pass" if val > 49 else "Fail"
                else:
                    marks  = ""
                    status = "Absent"
            else:
                nums = MARKS_NUM_RE.findall(text_norm)
                nums = [int(n) for n in nums if 0 <= int(n) <= 100]
                if nums:
                    val    = nums[-1]
                    marks  = val
                    status = "Pass" if val > 49 else "Fail"
                else:
                    marks  = ""
                    status = "Absent"

    result = {
        "pdf_name":      os.path.basename(fname),
        "pdf_bytes":     pdf_bytes,
        "hallticket":    str(hall).strip(),
        "marks":         marks,
        "status":        status,
        "text_snippet":  (text_norm[:2000] if show_ocr_debug else ""),
        "_from_cache":   False
    }
    cache_put(h, result)
    return result


# Recursive ZIP extraction — LOGIC UNCHANGED
def extract_from_zip_recursive(zip_bytes: bytes, ocr_dpi: int, ocr_lang_s: str,
                                progress: ProcessTracker = None, nested_call: bool = False):
    results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = zf.namelist()
            total = len(names)
            if progress is None and not nested_call:
                progress = ProcessTracker(total, "Processing ZIP files", show_ui=False)
            for i, name in enumerate(names, start=1):
                if progress and not nested_call:
                    progress.update(f"Processing: {os.path.basename(name)}")
                try:
                    data = zf.read(name)
                except Exception as e:
                    logger.warning("Cannot read entry %s: %s", name, e)
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
            if progress and not nested_call:
                progress.done("ZIP processing complete!")
    except zipfile.BadZipFile:
        raise
    return results


# Fill excel — LOGIC UNCHANGED
def fill_excel_using_pdf_data(df: pd.DataFrame, pdf_data: list, hall_col: str):
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

    marks_col  = "marks"
    status_col = "status"
    if marks_col not in df.columns:
        df[marks_col] = ""
    if status_col not in df.columns:
        df[status_col] = ""

    filled    = 0
    unmatched = []
    total     = len(df)
    progress  = ProcessTracker(total, "Filling Excel data", show_ui=False)

    for i, (idx, row) in enumerate(df.iterrows(), start=1):
        progress.update(f"Processing row {i}")
        ht = str(row.get(hall_col,"")).strip()
        if not ht:
            unmatched.append({"index": idx, "reason": "no_hallticket"})
            continue
        val = None
        # 1) Exact match
        if ht in pdf_map:
            val = pdf_map[ht]
        else:
            # 2) Strip non-digits from Excel hallticket and try exact match
            digits = re.sub(r"\D", "", ht)
            if digits and digits in pdf_map:
                val = pdf_map[digits]
            else:
                # 3) Try matching PDF key's digits exactly against Excel digits
                #    Both must be equal length AND identical — no suffix/prefix tricks
                for k in pdf_map.keys():
                    kd = re.sub(r"\D", "", str(k))
                    if kd and digits and kd == digits:
                        val = pdf_map[k]
                        break
        if val is None:
            df.at[idx, marks_col]  = ""
            df.at[idx, status_col] = "No PDF"
        else:
            m = val.get("marks")
            if isinstance(m, int):
                df.at[idx, marks_col]  = str(int(m))
                df.at[idx, status_col] = "Pass" if int(m) > 49 else "Fail"
            else:
                df.at[idx, marks_col]  = ""
                df.at[idx, status_col] = "Absent "
        filled += 1

    progress.update(f"Processed {i} rows")
    progress.done("Excel filling complete!")
    return df, filled, unmatched, pdf_map


# ZIP helpers — LOGIC UNCHANGED
def make_zip_bytes(file_entries):
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, content in file_entries:
            zf.writestr(fname, content)
    bio.seek(0)
    return bio.read()

def split_files_into_zip_parts(file_entries, max_bytes, zip_name_prefix="results"):
    if not file_entries:
        return []
    parts         = []
    current_files = []
    part_no       = 1
    for fname, content in file_entries:
        if current_files and len(make_zip_bytes(current_files + [(fname, content)])) > max_bytes:
            zip_name = f"{zip_name_prefix}_part{part_no}.zip"
            parts.append((zip_name, make_zip_bytes(current_files)))
            part_no += 1
            current_files = []
        current_files.append((fname, content))
        if len(make_zip_bytes(current_files)) > max_bytes:
            zip_name = f"{zip_name_prefix}_part{part_no}.zip"
            parts.append((zip_name, make_zip_bytes(current_files)))
            part_no += 1
            current_files = []
    if current_files:
        zip_name = f"{zip_name_prefix}_part{part_no}.zip"
        parts.append((zip_name, make_zip_bytes(current_files)))
    return parts


def detect_name_col(df: pd.DataFrame) -> str:
    for c in df.columns:
        if re.fullmatch(r"name|student[\s_]?name|candidate[\s_]?name|full[\s_]?name", c.strip(), re.IGNORECASE):
            return c
    return ""


# ================================================================
# Main UI
# ================================================================
st.markdown(
    "<div style='margin-top:1rem; padding:0.5rem 1rem; background:#0b74de; color:#fff; "
    "border-radius:6px; font-size:1rem; font-weight:600;'>Step 1 — Upload Files</div>",
    unsafe_allow_html=True
)
st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("Upload Excel or CSV", type=["xlsx","csv"])
with col2:
    uploaded_zip = st.file_uploader("Upload ZIP (nested zips with PDFs)", type=["zip"])

if uploaded_excel and uploaded_zip:
    try:
        if 'df' not in st.session_state or st.session_state.get('uploaded_excel_name') != uploaded_excel.name:
            if uploaded_excel.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_excel, dtype=str).fillna("")
            else:
                df = pd.read_excel(uploaded_excel, dtype=str, engine="openpyxl").fillna("")
            st.session_state['df']                  = df
            st.session_state['uploaded_excel_name'] = uploaded_excel.name
        df = st.session_state['df']
    except Exception as e:
        st.error(f"Failed to read Excel/CSV: {e}")
        st.stop()

    st.success(f"Excel loaded — {len(df)} rows")
    cols         = df.columns.tolist()
    hall_col     = st.selectbox("Select Hallticket column", cols)
    email_col    = st.selectbox("Select Email column", cols)
    location_col = st.selectbox("Select Location column", cols)

    auto_name        = detect_name_col(df)
    name_col_options = ["(None)"] + cols
    default_name_idx = name_col_options.index(auto_name) if auto_name in name_col_options else 0
    name_col = st.selectbox(
        "Select Name column (used in live send log)",
        name_col_options,
        index=default_name_idx
    )
    if name_col == "(None)":
        name_col = None

    process_key       = f"{uploaded_excel.name}|{uploaded_zip.name}"
    start_process     = st.button("Start processing (Run OCR & Fill results)")
    already_processed = st.session_state.get("processed_key") == process_key

    if start_process and not already_processed:
        try:
            zip_bytes = uploaded_zip.read()
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
                total_files = len([n for n in zf.namelist()
                                   if n.lower().endswith('.pdf') or n.lower().endswith('.zip')])
            main_progress = ProcessTracker(total_files, "Processing files", show_ui=True)
            pdf_data      = extract_from_zip_recursive(zip_bytes, ocr_dpi=ocr_dpi,
                                                        ocr_lang_s=ocr_lang, progress=main_progress)
            main_progress.done("Processing complete!")
            st.session_state['pdf_data']          = pdf_data
            st.session_state['uploaded_zip_name'] = uploaded_zip.name
            st.session_state['processed_key']     = process_key
            st.markdown(
                f"<div style='background:#eaf4fb; border:1px solid #b3d9f0; border-radius:6px; "
                f"padding:0.4rem 0.8rem; font-size:0.82rem; color:#1a5276;'>"
                f"SQLite Cache updated — <b>{cache_stats()}</b> PDFs now cached.</div>",
                unsafe_allow_html=True
            )
        except zipfile.BadZipFile:
            st.error("Uploaded file is not a valid ZIP archive.")
            st.session_state['pdf_data'] = []
        except Exception as e:
            st.error(f"An error occurred during ZIP processing: {e}")
            st.session_state['pdf_data'] = []

    if st.session_state.get('processed_key') == process_key:
        pdf_data    = st.session_state.get('pdf_data', [])
        cache_hits  = sum(1 for p in pdf_data if p.get("_from_cache"))
        ocr_fresh   = len(pdf_data) - cache_hits
        st.markdown(
            f"<div style='display:flex; gap:0.6rem; margin:0.5rem 0;'>"
            f"<div style='flex:1; background:#f8f9fa; border:1px solid #dee2e6; border-radius:6px; "
            f"padding:0.6rem 1rem; text-align:center;'>"
            f"<div style='font-size:1.4rem; font-weight:700; color:#0b74de;'>{len(pdf_data)}</div>"
            f"<div style='font-size:0.75rem; color:#555;'>Total PDFs</div></div>"
            f"<div style='flex:1; background:#f8f9fa; border:1px solid #dee2e6; border-radius:6px; "
            f"padding:0.6rem 1rem; text-align:center;'>"
            f"<div style='font-size:1.4rem; font-weight:700; color:#1a7a4a;'>{cache_hits}</div>"
            f"<div style='font-size:0.75rem; color:#555;'>From Cache</div></div>"
            f"<div style='flex:1; background:#f8f9fa; border:1px solid #dee2e6; border-radius:6px; "
            f"padding:0.6rem 1rem; text-align:center;'>"
            f"<div style='font-size:1.4rem; font-weight:700; color:#8B5E00;'>{ocr_fresh}</div>"
            f"<div style='font-size:0.75rem; color:#555;'>OCR Processed</div></div>"
            f"</div>",
            unsafe_allow_html=True
        )

        if show_ocr_debug and pdf_data:
            st.subheader("OCR debug (sample snippets)")
            debug_rows = [{"pdf_name": p["pdf_name"], "hallticket": p["hallticket"],
                           "marks": p["marks"], "status": p["status"],
                           "text_snippet": p.get("text_snippet","")[:500]} for p in pdf_data]
            st.dataframe(pd.DataFrame(debug_rows).head(200))

        updated_df, filled_count, unmatched, pdf_map = fill_excel_using_pdf_data(
            df.copy(), pdf_data, hall_col
        )
        st.session_state['updated_df'] = updated_df
        st.session_state['pdf_map']    = pdf_map
        st.success(f"Filled {filled_count} rows (marks/status).")
        if unmatched:
            st.warning(f"{len(unmatched)} rows had missing hallticket.")

        st.subheader("Preview updated results (first 100 rows)")
        st.dataframe(updated_df.head(100))
    else:
        st.info("Files uploaded. Click 'Start processing (Run OCR & Fill results)' to begin OCR and fill the Excel.")

    updated_df = st.session_state.get('updated_df', df).copy()
    pdf_data   = st.session_state.get('pdf_data', [])

    if 'marks' not in updated_df.columns:
        updated_df['marks'] = ""
    if 'status' not in updated_df.columns:
        updated_df['status'] = ""

    sheets       = {}
    total        = len(updated_df)
    pass_count   = int((updated_df['status'] == 'Pass').sum())
    fail_count   = int((updated_df['status'] == 'Fail').sum())
    absent_count = int((updated_df['status'] == 'Absent').sum())
    summary_overall = pd.DataFrame([{
        "Total": total, "Pass": pass_count, "Fail": fail_count, "Absent": absent_count
    }])
    sheets["results"]         = updated_df
    sheets["summary_overall"] = summary_overall
    if location_col in updated_df.columns:
        by_loc = []
        for loc, g in updated_df.groupby(location_col):
            by_loc.append({
                "Location": loc, "Total": len(g),
                "Pass":   int((g['status']=="Pass").sum()),
                "Fail":   int((g['status']=="Fail").sum()),
                "Absent": int((g['status']=="Absent").sum())
            })
        sheets["summary_by_location"] = pd.DataFrame(by_loc)

    out_buf = io.BytesIO()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    out_buf.seek(0)
    st.download_button(
        "Download results + summary (Excel)",
        data=out_buf,
        file_name=f"aiclex_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
    st.markdown(
        "<div style='padding:0.5rem 1rem; background:#0b74de; color:#fff; "
        "border-radius:6px; font-size:1rem; font-weight:600;'>Step 2 — Prepare ZIPs</div>",
        unsafe_allow_html=True
    )
    st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)

    pdf_map_multi = defaultdict(list)
    for p in pdf_data:
        k = str(p.get("hallticket","")).strip()
        if k:
            pdf_map_multi[k].append(p)

    recipients  = defaultdict(lambda: defaultdict(list))
    missing_log = []
    row_meta    = {}   # email_key -> [{name, hallticket}]

    for idx, row in updated_df.iterrows():
        ht         = str(row.get(hall_col,"")).strip()
        emails_raw = str(row.get(email_col,"")).strip()
        loc        = str(row.get(location_col,"")).strip() or "Unknown"
        name_val   = str(row.get(name_col,"")).strip() if name_col else ""

        if not emails_raw:
            continue
        emails_list   = [e.strip() for e in re.split(r"[;, \n]+", emails_raw) if e.strip()]
        if not emails_list:
            continue
        recipient_key = ", ".join(sorted(list(set(emails_list))))

        if recipient_key not in row_meta:
            row_meta[recipient_key] = []
        row_meta[recipient_key].append({"name": name_val, "hallticket": ht})

        found_any = False
        if ht and ht in pdf_map_multi:
            for p in pdf_map_multi[ht]:
                recipients[recipient_key][loc].append(
                    (f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}", p["pdf_bytes"])
                )
            found_any = True
        else:
            # Strict exact match: strip non-digits from both sides, must be identical
            digits = re.sub(r"\D", "", ht)
            if digits:
                for k, lst in pdf_map_multi.items():
                    kd = re.sub(r"\D", "", str(k))
                    if kd and kd == digits:
                        for p in lst:
                            recipients[recipient_key][loc].append(
                                (f"{p.get('hallticket') or 'noid'}_{p.get('pdf_name')}", p["pdf_bytes"])
                            )
                        found_any = True
                        break
        if not found_any:
            missing_log.append({"index": idx, "hallticket": ht, "emails": recipient_key, "location": loc})

    st.markdown(
        f"<div style='background:#f0f6ff; border:1px solid #b3d0f0; border-radius:6px; "
        f"padding:0.45rem 0.9rem; font-size:0.85rem; color:#1a3a6b; margin-bottom:0.4rem;'>"
        f"Recipients prepared: <b>{len(recipients)}</b> — preview below</div>",
        unsafe_allow_html=True
    )
    rec_preview = []
    for em, locs in list(recipients.items())[:200]:
        files_count = sum(len(lst) for lst in locs.values())
        rec_preview.append({"email": em, "locations": ", ".join(locs.keys()), "files": files_count})
    if rec_preview:
        st.dataframe(pd.DataFrame(rec_preview))
    if missing_log:
        st.warning(f"{len(missing_log)} rows had no matching PDFs (sample):")
        st.dataframe(pd.DataFrame(missing_log).head(50))

    if st.button("Prepare ZIPs (grouped by recipient->location)"):
        st.info("Preparing ZIP parts in memory (may use RAM).")
        max_bytes        = int(attachment_limit_mb * 1024 * 1024)
        prepared         = {}
        total_recipients = len(recipients)
        progress         = ProcessTracker(total_recipients, "Preparing ZIP files")
        for i, (em, locs) in enumerate(recipients.items(), start=1):
            progress.update(f"Preparing files for: {em}")
            prepared[em] = []
            for loc, files in locs.items():
                safe_prefix = re.sub(r"[^A-Za-z0-9]+","_", loc)[:40] or "loc"
                parts       = split_files_into_zip_parts(files, max_bytes, zip_name_prefix=safe_prefix)
                prepared[em].append((loc, parts))
        progress.done("ZIP preparation complete!")
        st.session_state["prepared"] = prepared
        st.session_state["row_meta"] = row_meta
        st.success("Prepared ZIP parts stored in session memory.")


# ================================================================
# Preview & Send
# ================================================================
if "prepared" in st.session_state:
    st.markdown(
        "<div style='padding:0.5rem 1rem; background:#0b74de; color:#fff; "
        "border-radius:6px; font-size:1rem; font-weight:600;'>Step 3 — ZIP Preview</div>",
        unsafe_allow_html=True
    )
    st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)
    preview_rows     = []
    location_summary = defaultdict(list)
    for em, locs in st.session_state["prepared"].items():
        for loc, parts in locs:
            for pname, pbytes in parts:
                sz = len(pbytes) if pbytes else 0
                preview_rows.append({
                    "email": em, "location": loc,
                    "zip_name": pname, "size": human_bytes(sz)
                })
                location_summary[loc].append(pname)
    if preview_rows:
        st.dataframe(pd.DataFrame(preview_rows).head(500))
    loc_summary_rows = []
    for loc, partnames in location_summary.items():
        loc_summary_rows.append({
            "Location": loc, "PartsCount": len(partnames),
            "Parts": ", ".join(partnames)
        })
    if loc_summary_rows:
        st.subheader("Location-wise parts summary")
        st.dataframe(pd.DataFrame(loc_summary_rows))

    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
    st.markdown(
        "<div style='padding:0.5rem 1rem; background:#0b74de; color:#fff; "
        "border-radius:6px; font-size:1rem; font-weight:600;'>Step 4 — Send Emails</div>",
        unsafe_allow_html=True
    )
    st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)

    smtp_user = None
    smtp_pass = None
    try:
        smtp_user = st.secrets["email_credentials"]["smtp_user"]
        smtp_pass = st.secrets["email_credentials"]["smtp_pass"]
        st.success(f"Email credentials loaded successfully for: **{smtp_user}**")
    except KeyError:
        st.error("Email credentials not found. Please create a `.streamlit/secrets.toml` file.")
        st.code("""
# .streamlit/secrets.toml Example
[email_credentials]
smtp_user = "your-email@gmail.com"
smtp_pass = "your-google-app-password"
""")

    test_mode     = st.checkbox("Test mode (send all to test email)", value=True)
    test_email    = st.text_input("Test email (if test mode ON)")
    subj_template = st.text_input("Subject template",
                                   value="Results for {location} (Part {part}/{total_parts})")
    body_template = st.text_area("Body template",
                                  value="Hello,\n\nPlease find attached results for {location} (Part {part}/{total_parts}).\n\nRegards,\nAiclex")

    if st.button("Start sending prepared ZIPs"):
        if not smtp_user or not smtp_pass:
            st.error("Cannot send emails. Please configure your email credentials in the `.streamlit/secrets.toml` file first.")
        else:
            prepared = st.session_state["prepared"]
            row_meta = st.session_state.get("row_meta", {})

            total_sends = sum(
                len(parts)
                for em, locs in prepared.items()
                for loc, parts in locs
            )

            if total_sends == 0:
                st.warning("No prepared ZIPs to send.")
            else:
                # ---- Live UI ----
                st.markdown(
                    f"<div style='background:#0b74de; color:#fff; border-radius:6px; "
                    f"padding:0.6rem 1.2rem; font-size:1rem; font-weight:600; margin-bottom:0.5rem;'>"
                    f"Total emails to send: {total_sends}</div>",
                    unsafe_allow_html=True
                )
                prog_bar       = st.progress(0)
                status_txt     = st.empty()
                counter_txt    = st.empty()
                live_table_hdr = st.empty()
                live_table     = st.empty()

                success_log = []
                failed_log  = []
                sent_count  = 0
                live_rows   = []

                try:
                    with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as s:
                        s.ehlo(); s.starttls(); s.ehlo()
                        s.login(smtp_user, smtp_pass)

                        for email_key, locs in prepared.items():
                            recipient_list_orig = [e.strip() for e in email_key.split(',') if e.strip()]
                            meta_list           = row_meta.get(email_key, [])
                            disp_name  = ", ".join(sorted({m["name"] for m in meta_list if m["name"]})) or "—"
                            disp_halls = ", ".join(sorted({m["hallticket"] for m in meta_list if m["hallticket"]})) or "—"

                            for loc, parts in locs:
                                total_parts = len(parts)
                                for part_idx, (zipname, zipbytes) in enumerate(parts, start=1):
                                    sent_count += 1

                                    final_recipients = ([test_email] if test_mode and test_email
                                                        else recipient_list_orig)

                                    if not final_recipients:
                                        err_msg = "Recipient email address is empty"
                                        failed_log.append({
                                            "recipients": email_key, "loc": loc,
                                            "zip": zipname, "error": err_msg
                                        })
                                        live_rows.append({
                                            "#": sent_count, "Name": disp_name,
                                            "Hallticket": disp_halls,
                                            "Email": email_key, "Location": loc,
                                            "ZIP": zipname, "Status": "Failed", "Error": err_msg
                                        })
                                        log_send(email_key, disp_name, disp_halls,
                                                 loc, zipname, "Failed", err_msg)
                                        continue

                                    msg = EmailMessage()
                                    msg["From"]    = smtp_user
                                    msg["To"]      = ", ".join(final_recipients)
                                    msg["Subject"] = subj_template.format(
                                        location=loc, part=part_idx, total_parts=total_parts
                                    )
                                    msg.set_content(body_template.format(
                                        location=loc, part=part_idx, total_parts=total_parts
                                    ))
                                    if zipbytes:
                                        msg.add_attachment(zipbytes, maintype="application",
                                                           subtype="zip", filename=zipname)

                                    counter_txt.markdown(
                                        f"**Sending: {sent_count} / {total_sends}** &nbsp;|&nbsp; "
                                        f"Sent: {len(success_log)} &nbsp;|&nbsp; "
                                        f"Failed: {len(failed_log)}"
                                    )
                                    status_txt.text(
                                        f"Sending to {final_recipients[0]}  |  {loc}  Part {part_idx}/{total_parts}"
                                    )

                                    try:
                                        s.send_message(msg)
                                        success_log.append({
                                            "timestamp":  datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                            "recipients": msg["To"],
                                            "subject":    msg["Subject"],
                                            "zip_name":   zipname,
                                            "status":     "Success"
                                        })
                                        row_status = "Sent"
                                        log_send(", ".join(final_recipients), disp_name, disp_halls,
                                                 loc, zipname, "Sent")
                                        time.sleep(send_delay)
                                    except Exception as e:
                                        logger.error("Failed to send to %s: %s", final_recipients, e)
                                        failed_log.append({
                                            "recipients": msg["To"], "loc": loc,
                                            "zip": zipname, "error": str(e)
                                        })
                                        row_status = "Failed"
                                        log_send(", ".join(final_recipients), disp_name, disp_halls,
                                                 loc, zipname, "Failed", str(e))

                                    live_rows.append({
                                        "#": sent_count, "Name": disp_name,
                                        "Hallticket": disp_halls,
                                        "Email": ", ".join(final_recipients),
                                        "Location": loc, "ZIP": zipname,
                                        "Status": row_status, "Error": ""
                                    })

                                    live_table_hdr.markdown("#### Live Send Log")
                                    live_table.dataframe(
                                        pd.DataFrame(live_rows[-200:]),
                                        use_container_width=True
                                    )
                                    prog_bar.progress(min(1.0, sent_count / total_sends))

                except Exception as e:
                    st.error(f"A critical error occurred with the SMTP connection: {e}")

                status_txt.empty()
                counter_txt.empty()
                st.markdown(
                    f"<div style='background:#1a7a4a; color:#fff; border-radius:6px; "
                    f"padding:0.6rem 1.2rem; font-size:0.95rem; font-weight:600; margin:0.5rem 0;'>"
                    f"Sending complete — Successful: {len(success_log)} &nbsp;|&nbsp; "
                    f"Failed: {len(failed_log)} &nbsp; (Total: {total_sends})</div>",
                    unsafe_allow_html=True
                )
                if success_log:
                    st.subheader("Success Log")
                    st.dataframe(pd.DataFrame(success_log))
                if failed_log:
                    st.subheader("Failure Log")
                    st.dataframe(pd.DataFrame(failed_log))

    # ---- Persistent logs ----
    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
    st.markdown(
        "<div style='padding:0.4rem 1rem; background:#f0f4f8; border:1px solid #cbd5e0; "
        "border-radius:6px; font-size:0.9rem; font-weight:600; color:#2d3748;'>"
        "Send History — All-time Log (SQLite)</div>",
        unsafe_allow_html=True
    )
    st.markdown("<div style='height:0.4rem'></div>", unsafe_allow_html=True)
    all_logs = load_send_logs()
    if not all_logs.empty:
        st.dataframe(all_logs, use_container_width=True)
        log_csv = all_logs.to_csv(index=False).encode()
        st.download_button(
            "Download all send logs (CSV)",
            data=log_csv,
            file_name=f"send_logs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    else:
        st.info("No send logs found yet. Logs will appear here after sending emails.")

st.markdown(
    f"<div style='margin-top:2rem; padding:0.6rem 1.2rem; background:#f0f4f8; "
    f"border-top:1px solid #cbd5e0; border-radius:6px; color:#555; font-size:0.78rem;'>"
    f"Built by {BRAND} &nbsp;|&nbsp; CRUX Result Sending System"
    f"</div>",
    unsafe_allow_html=True
)
