# streamlit_app.py
"""
Aiclex Hallticket Result Mailer
--------------------------------
- Upload Excel + ZIP of PDFs (nested allowed).
- Extract hallticket & marks/Absent from PDFs.
- Fill Excel with marks + status.
- Location-wise grouping, split ZIPs if >3MB.
- Send emails with Gmail SMTP + progress bar.
"""

import os, io, re, zipfile, tempfile, time, smtplib
from email.message import EmailMessage
from collections import defaultdict
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import pytesseract

# ---------------- Config ----------------
MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024  # 3MB limit for attachments

st.set_page_config(page_title="Aiclex Hallticket Mailer", layout="wide")
st.title("📩 Aiclex Technologies — Hallticket Result Mailer")

# ---------------- Regex ----------------
HALL_RE = re.compile(r"\b[0-9]{4,}\b")
ABSENT_RE = re.compile(r"\b(absent|not present|a\s*b\s*s\s*e\s*n\s*t)\b", re.IGNORECASE)
MARKS_RE = re.compile(r"(?:marks|mark|score|total)[:\s\-]*([0-9]{1,3})", re.IGNORECASE)

# ---------------- PDF Helpers ----------------
def extract_text_from_pdf(pdf_bytes):
    """Extract text from PDF; fallback to OCR if needed."""
    text_parts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for p in pdf.pages:
                t = p.extract_text() or ""
                if t.strip():
                    text_parts.append(t)
    except Exception:
        pass
    if text_parts:
        return "\n".join(text_parts)
    # OCR fallback
    ocr_texts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for p in pdf.pages:
                try:
                    img = p.to_image(resolution=150).original
                    ocr_texts.append(pytesseract.image_to_string(img))
                except Exception:
                    continue
    except Exception:
        pass
    return "\n".join(ocr_texts)

def find_hallticket(text):
    c = HALL_RE.findall(text or "")
    return max(c, key=len) if c else None

def find_marks_or_absent(text):
    if not text: return None
    lower = text.lower()
    if ABSENT_RE.search(lower): return "Absent"
    m = MARKS_RE.search(text)
    if m:
        try:
            val = int(m.group(1))
            if 0 <= val <= 100: return val
        except: pass
    # fallback: any number 0–100
    nums = re.findall(r"\b(\d{1,3})\b", text)
    nums = [int(n) for n in nums if 0 <= int(n) <= 100]
    return nums[-1] if nums else None

# ---------------- ZIP Processing ----------------
def process_zip(zip_bytes):
    """Recursively extract PDFs from ZIP and return list of dicts {hallticket, marks, pdf_bytes, pdf_name}"""
    results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith("/"): continue
                data = zf.read(name)
                if name.lower().endswith(".zip"):
                    results.extend(process_zip(data))
                elif name.lower().endswith(".pdf"):
                    text = extract_text_from_pdf(data)
                    hall = find_hallticket(text)
                    marks = find_marks_or_absent(text)
                    results.append({
                        "hallticket": hall,
                        "marks": marks,
                        "pdf_bytes": data,
                        "pdf_name": os.path.basename(name)
                    })
    except Exception:
        pass
    return results

# ---------------- Excel Fill ----------------
def fill_excel(df, pdf_data, hall_col):
    pdf_map = {str(p["hallticket"]).strip(): p["marks"] for p in pdf_data if p["hallticket"]}
    if "marks" not in df.columns: df["marks"] = ""
    if "status" not in df.columns: df["status"] = ""
    filled = 0
    for idx,row in df.iterrows():
        ht = str(row.get(hall_col,"")).strip()
        if not ht: continue
        mval = None
        if ht in pdf_map:
            mval = pdf_map[ht]
        else:
            digits = re.sub(r"\D","",ht)
            if digits and digits in pdf_map: mval = pdf_map[digits]
        # Decide
        if mval is None or (isinstance(mval,str) and mval.lower().startswith("abs")):
            df.at[idx,"marks"] = ""
            df.at[idx,"status"] = "Absent"
        else:
            try:
                mm = int(mval)
                df.at[idx,"marks"] = mm
                df.at[idx,"status"] = "Pass" if mm>49 else "Fail"
            except:
                df.at[idx,"marks"] = ""
                df.at[idx,"status"] = "Absent"
        filled += 1
    return df, filled

# ---------------- ZIP Split ----------------
def make_zip(parts):
    bio = io.BytesIO()
    with zipfile.ZipFile(bio,"w",compression=zipfile.ZIP_DEFLATED) as z:
        for fname,b in parts:
            z.writestr(fname,b)
    bio.seek(0); return bio.read()

def split_zip(files, prefix):
    parts=[]; current=[]; 
    for fname,b in files:
        test=current+[(fname,b)]
        test_zip=make_zip(test)
        if len(test_zip)<=MAX_ATTACHMENT_BYTES:
            current=test
        else:
            if current:
                parts.append((f"{prefix}_part{len(parts)+1}.zip",make_zip(current)))
            current=[(fname,b)]
    if current: parts.append((f"{prefix}_part{len(parts)+1}.zip",make_zip(current)))
    return parts

# ---------------- Email ----------------
def send_email(user,pwd,to,subject,body,attachments):
    msg=EmailMessage()
    msg["From"]=user; msg["To"]=to; msg["Subject"]=subject
    msg.set_content(body)
    for fname,b in attachments:
        msg.add_attachment(b,maintype="application",subtype="zip",filename=fname)
    with smtplib.SMTP("smtp.gmail.com",587,timeout=60) as s:
        s.starttls(); s.login(user,pwd); s.send_message(msg)

# ---------------- UI ----------------
st.header("1. Upload Excel & ZIP")
excel_file = st.file_uploader("Excel/CSV",type=["xlsx","csv"])
zip_file = st.file_uploader("ZIP (with PDFs)",type=["zip"])

if excel_file and zip_file:
    df = pd.read_csv(excel_file,dtype=str).fillna("") if excel_file.name.endswith(".csv") else pd.read_excel(excel_file,dtype=str).fillna("")
    st.success(f"Excel loaded with {len(df)} rows")
    hall_col = st.selectbox("Hallticket column",df.columns)
    email_col = st.selectbox("Email column",df.columns)
    loc_col = st.selectbox("Location column",df.columns)

    pdf_data = process_zip(zip_file.read())
    st.info(f"Processed {len(pdf_data)} PDFs")

    df, filled = fill_excel(df,pdf_data,hall_col)
    st.success(f"Updated {filled} rows with marks/status")
    st.dataframe(df.head(50))

    # Download updated Excel
    buf=io.BytesIO(); df.to_excel(buf,index=False,engine="openpyxl"); buf.seek(0)
    st.download_button("Download updated Excel",buf,file_name=f"updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    # Email system
    st.header("2. Send Emails")
    smtp_user=st.text_input("Gmail address",value=os.environ.get("SMTP_USER",""))
    smtp_pass=st.text_input("App Password",type="password",value=os.environ.get("SMTP_PASS",""))
    if st.button("Start Sending"):
        if not smtp_user or not smtp_pass:
            st.error("Provide Gmail + App Password")
        else:
            prog=st.progress(0); status=st.empty()
            recipients = []
            for idx,row in df.iterrows():
                recipients.append((row[email_col],row[loc_col],row[hall_col]))
            total=len(recipients)
            done=0
            for email,loc,ht in recipients:
                files=[(f"{ht}.pdf",p["pdf_bytes"]) for p in pdf_data if str(p["hallticket"])==str(ht)]
                if not files: continue
                zips=split_zip(files,loc)
                send_email(smtp_user,smtp_pass,email,f"Result for {loc}",f"Dear Participant,\n\nPlease find attached your result.\n\nRegards,\nAiclex",zips)
                done+=1
                prog.progress(done/total); status.text(f"Sent {done}/{total} → {email}")
                time.sleep(1)
            st.success("All emails sent ✅")
