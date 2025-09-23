# streamlit_app.py
"""
Aiclex — PDF ≠ Excel Matcher + Email System
Features:
- Upload Excel/CSV with Hallticket, Email, Location
- Upload nested ZIP containing PDFs
- Extract hallticket + marks/Absent from PDFs (OCR fallback)
- Fill Excel with marks & status
- Group PDFs by Location -> ZIP -> Split if >3MB
- Send results by Gmail SMTP with Streamlit progress bar
"""

import os
import io
import re
import zipfile
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import smtplib
from email.message import EmailMessage

import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
from PIL import Image

# ------------------ Config ------------------
APP_TITLE = "Aiclex — PDF ≠ Excel Matcher"
BRAND_COLOR = "#0b74de"
LOGO_PATH = None
MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024  # 3MB

st.set_page_config(page_title="Aiclex PDF-Excel Matcher", layout="wide")

# Branding header
header_html = (
    "<div style='display:flex; align-items:center;'>"
    + ("<div style='margin-right:16px;'><img src='{}' width='120' /></div>".format(LOGO_PATH) if LOGO_PATH and os.path.exists(LOGO_PATH) else "")
    + "<div><h1 style='color:{}; margin:0;'>{}</h1>"
    + "<div style='color:gray'>Built by Aiclex Technologies</div></div></div>"
).format(BRAND_COLOR, APP_TITLE)
st.markdown(header_html, unsafe_allow_html=True)
st.markdown("---")

# ------------------ Patterns ------------------
HALLTICKET_PATTERN = re.compile(r"\b[0-9]{4,}\b")
ABSENT_PATTERN = re.compile(r"\b(?:absent|not present|a\s*b\s*s\s*e\s*n\s*t|a\.?b\.?s\b|abs\b)\b", re.IGNORECASE)
MARKS_PATTERN = re.compile(r"(?:marks|mark|score|total)[:\s\-]*([0-9]{1,3})", re.IGNORECASE)

# ------------------ Helpers ------------------
def extract_text_from_pdf_bytes(pdf_bytes):
    text_parts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for p in pdf.pages:
                t = p.extract_text() or ""
                if t.strip():
                    text_parts.append(t)
    except Exception:
        pass
    combined = "\n".join(text_parts).strip()
    if combined:
        return combined
    # OCR fallback
    ocr_texts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for p in pdf.pages:
                try:
                    img = p.to_image(resolution=150).original
                    ocr = pytesseract.image_to_string(img)
                    ocr_texts.append(ocr)
                except Exception:
                    continue
    except Exception:
        pass
    return "\n".join(ocr_texts)

def find_hallticket_in_text(text):
    c = HALLTICKET_PATTERN.findall(text or "")
    return max(c, key=len) if c else None

def find_marks_or_absent(text):
    if not text:
        return None
    txt = text.replace('\xa0', ' ')
    lower = txt.lower()
    if ABSENT_PATTERN.search(lower):
        return 'Absent'
    for line in lower.splitlines():
        if 'absent' in line or 'not present' in line:
            return 'Absent'
    m = MARKS_PATTERN.search(txt)
    if m:
        try:
            val = int(m.group(1))
            if 0 <= val <= 100:
                return val
        except:
            pass
    # fallback numbers
    candidates = []
    for line in txt.splitlines():
        nums = re.findall(r"\b(\d{1,3})\b", line)
        for n in nums:
            v = int(n)
            if 0 <= v <= 100:
                weight = 2 if re.search(r'(?:marks|mark|score|total)', line, re.IGNORECASE) else 1
                candidates.append((weight, v))
    if candidates:
        maxw = max(c[0] for c in candidates)
        return [c[1] for c in candidates if c[0] == maxw][-1]
    return None

def process_uploaded_zip_bytes(zip_bytes):
    results = []
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as outer:
            for name in outer.namelist():
                if name.endswith('/'): continue
                data = outer.read(name)
                if name.lower().endswith('.zip'):
                    results.extend(process_uploaded_zip_bytes(data))
                elif name.lower().endswith('.pdf'):
                    text = extract_text_from_pdf_bytes(data)
                    hall = find_hallticket_in_text(text) or ""
                    marks = find_marks_or_absent(text)
                    results.append({
                        'hallticket': str(hall).strip(),
                        'marks': marks,
                        'source_pdf': name,
                        'pdf_bytes': data
                    })
    except zipfile.BadZipFile:
        pass
    return results

def fill_excel_with_pdf_data(df, pdf_data):
    df = df.copy()
    df_cols = {c.lower(): c for c in df.columns}
    hall_col = next((v for k,v in df_cols.items() if 'hall' in k and 'ticket' in k), None)
    marks_col = df_cols.get('marks','marks'); 
    status_col = df_cols.get('status','status')
    email_col = next((v for k,v in df_cols.items() if 'email' in k), None)
    location_col = next((v for k,v in df_cols.items() if 'location' in k), None)
    if marks_col not in df.columns: df[marks_col] = None
    if status_col not in df.columns: df[status_col] = None
    lookup = {str(p['hallticket']).strip(): p for p in pdf_data if p['hallticket']}
    filled = 0; unmatched=[]
    for idx,row in df.iterrows():
        hall = str(row.get(hall_col,"")).strip()
        if not hall: continue
        m=None; rec=None
        if hall in lookup: rec=lookup[hall]; m=rec['marks']
        if m is not None:
            df.at[idx,marks_col]=m
            if isinstance(m,str) and m.lower().startswith('abs'):
                df.at[idx,status_col]='Absent'
            else:
                try: df.at[idx,status_col]='Fail' if int(m)<49 else 'Pass'
                except: df.at[idx,status_col]='Unknown'
            filled+=1
        else:
            unmatched.append({'row':idx,'hallticket':hall})
    return df, filled, unmatched, email_col, location_col

# ------------------ Email Helpers ------------------
def make_zip_bytes(file_entries):
    bio=io.BytesIO()
    with zipfile.ZipFile(bio,"w",compression=zipfile.ZIP_DEFLATED) as zf:
        for fname,b in file_entries:
            zf.writestr(fname,b)
    bio.seek(0); return bio.read()

def split_files_into_zips(file_entries,max_bytes=MAX_ATTACHMENT_BYTES,zip_prefix="results"):
    parts=[]; current=[]
    for fname,b in file_entries:
        test=current+[(fname,b)]
        test_zip=make_zip_bytes(test)
        if len(test_zip)<=max_bytes: current=test
        else:
            if current: parts.append((f"{zip_prefix}_part{len(parts)+1}.zip",make_zip_bytes(current)))
            current=[(fname,b)]
    if current: parts.append((f"{zip_prefix}_part{len(parts)+1}.zip",make_zip_bytes(current)))
    return parts

def send_email_with_attachments_gmail(user,pwd,to,subject,body,attachments):
    msg=EmailMessage()
    msg['From']=user; msg['To']=to; msg['Subject']=subject
    msg.set_content(body)
    for fname,b in attachments:
        msg.add_attachment(b,maintype="application",subtype="zip",filename=fname)
    try:
        with smtplib.SMTP("smtp.gmail.com",587,timeout=60) as s:
            s.starttls(); s.login(user,pwd); s.send_message(msg)
        return True,None
    except Exception as e: return False,str(e)

def send_grouped_results(df,pdf_data,email_col,loc_col,smtp_user,smtp_pass,progress_cb=None):
    # map hallticket->pdf
    pdf_map=defaultdict(list)
    for p in pdf_data: pdf_map[p['hallticket']].append(p)
    # group by location
    rec_by_loc=defaultdict(list)
    for _,row in df.iterrows():
        hall=str(row.get('Hallticket') or row.get('hallticket') or "").strip()
        email=str(row.get(email_col) or "").strip()
        loc=str(row.get(loc_col) or "Unknown").strip()
        if email: rec_by_loc[loc].append({'email':email,'hall':hall})
    total=sum(len(v) for v in rec_by_loc.values()); sent=0; fails=[]
    for loc,recips in rec_by_loc.items():
        file_entries=[]
        for r in recips:
            for p in pdf_map.get(r['hall'],[]): 
                fname=f"{r['hall']}_{Path(p['source_pdf']).name}"
                file_entries.append((fname,p['pdf_bytes']))
        zips=split_files_into_zips(file_entries,MAX_ATTACHMENT_BYTES,f"{loc}_results")
        for r in recips:
            ok,err=send_email_with_attachments_gmail(smtp_user,smtp_pass,r['email'],
                f"Results for {loc}",
                f"Dear Participant,\n\nPlease find attached results for {loc}.\n\nRegards,\nAiclex",
                zips)
            sent+=ok; 
            if not ok: fails.append({'email':r['email'],'error':err})
            if progress_cb: progress_cb((sent+len(fails))/total,f"Sent to {r['email']}")
    return {"sent":sent,"failed":fails,"total":total}

# ------------------ Streamlit UI ------------------
st.write("Upload Excel/CSV and a ZIP (nested ZIPs allowed) containing PDFs.")
col1,col2=st.columns(2)
with col1: uploaded_excel=st.file_uploader("Excel/CSV file",type=["xlsx","xls","csv"])
with col2: uploaded_zip=st.file_uploader("ZIP file",type=["zip"])
if st.button("Process and Match"):
    if not uploaded_excel or not uploaded_zip: st.error("Please upload both files."); st.stop()
    # read excel
    df=pd.read_csv(uploaded_excel,dtype=str) if uploaded_excel.name.endswith(".csv") else pd.read_excel(uploaded_excel,dtype=str,engine="openpyxl")
    st.success(f"Excel loaded {len(df)} rows")
    # process zip
    pdf_data=process_uploaded_zip_bytes(uploaded_zip.read())
    st.info(f"Extracted {len(pdf_data)} pdf records")
    updated_df,filled,unmatched,email_col,loc_col=fill_excel_with_pdf_data(df,pdf_data)
    st.success(f"Filled {filled} rows")
    st.dataframe(updated_df.head(50))
    # download
    buf=io.BytesIO(); updated_df.to_excel(buf,index=False,engine="openpyxl"); buf.seek(0)
    st.download_button("Download updated Excel",buf,file_name=f"updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    # email
    if st.checkbox("Send grouped results by Gmail"):
        smtp_user=st.text_input("Gmail address",value=os.environ.get("SMTP_USER",""))
        smtp_pass=st.text_input("Gmail App Password",type="password",value=os.environ.get("SMTP_PASS",""))
        if st.button("Start Sending"):
            prog=st.progress(0); status=st.empty()
            def cb(frac,text): prog.progress(min(1,frac)); status.text(text)
            result=send_grouped_results(updated_df,pdf_data,email_col,loc_col,smtp_user,smtp_pass,cb)
            st.success(f"Emails sent {result['sent']} / {result['total']}")
            if result['failed']: st.error("Some failed"); st.dataframe(pd.DataFrame(result['failed']))
