import os, io, re, zipfile
from collections import defaultdict
from datetime import datetime
from email.message import EmailMessage

import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
from PIL import Image

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE = True
except ImportError:
    PDF2IMAGE = False

# ---------------- CONFIG ----------------
MAX_ATTACHMENT_MB = 3
OCR_DPI = 200
OCR_LANG = "eng"

st.set_page_config(page_title="Aiclex — Result Showing", layout="wide")
st.title("📊 Aiclex — Result Showing")

# Regex
LABEL_RE = re.compile(r"Marks\s*Obtained", re.IGNORECASE)
MARKS_NUM_RE = re.compile(r"\b([0-9]{1,3})\b")
ABSENT_RE = re.compile(r"\b(absent)\b", re.IGNORECASE)
PASSFAIL_RE = re.compile(r"([0-9]{1,3})\s*(PASS|FAIL)", re.IGNORECASE)
HALL_RE = re.compile(r"\b[0-9]{3,}\b")

# ---------------- Helpers ----------------
def is_pdf_bytes(b: bytes) -> bool:
    """Check PDF header"""
    return bool(b) and b.lstrip().startswith(b"%PDF")

def extract_text(pdf_bytes):
    """Extract text with pdfplumber -> pdf2image+pytesseract fallback"""
    texts = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for p in pdf.pages:
                t = p.extract_text() or ""
                if t.strip():
                    texts.append(t)
    except: pass
    if texts: return "\n".join(texts)

    if PDF2IMAGE:
        try:
            imgs = convert_from_bytes(pdf_bytes, dpi=OCR_DPI)
            return "\n".join(pytesseract.image_to_string(im, lang=OCR_LANG) for im in imgs)
        except: pass

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            ocr_texts=[]
            for p in pdf.pages:
                try:
                    im=p.to_image(resolution=OCR_DPI).original
                    ocr_texts.append(pytesseract.image_to_string(im, lang=OCR_LANG))
                except: continue
            return "\n".join(ocr_texts)
    except: pass
    return ""

def parse_pdf(pdf_bytes, fname=""):
    text = extract_text(pdf_bytes) or ""
    # Hallticket
    hall=""
    h_cands=HALL_RE.findall(text)
    if h_cands: hall=max(h_cands,key=len)
    else:
        digits=re.findall(r"\d+",os.path.basename(fname))
        hall=digits[-1] if digits else ""

    marks=None
    status="Absent"
    # Absent check
    if ABSENT_RE.search(text):
        marks=""
        status="Absent"
    else:
        pf=PASSFAIL_RE.search(text)
        if pf:
            val=int(pf.group(1))
            marks=val
            status="Pass" if val>49 else "Fail"
        else:
            lbl=LABEL_RE.search(text)
            if lbl:
                snippet=text[lbl.end():lbl.end()+100]
                mnum=re.search(r"([0-9]{1,3})",snippet)
                if mnum:
                    val=int(mnum.group(1))
                    marks=val
                    status="Pass" if val>49 else "Fail"
                else:
                    marks=""
                    status="Absent"
            else:
                nums=[int(n) for n in MARKS_NUM_RE.findall(text) if 0<=int(n)<=100]
                if nums:
                    val=nums[-1]
                    marks=val
                    status="Pass" if val>49 else "Fail"
                else:
                    marks=""
                    status="Absent"

    return {
        "pdf_name": os.path.basename(fname),
        "pdf_bytes": pdf_bytes,
        "hallticket": hall,
        "marks": marks,
        "status": status,
        "text_snippet": text[:500]
    }

def extract_zip_safe(zip_bytes):
    results=[]
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith("/"): continue
                try:
                    data=zf.read(name)
                except: continue
                lname=name.lower()
                if lname.endswith(".zip"):
                    try:
                        results.extend(extract_zip_safe(data))
                    except zipfile.BadZipFile:
                        if is_pdf_bytes(data):
                            results.append(parse_pdf(data,name))
                        else:
                            st.warning(f"Skipping non-zip/non-pdf: {name}")
                elif lname.endswith(".pdf") or is_pdf_bytes(data):
                    try:
                        results.append(parse_pdf(data,name))
                    except Exception as e:
                        st.warning(f"Failed PDF {name}: {e}")
                else:
                    continue
    except zipfile.BadZipFile:
        st.error("Top-level is not a valid ZIP")
    return results

def fill_excel(df,pdfs,hall_col):
    pdf_map={p["hallticket"]:p for p in pdfs if p["hallticket"]}
    if "marks" not in df: df["marks"]=""
    if "status" not in df: df["status"]=""
    for i,row in df.iterrows():
        ht=str(row.get(hall_col,"")).strip()
        if ht in pdf_map:
            df.at[i,"marks"]=pdf_map[ht]["marks"]
            df.at[i,"status"]=pdf_map[ht]["status"]
    return df

def make_zip(files):
    bio=io.BytesIO()
    with zipfile.ZipFile(bio,"w",compression=zipfile.ZIP_DEFLATED) as z:
        for fn,b in files: z.writestr(fn,b)
    bio.seek(0)
    return bio.read()

def split_zip(files,max_bytes,prefix):
    parts=[];cur=[]
    for fn,b in files:
        test=cur+[(fn,b)]
        if len(make_zip(test))<=max_bytes:
            cur=test
        else:
            if cur: parts.append((f"{prefix}_part{len(parts)+1}.zip",make_zip(cur)))
            cur=[(fn,b)]
    if cur: parts.append((f"{prefix}_part{len(parts)+1}.zip",make_zip(cur)))
    return parts

def send_email(user,pwd,to,subj,body,atts):
    msg=EmailMessage()
    msg["From"]=user; msg["To"]=to; msg["Subject"]=subj
    msg.set_content(body)
    for fn,b in atts:
        msg.add_attachment(b,maintype="application",subtype="zip",filename=fn)
    import smtplib
    with smtplib.SMTP("smtp.gmail.com",587,timeout=60) as s:
        s.starttls(); s.login(user,pwd); s.send_message(msg)

# ---------------- UI ----------------
st.header("1. Upload Excel & ZIP")
excel=st.file_uploader("Excel/CSV",type=["xlsx","csv"])
zipf=st.file_uploader("ZIP (nested ok)",type=["zip"])

if excel and zipf:
    df=pd.read_csv(excel,dtype=str).fillna("") if excel.name.endswith("csv") else pd.read_excel(excel,dtype=str).fillna("")
    st.success(f"Excel {len(df)} rows loaded")
    hall_col=st.selectbox("Hallticket col",df.columns)
    email_col=st.selectbox("Email col",df.columns)
    loc_col=st.selectbox("Location col",df.columns)

    pdfs=extract_zip_safe(zipf.read())
    st.info(f"Extracted {len(pdfs)} PDFs")

    df=fill_excel(df,pdfs,hall_col)

    # summary
    total=len(df)
    passc=(df["status"]=="Pass").sum()
    failc=(df["status"]=="Fail").sum()
    absc=(df["status"]=="Absent").sum()
    st.write({"Total":total,"Pass":int(passc),"Fail":int(failc),"Absent":int(absc)})

    buf=io.BytesIO(); df.to_excel(buf,index=False,engine="openpyxl"); buf.seek(0)
    st.download_button("Download filled Excel",buf,file_name="results_filled.xlsx")

    # Prepare zips per email+location
    recips=defaultdict(lambda:defaultdict(list))
    for r in df.itertuples():
        em=str(getattr(r,email_col)).strip()
        loc=str(getattr(r,loc_col)).strip()
        ht=str(getattr(r,hall_col)).strip()
        for p in pdfs:
            if p["hallticket"]==ht:
                recips[em][loc].append((p["pdf_name"],p["pdf_bytes"]))
    max_bytes=int(MAX_ATTACHMENT_MB*1024*1024)
    prepared={em:{loc:split_zip(files,max_bytes,loc) for loc,files in locs.items()} for em,locs in recips.items()}

    st.subheader("Prepared ZIP preview")
    rows=[]
    for em,locs in prepared.items():
        for loc,parts in locs.items():
            for pn,pb in parts:
                rows.append({"email":em,"location":loc,"zip":pn,"size":len(pb)})
    st.dataframe(pd.DataFrame(rows))

    st.header("2. Send Emails")
    user=st.text_input("Gmail",value="")
    pwd=st.text_input("App Password",type="password")
    test_mode=st.checkbox("Test mode",value=True)
    test_email=st.text_input("Test email if test mode")
    subj=st.text_input("Subject template","Results for {location}")
    body=st.text_area("Body template","Dear,\n\nPlease find attached result for {location}.\n\nRegards,\nAiclex")

    if st.button("Send"):
        if not user or not pwd: st.error("Need SMTP creds")
        else:
            cnt=0; total_parts=sum(len(parts) for locs in prepared.values() for parts in locs.values())
            prog=st.progress(0)
            for em,locs in prepared.items():
                recs=[test_email] if test_mode else [em]
                for loc,parts in locs.items():
                    for pn,pb in parts:
                        for r in recs:
                            send_email(user,pwd,r,subj.format(location=loc),body.format(location=loc),[(pn,pb)])
                            cnt+=1; prog.progress(cnt/total_parts)
            st.success("Emails sent successfully ✅")

