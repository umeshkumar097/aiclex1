# streamlit_app.py
"""
Aiclex — Result Showing (Final Streamlit App)
"""

import os, io, re, time, zipfile, logging
from collections import defaultdict
from datetime import datetime
from email.message import EmailMessage

import streamlit as st
import pandas as pd
import pdfplumber
from PIL import Image
import pytesseract

try:
    from pdf2image import convert_from_bytes
    PDF2IMAGE = True
except:
    PDF2IMAGE = False

# ---------------- CONFIG ----------------
APP_TITLE = "Aiclex — Result Showing"
BRAND = "Aiclex Technologies"
MAX_ATTACHMENT_MB = 3
DEFAULT_OCR_DPI = 200
DEFAULT_OCR_LANG = "eng"

logger = logging.getLogger("aiclex")
if not logger.handlers:
    h = logging.StreamHandler()
    h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(h)
logger.setLevel(logging.INFO)

# ---------------- UI ----------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.markdown(f"<h1 style='color:#0b74de'>{APP_TITLE}</h1><div style='color:gray'>Built by {BRAND}</div>", unsafe_allow_html=True)
st.write("---")

LABEL_RE = re.compile(r"Marks\s*Obtained", re.IGNORECASE)
MARKS_NUM_RE = re.compile(r"\b([0-9]{1,3})\b")
ABSENT_RE = re.compile(r"\b(absent|not present)\b", re.IGNORECASE)
PASSFAIL_RE = re.compile(r"([0-9]{1,3})\s*(PASS|FAIL)", re.IGNORECASE)
HALL_RE = re.compile(r"\b[0-9]{3,}\b")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# ---------------- Helpers ----------------
def is_pdf_bytes(b: bytes) -> bool:
    return bool(b) and b.lstrip().startswith(b"%PDF")

def extract_text(pdf_bytes, dpi=200, lang="eng"):
    texts=[]
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for p in pdf.pages:
                t=p.extract_text() or ""
                if t.strip(): texts.append(t)
    except: pass
    if texts: return "\n".join(texts)

    if PDF2IMAGE:
        try:
            pages=convert_from_bytes(pdf_bytes,dpi=dpi)
            return "\n".join(pytesseract.image_to_string(im,lang=lang) for im in pages)
        except: pass

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            ocr=[]
            for p in pdf.pages:
                try:
                    im=p.to_image(resolution=dpi).original
                    ocr.append(pytesseract.image_to_string(im,lang=lang))
                except: continue
            return "\n".join(ocr)
    except: pass
    return ""

def parse_pdf(pdf_bytes,fname="",dpi=DEFAULT_OCR_DPI,lang=DEFAULT_OCR_LANG):
    text=extract_text(pdf_bytes,dpi=dpi,lang=lang)
    hall=""
    h=HALL_RE.findall(text)
    if h: hall=max(h,key=len)
    else:
        digits=re.findall(r"\d+",os.path.basename(fname))
        hall=digits[-1] if digits else ""

    marks=None; status="Absent"
    if ABSENT_RE.search(text):
        marks=""; status="Absent"
    else:
        pf=PASSFAIL_RE.search(text)
        if pf:
            val=int(pf.group(1))
            marks=val; status="Pass" if val>49 else "Fail"
        else:
            lbl=LABEL_RE.search(text)
            if lbl:
                snip=text[lbl.end():lbl.end()+100]
                mnum=re.search(r"([0-9]{1,3})",snip)
                if mnum:
                    val=int(mnum.group(1))
                    marks=val; status="Pass" if val>49 else "Fail"
                else:
                    marks=""; status="Absent"
            else:
                nums=[int(n) for n in MARKS_NUM_RE.findall(text) if 0<=int(n)<=100]
                if nums:
                    val=nums[-1]; marks=val; status="Pass" if val>49 else "Fail"
                else:
                    marks=""; status="Absent"
    return {"pdf_name":os.path.basename(fname),"pdf_bytes":pdf_bytes,"hallticket":hall,"marks":marks,"status":status}

def extract_zip_safe(zip_bytes,dpi,lang):
    results=[]
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            for name in zf.namelist():
                if name.endswith("/"): continue
                try: data=zf.read(name)
                except: continue
                lname=name.lower()
                if lname.endswith(".zip"):
                    try: results.extend(extract_zip_safe(data,dpi,lang))
                    except zipfile.BadZipFile:
                        if is_pdf_bytes(data): results.append(parse_pdf(data,name,dpi,lang))
                elif lname.endswith(".pdf") or is_pdf_bytes(data):
                    results.append(parse_pdf(data,name,dpi,lang))
    except: pass
    return results

def fill_excel(df,pdfs,hall_col):
    pdf_map={p["hallticket"]:p for p in pdfs if p["hallticket"]}
    if "marks" not in df: df["marks"]=""
    if "status" not in df: df["status"]=""
    for i,row in df.iterrows():
        ht=str(row.get(hall_col,"")).strip()
        if ht in pdf_map:
            m=pdf_map[ht]["marks"]; df.at[i,"marks"]=m
            if isinstance(m,int): df.at[i,"status"]="Pass" if m>49 else "Fail"
            elif m=="": df.at[i,"status"]="Absent"
    return df

def make_zip(files):
    bio=io.BytesIO()
    with zipfile.ZipFile(bio,"w",compression=zipfile.ZIP_DEFLATED) as z:
        for fn,b in files: z.writestr(fn,b)
    bio.seek(0); return bio.read()

def split_zip(files,max_bytes,prefix):
    parts=[];cur=[]
    for fn,b in files:
        test=cur+[(fn,b)]
        if len(make_zip(test))<=max_bytes: cur=test
        else:
            if cur: parts.append((f"{prefix}_part{len(parts)+1}.zip",make_zip(cur)))
            cur=[(fn,b)]
    if cur: parts.append((f"{prefix}_part{len(parts)+1}.zip",make_zip(cur)))
    return parts

def send_email(user,pwd,to_emails,subj,body,atts):
    msg=EmailMessage()
    msg["From"]=user
    msg["To"]=", ".join(to_emails) if isinstance(to_emails,list) else to_emails
    msg["Subject"]=subj
    msg.set_content(body)
    for fn,b in atts:
        msg.add_attachment(b,maintype="application",subtype="zip",filename=fn)
    import smtplib
    with smtplib.SMTP("smtp.gmail.com",587,timeout=60) as s:
        s.starttls(); s.login(user,pwd); s.send_message(msg)

# ---------------- UI Flow ----------------
st.header("Step 1 — Upload Excel & ZIP")
excel=st.file_uploader("Excel/CSV",type=["xlsx","csv"])
zipf=st.file_uploader("ZIP (nested allowed)",type=["zip"])

if excel and zipf:
    df=pd.read_csv(excel,dtype=str).fillna("") if excel.name.endswith("csv") else pd.read_excel(excel,dtype=str).fillna("")
    st.success(f"Excel {len(df)} rows loaded")
    hall_col=st.selectbox("Hallticket col",df.columns)
    email_col=st.selectbox("Email col",df.columns)
    loc_col=st.selectbox("Location col",df.columns)

    pdfs=extract_zip_safe(zipf.read(),DEFAULT_OCR_DPI,DEFAULT_OCR_LANG)
    st.info(f"Extracted {len(pdfs)} PDFs")

    df=fill_excel(df,pdfs,hall_col)
    st.dataframe(df.head(50))

    buf=io.BytesIO(); df.to_excel(buf,index=False,engine="openpyxl"); buf.seek(0)
    st.download_button("Download filled Excel",buf,"results.xlsx")

    # prepare recipients
    recips=defaultdict(lambda:defaultdict(list))
    for r in df.itertuples():
        ems=str(getattr(r,email_col)).strip()
        loc=str(getattr(r,loc_col)).strip()
        ht=str(getattr(r,hall_col)).strip()
        if not ems: continue
        emails=[e.strip() for e in re.split(r"[;, \n]+",ems) if e.strip()]
        for e in emails:
            for p in pdfs:
                if p["hallticket"]==ht:
                    recips[e][loc].append((p["pdf_name"],p["pdf_bytes"]))

    max_bytes=int(MAX_ATTACHMENT_MB*1024*1024)
    prepared={em:{loc:split_zip(files,max_bytes,loc) for loc,files in locs.items()} for em,locs in recips.items()}
    st.session_state["prepared"]=prepared

if "prepared" in st.session_state:
    st.header("Step 2 — Send Emails")
    smtp_user=st.text_input("Gmail",value="info@cruxmanagement.com")
    smtp_pass=st.text_input("App Password",type="password",value="norx wxop hvsm bvfu")
    test_mode=st.checkbox("Test mode",value=True)
    test_email=st.text_input("Test email if test mode")
    subj=st.text_input("Subject template","Results for {location} (Part {part}/{total_parts})")
    body=st.text_area("Body template","Dear,\n\nPlease find attached result for {location} (Part {part}/{total_parts}).\n\nRegards,\nAiclex")

    if st.button("Send"):
        total_parts=sum(len(parts) for locs in st.session_state["prepared"].values() for parts in locs.values())
        cur=0; prog=st.progress(0)
        for em,locs in st.session_state["prepared"].items():
            rec_list=[e.strip() for e in re.split(r"[;, \n]+",em) if e.strip()]
            if test_mode: rec_list=[test_email] if test_email else []
            for loc,parts in locs.items():
                total=len(parts)
                for idx,(pn,pb) in enumerate(parts,1):
                    s=subj.format(location=loc,part=idx,total_parts=total)
                    b=body.format(location=loc,part=idx,total_parts=total)
                    send_email(smtp_user,smtp_pass,rec_list,s,b,[(pn,pb)])
                    cur+=1; prog.progress(cur/total_parts)
        st.success("Emails sent successfully ✅")
