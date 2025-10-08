# --- Aiclex Media Processor (Streamlit UI) ---
# Author: Aiclex Technologies (aiclex.in)
# Version: 1.2.0
# ---------------------------------------------

import streamlit as st
import pandas as pd
import os
import zipfile
import shutil
import tempfile
from pathlib import Path
from PIL import Image, ImageOps

# Third-party libraries (install from requirements.txt)
try:
    from pdf2image import convert_from_path
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# --- CONFIGURATION ---

TARGET_SIZES_KB = {
    "id_proof": (10, 24), "photo": (10, 24), "signature": (10, 24),
    "qualification_proof": (50, 200), "unknown": (10, 500)
}
CATEGORY_KEYWORDS = {
    "id_proof": ["id", "aadhaar", "pan", "passport", "dl", "driving"],
    "photo": ["photo", "pic", "profile", "image"], "signature": ["sign", "signature"],
    "qualification_proof": ["degree", "certificate", "marksheet", "qualification", "qual"]
}
SUPPORTED_MEDIA_EXT = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.pdf']
MIN_QUALITY = 10
DPI = 200

# --- CORE PROCESSING FUNCTIONS ---

def extract_all_zips(start_dir, status_placeholder):
    """Recursively finds and extracts all nested ZIP files."""
    while True:
        zip_found_in_scan = False
        for root, _, files in os.walk(start_dir):
            for filename in files:
                if filename.lower().endswith('.zip'):
                    zip_path = Path(root) / filename
                    extract_folder = zip_path.with_suffix('')
                    status_placeholder.text(f"  -> Extracting nested ZIP: {filename}...")
                    try:
                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                            zip_ref.extractall(extract_folder)
                        zip_path.unlink()
                        zip_found_in_scan = True
                    except zipfile.BadZipFile:
                        st.warning(f"Bad ZIP file, cannot extract: {filename}")
                    break
            if zip_found_in_scan:
                break
        if not zip_found_in_scan:
            break

def guess_category_by_filename(filename):
    fn_lower = Path(filename).stem.lower()
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(keyword in fn_lower for keyword in keywords):
            return category
    return "unknown"

def convert_any_to_jpg(input_path, output_jpg_path):
    try:
        ext = input_path.suffix.lower()
        output_jpg_path.parent.mkdir(parents=True, exist_ok=True)
        if ext == '.pdf':
            if not PDF_SUPPORT: return False
            pages = convert_from_path(input_path, dpi=DPI, first_page=1, last_page=1)
            if pages: pages[0].save(output_jpg_path, 'JPEG')
            return True
        elif ext in SUPPORTED_MEDIA_EXT:
            with Image.open(input_path) as img:
                img = ImageOps.exif_transpose(img)
                if img.mode != 'RGB': img = img.convert('RGB')
                img.save(output_jpg_path, 'JPEG')
            return True
        return False
    except Exception:
        return False

def compress_jpg_to_target(input_path, output_path, min_kb, max_kb):
    try:
        img = Image.open(input_path)
        img_copy = img.copy()
        for quality in range(95, MIN_QUALITY - 1, -5):
            img_copy.save(output_path, 'JPEG', quality=quality, optimize=True)
            if min_kb <= (output_path.stat().st_size / 1024) <= max_kb:
                return "ok", output_path.stat().st_size / 1024
        best_attempt_path = Path(str(output_path) + ".best")
        shutil.copy(output_path, best_attempt_path)
        for scale in [0.9, 0.75, 0.5]:
            new_size = (int(img_copy.width * scale), int(img_copy.height * scale))
            img_resized = img_copy.resize(new_size, Image.Resampling.LANCZOS)
            for quality in range(90, MIN_QUALITY - 1, -10):
                img_resized.save(output_path, 'JPEG', quality=quality, optimize=True)
                size_kb = output_path.stat().st_size / 1024
                if min_kb <= size_kb <= max_kb:
                    best_attempt_path.unlink()
                    return "ok", size_kb
                if size_kb < best_attempt_path.stat().st_size / 1024:
                     shutil.copy(output_path, best_attempt_path)
        shutil.move(best_attempt_path, output_path)
        return "partial", output_path.stat().st_size / 1024
    except Exception:
        return "compress_failed", 0

def process_tree(source_dir, processed_dir, status_placeholder):
    """Processes all files, updating the Streamlit UI."""
    report_data = []
    status_placeholder.text("Starting file processing...")
    files_to_process = [p for p in source_dir.rglob('*') if p.is_file() and not p.name.lower().endswith('.zip')]
    progress_bar = st.progress(0)

    for i, input_path in enumerate(files_to_process):
        relative_path = input_path.relative_to(source_dir)
        status_placeholder.text(f"Processing: {relative_path}")
        original_kb = input_path.stat().st_size / 1024
        ext = input_path.suffix.lower()

        if ext in SUPPORTED_MEDIA_EXT:
            temp_jpg_path = Path(tempfile.gettempdir()) / f"temp_{relative_path.name}.jpg"
            if not convert_any_to_jpg(input_path, temp_jpg_path):
                report_data.append([relative_path, 'unknown', f"{original_kb:.2f}", 0, "convert_failed", "name"])
                continue
            category = guess_category_by_filename(input_path.name)
            min_kb, max_kb = TARGET_SIZES_KB.get(category, TARGET_SIZES_KB["unknown"])
            output_path = (processed_dir / relative_path).with_suffix(".jpg")
            output_path.parent.mkdir(parents=True, exist_ok=True)
            status, final_kb = compress_jpg_to_target(temp_jpg_path, output_path, min_kb, max_kb)
            report_data.append([relative_path, category, f"{original_kb:.2f}", f"{final_kb:.2f}", status, "name"])
            if temp_jpg_path.exists(): temp_jpg_path.unlink()
        else:
            output_path = processed_dir / relative_path
            output_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(input_path, output_path)
            report_data.append([relative_path, 'document', f"{original_kb:.2f}", f"{original_kb:.2f}", 'copied_as_is', 'name'])
        
        progress_bar.progress((i + 1) / len(files_to_process))

    status_placeholder.text("Processing complete!")
    return report_data

# --- STREAMLIT UI ---

st.set_page_config(page_title="Aiclex Media Processor", layout="wide")
st.title("🗂️ Aiclex Media Processor")
st.caption("Developed by Aiclex Technologies | aiclex.in")

if not PDF_SUPPORT:
    st.error("PDF processing is disabled. Please install 'poppler-utils' on the system to enable it.")

uploaded_file = st.file_uploader(
    "Apni .zip file yahaan upload karein",
    type=["zip"],
    help="Ek .zip file upload karein jismein aapke saare documents (images, PDFs, etc.) hon."
)

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        source_files_dir = temp_path / "source"
        processed_files_dir = temp_path / "processed"
        source_files_dir.mkdir()
        processed_files_dir.mkdir()

        input_zip_path = temp_path / uploaded_file.name
        with open(input_zip_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.info(f"Processing '{uploaded_file.name}'...")
        status_box = st.empty()

        with st.spinner('Extracting ZIP file(s)...'):
            with zipfile.ZipFile(input_zip_path, 'r') as zip_ref:
                zip_ref.extractall(source_files_dir)
            extract_all_zips(source_files_dir, status_box)

        report_rows = process_tree(source_files_dir, processed_files_dir, status_box)
        report_headers = ["path", "category", "original_kb", "final_kb", "status", "decided_via"]
        df_report = pd.DataFrame(report_rows, columns=report_headers)
        
        base_name = uploaded_file.name.replace('.zip', '')
        output_zip_path = temp_path / f"processed_{base_name}"
        shutil.make_archive(str(output_zip_path), 'zip', processed_files_dir)

        st.success("✅ Sabhi files process ho gayi hain!")
        st.header("Processing Report")
        st.dataframe(df_report)
        
        col1, col2 = st.columns(2)
        with open(f"{output_zip_path}.zip", "rb") as fp:
            col1.download_button(
                label="📂 Processed ZIP Download Karein",
                data=fp,
                file_name=f"processed_{uploaded_file.name}",
                mime="application/zip"
            )

        csv_data = df_report.to_csv(index=False).encode('utf-8')
        col2.download_button(
            label="📄 Report (CSV) Download Karein",
            data=csv_data,
            file_name=f"processing_report_{base_name}.csv",
            mime='text/csv'
        )
