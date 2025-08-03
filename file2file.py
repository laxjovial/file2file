import streamlit as st
import pandas as pd
from io import BytesIO
import os
from docx import Document
from pdf2docx import Converter
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import pypandoc

# Ensure pandoc is installed
try:
    pypandoc.get_pandoc_path()
except OSError:
    pypandoc.download_pandoc()

st.set_page_config(page_title="File2File Converter", layout="centered")
st.title("📁 File2File Converter")
st.markdown("Convert between PDF, DOCX, TXT, CSV, XLS, XLSX. Batch uploads supported. Real PDF export.")

# Supported formats
doc_types = ["pdf", "docx", "txt"]
sheet_types = ["csv", "xls", "xlsx"]
all_types = doc_types + sheet_types

# Select format
source_format = st.selectbox("From format", all_types)
target_format = st.selectbox("To format", [f for f in all_types if f != source_format])

# Upload
uploaded_files = st.file_uploader("Upload files", type=[source_format], accept_multiple_files=True)

# Optional output file name
custom_name = st.text_input("Optional: base name for output file(s)", "converted")

# Preview section
def preview_file(file, file_type):
    st.subheader("🔍 Preview")
    if file_type in ["txt"]:
        text = file.read().decode("utf-8")
        st.text(text[:1000] + ("..." if len(text) > 1000 else ""))
    elif file_type in ["csv"]:
        df = pd.read_csv(file)
        st.dataframe(df.head())
    elif file_type in ["xls", "xlsx"]:
        df = pd.read_excel(file)
        st.dataframe(df.head())
    elif file_type == "docx":
        doc = Document(file)
        text = "\n".join([p.text for p in doc.paragraphs])
        st.text(text[:1000] + ("..." if len(text) > 1000 else ""))
    elif file_type == "pdf":
        with pdfplumber.open(file) as pdf:
            text = "\n".join([page.extract_text() or "" for page in pdf.pages])
            st.text(text[:1000] + ("..." if len(text) > 1000 else ""))

# DOC/TXT/PDF conversion
def convert_doc_file(uploaded_file, source, target):
    result = BytesIO()
    name = os.path.splitext(uploaded_file.name)[0]

    if source == "pdf":
        if target == "docx":
            with open("temp_input.pdf", "wb") as f:
                f.write(uploaded_file.read())
            cv = Converter("temp_input.pdf")
            cv.convert("temp_output.docx", start=0, end=None)
            cv.close()
            with open("temp_output.docx", "rb") as f:
                result.write(f.read())
            os.remove("temp_input.pdf")
            os.remove("temp_output.docx")
        elif target == "txt":
            with pdfplumber.open(uploaded_file) as pdf:
                text = "\n".join([page.extract_text() or "" for page in pdf.pages])
                result.write(text.encode("utf-8"))

    elif source == "docx":
        if target == "pdf":
            with open("temp.docx", "wb") as f:
                f.write(uploaded_file.read())
            pypandoc.convert_file("temp.docx", "pdf", outputfile="temp.pdf")
            with open("temp.pdf", "rb") as f:
                result.write(f.read())
            os.remove("temp.docx")
            os.remove("temp.pdf")
        elif target == "txt":
            doc = Document(uploaded_file)
            text = "\n".join([p.text for p in doc.paragraphs])
            result.write(text.encode("utf-8"))

    elif source == "txt":
        text = uploaded_file.read().decode("utf-8")
        if target == "pdf":
            c = canvas.Canvas(result, pagesize=letter)
            y = 750
            for line in text.splitlines():
                c.drawString(50, y, line)
                y -= 15
                if y < 50:
                    c.showPage()
                    y = 750
            c.save()
        elif target == "docx":
            doc = Document()
            for line in text.splitlines():
                doc.add_paragraph(line)
            doc.save(result)

    elif source == "docx" and target == "txt":
        doc = Document(uploaded_file)
        text = "\n".join([p.text for p in doc.paragraphs])
        result.write(text.encode("utf-8"))

    elif source == "txt" and target == "docx":
        text = uploaded_file.read().decode("utf-8")
        doc = Document()
        for line in text.splitlines():
            doc.add_paragraph(line)
        doc.save(result)

    result.seek(0)
    return result

# Spreadsheet conversion
def convert_sheet_file(uploaded_file, source, target):
    result = BytesIO()
    if source == "csv":
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    if target == "csv":
        df.to_csv(result, index=False)
    else:
        df.to_excel(result, index=False, engine="openpyxl" if target == "xlsx" else "xlrd")

    result.seek(0)
    return result

# Convert and download section
if uploaded_files:
    for idx, uploaded_file in enumerate(uploaded_files):
        st.divider()
        st.subheader(f"📄 File {idx + 1}: {uploaded_file.name}")
        preview_file(uploaded_file, source_format)

        with st.spinner("Converting..."):
            if source_format in doc_types and target_format in doc_types:
                output = convert_doc_file(uploaded_file, source_format, target_format)
            elif source_format in sheet_types and target_format in sheet_types:
                output = convert_sheet_file(uploaded_file, source_format, target_format)
            else:
                st.error("❌ Cross-type conversions (e.g., DOCX → CSV) not supported.")
                continue

        file_base = custom_name if custom_name else os.path.splitext(uploaded_file.name)[0]
        download_name = f"{file_base}_{idx + 1}.{target_format}" if len(uploaded_files) > 1 else f"{file_base}.{target_format}"

        st.success(f"✅ Done: {download_name}")
        st.download_button("⬇️ Download", data=output, file_name=download_name)

