import streamlit as st
import pandas as pd
from io import BytesIO
import os

from docx import Document
from PyPDF2 import PdfReader
from fpdf import FPDF

# Set title
st.title("📁 File2File Converter")
st.write("Upload a file and convert it to another format.")

# Define supported formats
doc_types = ["pdf", "docx", "txt"]
sheet_types = ["csv", "xls", "xlsx"]

all_types = doc_types + sheet_types

# Select source and target formats
source_format = st.selectbox("Select the source format", all_types)
target_format = st.selectbox("Select the target format", [f for f in all_types if f != source_format])

# File uploader
uploaded_file = st.file_uploader(f"Upload a {source_format.upper()} file", type=[source_format])

# Convert to bytes for download
def download_button(data, filename, label):
    st.download_button(label, data=data, file_name=filename)

# Helper: Convert PDF to text
def pdf_to_text(file):
    reader = PdfReader(file)
    text = "\n".join([page.extract_text() or '' for page in reader.pages])
    return text

# Helper: Text to PDF using FPDF
def text_to_pdf(text):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=12)
    for line in text.splitlines():
        pdf.cell(0, 10, txt=line, ln=True)
    output = BytesIO()
    pdf.output(output)
    return output.getvalue()

# Handle conversion logic
def convert_file(uploaded_file, source_format, target_format):
    result = BytesIO()

    # Text-based conversions
    if source_format in doc_types and target_format in doc_types:
        text = ""

        # Read original
        if source_format == "pdf":
            text = pdf_to_text(uploaded_file)
        elif source_format == "docx":
            doc = Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
        elif source_format == "txt":
            text = uploaded_file.read().decode("utf-8")

        # Write to target
        if target_format == "txt":
            result.write(text.encode("utf-8"))
        elif target_format == "docx":
            doc = Document()
            for line in text.splitlines():
                doc.add_paragraph(line)
            doc.save(result)
        elif target_format == "pdf":
            pdf_bytes = text_to_pdf(text)
            result.write(pdf_bytes)

    # Spreadsheet conversions
    elif source_format in sheet_types and target_format in sheet_types:
        # Read source
        if source_format == "csv":
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        # Write to target
        if target_format == "csv":
            df.to_csv(result, index=False)
        else:
            df.to_excel(result, index=False, engine="openpyxl" if target_format == "xlsx" else "xlrd")

    else:
        st.error("❌ This type of conversion is not supported.")
        return None, None

    result.seek(0)
    out_ext = f".{target_format}"
    out_name = f"converted{out_ext}"
    return result, out_name

# When file is uploaded, perform conversion
if uploaded_file:
    output, filename = convert_file(uploaded_file, source_format, target_format)
    if output:
        st.success("✅ Conversion successful!")
        download_button(output, filename, "⬇️ Download Converted File")
