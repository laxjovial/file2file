import streamlit as st
import pandas as pd
import os
from io import BytesIO
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

st.title("📁 File2File Converter")
st.markdown("Convert between PDF, DOCX, TXT, CSV, XLS, XLSX. All conversions are real and bidirectional.")

doc_types = ["pdf", "docx", "txt"]
sheet_types = ["csv", "xls", "xlsx"]
all_types = doc_types + sheet_types

source_format = st.selectbox("From format", all_types)
target_format = st.selectbox("To format", [f for f in all_types if f != source_format])

uploaded_file = st.file_uploader("Upload your file", type=[source_format])

def convert_doc_file(uploaded_file, source, target):
    result = BytesIO()

    # Read input content
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
            output = pypandoc.convert_file("temp.docx", "pdf", outputfile="temp.pdf")
            with open("temp.pdf", "rb") as f:
                result.write(f.read())
            os.remove("temp.docx")
            os.remove("temp.pdf")

        elif target == "txt":
            doc = Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
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

    elif source == "txt" and target == "docx":
        text = uploaded_file.read().decode("utf-8")
        doc = Document()
        for line in text.splitlines():
            doc.add_paragraph(line)
        doc.save(result)

    elif source == "docx" and target == "txt":
        doc = Document(uploaded_file)
        text = "\n".join([p.text for p in doc.paragraphs])
        result.write(text.encode("utf-8"))

    elif source == "pdf" and target == "txt":
        with pdfplumber.open(uploaded_file) as pdf:
            text = "\n".join([page.extract_text() or "" for page in pdf.pages])
            result.write(text.encode("utf-8"))

    result.seek(0)
    return result

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

if uploaded_file:
    st.info("Converting...")

    if source_format in doc_types and target_format in doc_types:
        output = convert_doc_file(uploaded_file, source_format, target_format)
    elif source_format in sheet_types and target_format in sheet_types:
        output = convert_sheet_file(uploaded_file, source_format, target_format)
    else:
        st.error("❌ Cross-type conversions (e.g., DOCX → CSV) are not supported.")
        output = None

    if output:
        download_name = f"converted.{target_format}"
        st.success("✅ Conversion complete!")
        st.download_button("⬇️ Download File", data=output, file_name=download_name)
