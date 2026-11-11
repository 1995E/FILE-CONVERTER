import streamlit as st
from pdf2docx import Converter
import os
import tempfile

st.set_page_config(page_title="PDF to Word Converter", layout="centered")

st.title("📄 PDF to Editable Word Converter")
st.write("Convert your PDF files to editable DOCX format while preserving layouts, images, and text.")

uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file is not None:
    # Create a temporary file to save the uploaded PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(uploaded_file.read())
        temp_pdf_path = temp_pdf.name

    # Prepare the output DOCX path
    temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")

    if st.button("Convert to Word"):
        with st.spinner("Converting... This may take a moment depending on the file size."):
            try:
                # Initialize converter
                cv = Converter(temp_pdf_path)
                # Convert all pages (start=0, end=None)
                cv.convert(temp_docx_path, start=0, chunk_size=100)
                cv.close()
                
                st.success("Conversion completed successfully!")
                
                # Read the converted file for download
                with open(temp_docx_path, "rb") as docx_file:
                    docx_bytes = docx_file.read()
                
                # Provide download button
                st.download_button(
                    label="Download Word Document",
                    data=docx_bytes,
                    file_name=uploaded_file.name.replace(".pdf", ".docx"),
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"An error occurred during conversion: {e}")
            finally:
                # Clean up temporary files
                if os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)
                if os.path.exists(temp_docx_path):
                    os.remove(temp_docx_path)
