
import streamlit as st
from pdf2docx import Converter
import os
import tempfile
from docx import Document

st.set_page_config(page_title="PDF to Word Converter & Editor", layout="wide")

st.title("📄 PDF to Editable Word Converter & Platform")
st.write("Convert your PDF files to Word, edit the content directly in this browser interface, and export the customized DOCX file.")

# Initialize session state variables to hold document content and state
if "doc_paragraphs" not in st.session_state:
    st.session_state.doc_paragraphs = []
if "conversion_done" not in st.session_state:
    st.session_state.conversion_done = False
if "file_name" not in st.session_state:
    st.session_state.file_name = ""

uploaded_file = st.file_uploader("Choose a PDF file to begin", type=["pdf"])

if uploaded_file is not None:
    if st.button("Convert and Extract Text for Editing"):
        with st.spinner("Converting PDF... This may take a moment."):
            try:
                # Set up temporary files
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(uploaded_file.read())
                    temp_pdf_path = temp_pdf.name

                temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")

                # Convert PDF to Word document
                cv = Converter(temp_pdf_path)
                cv.convert(temp_docx_path, start=0, chunk_size=100)
                cv.close()

                # Read content from converted Word document
                doc = Document(temp_docx_path)
                # Save original paragraphs to session state
                st.session_state.doc_paragraphs = [p.text for p in doc.paragraphs if p.text.strip() != ""]
                st.session_state.conversion_done = True
                st.session_state.file_name = uploaded_file.name.replace(".pdf", "_edited.docx")

                # Cleanup temporary files
                if os.path.exists(temp_pdf_path):
                    os.remove(temp_pdf_path)
                if os.path.exists(temp_docx_path):
                    os.remove(temp_docx_path)

                st.success("Document successfully converted! You can now edit the content below.")
            except Exception as e:
                st.error(f"An error occurred during conversion: {e}")

# Editor Interface
if st.session_state.conversion_done:
    st.subheader("✍️ Live Document Editor")
    st.info("Modify the paragraphs below. Any changes made will be packed into the final exported Word document.")
    
    edited_paragraphs = []
    # Display and allow editing of each paragraph
    for i, para in enumerate(st.session_state.doc_paragraphs):
        edited_text = st.text_area(f"Paragraph {i+1}", value=para, key=f"para_{i}", height=100)
        edited_paragraphs.append(edited_text)

    # Save and Export action
    if st.button("Rebuild & Download DOCX"):
        try:
            output_doc = Document()
            for para in edited_paragraphs:
                if para.strip():
                    output_doc.add_paragraph(para)

            # Save the newly constructed document
            export_path = tempfile.mktemp(suffix=".docx")
            output_doc.save(export_path)

            with open(export_path, "rb") as f:
                docx_bytes = f.read()

            st.download_button(
                label="Download Final Edited Word Document",
                data=docx_bytes,
                file_name=st.session_state.file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            if os.path.exists(export_path):
                os.remove(export_path)
        except Exception as e:
            st.error(f"An error occurred while compiling the final document: {e}")
