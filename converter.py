
import streamlit as st
from pdf2docx import Converter
from docx import Document
from fpdf import FPDF
import os
import tempfile

st.set_page_config(page_title="Enterprise Doc Studio", layout="wide")

st.title("📝 Enterprise PDF-Word Workspace")
st.write("Convert PDF to Word, edit in-app with Microsoft styling alignments, track/revert drafts, and export back to PDF.")

# Initialize Session States
if "original_paragraphs" not in st.session_state:
    st.session_state.original_paragraphs = []
if "edited_paragraphs" not in st.session_state:
    st.session_state.edited_paragraphs = []
if "has_converted" not in st.session_state:
    st.session_state.has_converted = False
if "doc_name" not in st.session_state:
    st.session_state.doc_name = ""

uploaded_file = st.file_uploader("Upload Source PDF", type=["pdf"])

if uploaded_file is not None:
    if st.button("Convert and Initialize Workspace"):
        with st.spinner("Initializing document engine..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(uploaded_file.read())
                    temp_pdf_path = temp_pdf.name
                
                temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")
                
                # Convert using pdf2docx
                cv = Converter(temp_pdf_path)
                cv.convert(temp_docx_path, start=0, chunk_size=100)
                cv.close()
                
                # Load converted elements
                doc = Document(temp_docx_path)
                paragraphs = [p.text for p in doc.paragraphs if p.text.strip() != ""]
                
                st.session_state.original_paragraphs = list(paragraphs)
                st.session_state.edited_paragraphs = list(paragraphs)
                st.session_state.has_converted = True
                st.session_state.doc_name = uploaded_file.name.replace(".pdf", "")
                
                # Cleanup
                os.remove(temp_pdf_path)
                os.remove(temp_docx_path)
                st.success("Workspace initialized!")
            except Exception as e:
                st.error(f"Error initializing document workspace: {e}")

if st.session_state.has_converted:
    st.markdown("---")
    
    # Multi-panel Layout
    col_editor, col_preview = st.columns([1, 1])
    
    with col_editor:
        st.subheader("🛠️ Document Editor")
        
        # Edit paragraphs
        current_edits = []
        for i, para in enumerate(st.session_state.edited_paragraphs):
            val = st.text_area(f"Paragraph {i+1}", value=para, key=f"edit_{i}", height=100)
            current_edits.append(val)
            
        # Active Action Buttons
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            if st.button("💾 Save Draft Changes"):
                st.session_state.edited_paragraphs = list(current_edits)
                st.success("Draft changes successfully saved in-memory!")
        with btn_col2:
            if st.button("↩️ Revert (Unsave) All Changes"):
                st.session_state.edited_paragraphs = list(st.session_state.original_paragraphs)
                st.info("Reverted to the original document state!")
                st.rerun()

    with col_preview:
        st.subheader("👁️ Live Word Document Preview")
        
        preview_content = "\n\n".join(st.session_state.edited_paragraphs)
        st.text_area("ReadOnly Live Preview (Page layout representation)", value=preview_content, height=450, disabled=True)
        
        st.subheader("📤 Export Center")
        
        # EXPORT 1: Word Document
        if st.button("Compile & Export to Word (DOCX)"):
            try:
                out_doc = Document()
                for p in st.session_state.edited_paragraphs:
                    if p.strip():
                        out_doc.add_paragraph(p)
                
                docx_temp = tempfile.mktemp(suffix=".docx")
                out_doc.save(docx_temp)
                
                with open(docx_temp, "rb") as f:
                    st.download_button(
                        label="Download .DOCX",
                        data=f.read(),
                        file_name=f"{st.session_state.doc_name}_modified.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"Failed to compile Word doc: {e}")
                
        # EXPORT 2: Convert edited document to PDF
        if st.button("Compile & Export to PDF"):
            try:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Helvetica", size=12)
                
                for p in st.session_state.edited_paragraphs:
                    if p.strip():
                        # Clean unicode characters for compatibility with FPDF basic font
                        clean_p = p.encode('latin-1', 'ignore').decode('latin-1')
                        pdf.multi_cell(0, 10, clean_p)
                        pdf.ln(4)
                
                pdf_temp = tempfile.mktemp(suffix=".pdf")
                pdf.output(pdf_temp)
                
                with open(pdf_temp, "rb") as f:
                    st.download_button(
                        label="Download .PDF",
                        data=f.read(),
                        file_name=f"{st.session_state.doc_name}_modified.pdf",
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Failed to compile PDF: {e}")
