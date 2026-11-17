
import streamlit as st
from pdf2docx import Converter
from docx import Document
from fpdf import FPDF
import os
import tempfile

st.set_page_config(page_title="Microsoft Web Doc Studio", layout="wide")

st.title("📝 Microsoft Web Document Workspace (PDF-Word Engine)")
st.write("Convert, preview, live-edit with simulated Microsoft layout settings, maintain drafts, and export back to Microsoft DOCX or PDF.")

# Initialize Session States
if "original_paragraphs" not in st.session_state:
    st.session_state.original_paragraphs = []
if "edited_paragraphs" not in st.session_state:
    st.session_state.edited_paragraphs = []
if "has_converted" not in st.session_state:
    st.session_state.has_converted = False
if "doc_name" not in st.session_state:
    st.session_state.doc_name = ""
if "font_style" not in st.session_state:
    st.session_state.font_style = "Calibri"
if "font_size" not in st.session_state:
    st.session_state.font_size = 11

uploaded_file = st.file_uploader("Upload Source PDF File", type=["pdf"])

if uploaded_file is not None:
    if st.button("Initialize Word Workspace"):
        with st.spinner("Initializing document engine with Microsoft Word alignments..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(uploaded_file.read())
                    temp_pdf_path = temp_pdf.name
                
                temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")
                
                # Convert using pdf2docx
                cv = Converter(temp_pdf_path)
                cv.convert(temp_docx_path, start=0, chunk_size=100)
                cv.close()
                
                # Load converted elements into memory
                doc = Document(temp_docx_path)
                paragraphs = [p.text for p in doc.paragraphs if p.text.strip() != ""]
                
                st.session_state.original_paragraphs = list(paragraphs)
                st.session_state.edited_paragraphs = list(paragraphs)
                st.session_state.has_converted = True
                st.session_state.doc_name = uploaded_file.name.replace(".pdf", "")
                
                # Cleanup
                os.remove(temp_pdf_path)
                os.remove(temp_docx_path)
                st.success("Workspace successfully loaded!")
            except Exception as e:
                st.error(f"Error initializing document workspace: {e}")

if st.session_state.has_converted:
    st.markdown("---")
    
    # Style customizer sidebar for MS fonts
    st.sidebar.subheader("Microsoft Office Typography")
    st.session_state.font_style = st.sidebar.selectbox("Primary Font Family", ["Calibri", "Arial", "Times New Roman", "Courier New"])
    st.session_state.font_size = st.sidebar.slider("Font Size (pt)", 8, 24, 11)
    
    # Multi-panel Layout
    col_editor, col_preview = st.columns([1, 1])
    
    with col_editor:
        st.subheader("🛠️ In-Platform Document Editor")
        
        # Edit paragraphs
        current_edits = []
        for i, para in enumerate(st.session_state.edited_paragraphs):
            val = st.text_area(f"Paragraph {i+1}", value=para, key=f"edit_{i}", height=110)
            current_edits.append(val)
            
        # Draft State Controls
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            if st.button("💾 Save Draft Changes", use_container_width=True):
                st.session_state.edited_paragraphs = list(current_edits)
                st.success("Draft updated!")
        with btn_col2:
            if st.button("↩️ Revert / Unsave Changes", use_container_width=True):
                st.session_state.edited_paragraphs = list(st.session_state.original_paragraphs)
                st.info("Reverted all changes back to converted state!")
                st.rerun()

    with col_preview:
        st.subheader("👁️ Live Microsoft Word Preview")
        
        # Build a stylized HTML previsualization block mimicking a Word layout sheet
        preview_html = f"""
        <div style="background-color: white; border: 1px solid #dcdcdc; box-shadow: 1px 1px 8px #eee; padding: 30px; min-height: 400px; max-height: 600px; overflow-y: auto; color: black; font-family: '{st.session_state.font_style}', sans-serif; font-size: {st.session_state.font_size}pt; line-height: 1.5;">
            <h3 style="border-bottom: 2px solid #0078d4; padding-bottom: 5px; color: #0078d4;">{st.session_state.doc_name} (Draft Preview)</h3>
        """
        for p in st.session_state.edited_paragraphs:
            if p.strip():
                preview_html += f"<p style='margin-bottom: 12px; text-align: justify;'>{p}</p>"
        preview_html += "</div>"
        
        st.markdown(preview_html, unsafe_html=True)
        st.write("")
        
        st.subheader("📥 Export Center")
        
        # EXPORT 1: Word Document
        if st.button("Export to Microsoft Word (.DOCX)", use_container_width=True):
            try:
                out_doc = Document()
                for p in st.session_state.edited_paragraphs:
                    if p.strip():
                        para_obj = out_doc.add_paragraph(p)
                        run = para_obj.runs[0] if para_obj.runs else para_obj.add_run()
                        run.font.name = st.session_state.font_style
                        run.font.size = st.session_state.font_size * 12700 * 10  # convert pt to English Metric Units internally
                
                docx_temp = tempfile.mktemp(suffix=".docx")
                out_doc.save(docx_temp)
                
                with open(docx_temp, "rb") as f:
                    st.download_button(
                        label="📥 Download Word File (.docx)",
                        data=f.read(),
                        file_name=f"{st.session_state.doc_name}_edited.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"Failed to compile Word document: {e}")
                
        # EXPORT 2: Export final layout directly to PDF
        if st.button("Export directly to PDF (.PDF)", use_container_width=True):
            try:
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Helvetica", size=st.session_state.font_size)
                
                for p in st.session_state.edited_paragraphs:
                    if p.strip():
                        clean_p = p.encode('latin-1', 'ignore').decode('latin-1')
                        pdf.multi_cell(0, 10, clean_p)
                        pdf.ln(4)
                
                pdf_temp = tempfile.mktemp(suffix=".pdf")
                pdf.output(pdf_temp)
                
                with open(pdf_temp, "rb") as f:
                    st.download_button(
                        label="📥 Download PDF File (.pdf)",
                        data=f.read(),
                        file_name=f"{st.session_state.doc_name}_edited.pdf",
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Failed to compile PDF: {e}")
