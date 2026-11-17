%%writefile app.py
import streamlit as st
from pdf2docx import Converter
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF
import os
import tempfile

st.set_page_config(page_title="Microsoft Word Cloud Studio", layout="wide")

st.title("💻 Microsoft Word Cloud Platform & PDF Engine")
st.write("Experience a rich Microsoft Word environment in-browser. Customize layout designs, paragraph rules, and fonts, preview live modifications, and export with MS Word styles preserved.")

# Initialize Session States
if "original_paragraphs" not in st.session_state:
    st.session_state.original_paragraphs = []
if "edited_paragraphs" not in st.session_state:
    st.session_state.edited_paragraphs = []
if "has_converted" not in st.session_state:
    st.session_state.has_converted = False
if "doc_name" not in st.session_state:
    st.session_state.doc_name = ""

uploaded_file = st.file_uploader("Upload Source PDF File", type=["pdf"])

if uploaded_file is not None:
    if st.button("Launch Interactive Word Engine"):
        with st.spinner("Initializing document engine with Microsoft Word environments..."):
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
                st.success("Workspace loaded successfully!")
            except Exception as e:
                st.error(f"Error initializing document workspace: {e}")

if st.session_state.has_converted:
    st.markdown("---")

    # Word Sidebar Layout Engine
    st.sidebar.subheader("📂 Page Layout Settings")
    font_style = st.sidebar.selectbox("Primary Font Family", ["Calibri", "Arial", "Times New Roman", "Garamond", "Courier New"])
    font_size = st.sidebar.slider("Font Size (pt)", 8, 28, 11)
    text_color = st.sidebar.color_picker("Font Color", "#333333")
    bg_color = st.sidebar.color_picker("Page Background Color", "#ffffff")
    
    st.sidebar.subheader("↔️ Paragraph Formatting")
    alignment = st.sidebar.selectbox("Text Alignment", ["Justified", "Left", "Center", "Right"])
    line_spacing = st.sidebar.selectbox("Line Spacing (MS-Standard)", [1.15, 1.0, 1.5, 2.0])
    paragraph_spacing = st.sidebar.slider("Spacing After Paragraph (px)", 0, 30, 10)

    # Side-By-Side Design Studio
    col_editor, col_preview = st.columns([1, 1])

    with col_editor:
        st.subheader("🛠️ Microsoft Word Live Editor")
        
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
                st.success("Draft updated successfully!")
        with btn_col2:
            if st.button("↩️ Revert / Unsave Changes", use_container_width=True):
                st.session_state.edited_paragraphs = list(st.session_state.original_paragraphs)
                st.info("Reverted all changes back to converted state!")
                st.rerun()

    with col_preview:
        st.subheader("👁️ Live Word Canvas Preview")

        # Build a stylized HTML previsualization block mimicking a physical MS Word sheet layout
        align_map = {"Justified": "justify", "Left": "left", "Center": "center", "Right": "right"}
        css_align = align_map[alignment]
        
        preview_html = f"""
        <div style="background-color: {bg_color}; border: 1px solid #c8c8c8; box-shadow: 0px 4px 16px rgba(0,0,0,0.1); padding: 40px; min-height: 500px; max-height: 600px; overflow-y: auto; color: {text_color}; font-family: '{font_style}', sans-serif; font-size: {font_size}pt; line-height: {line_spacing};">
            <h2 style="border-bottom: 2px solid #0078d4; padding-bottom: 8px; color: #0078d4; font-family: inherit; margin-top: 0;">{st.session_state.doc_name}</h2>
        """
        for p in st.session_state.edited_paragraphs:
            if p.strip():
                preview_html += f"<p style='margin-bottom: {paragraph_spacing}px; text-align: {css_align};'>{p}</p>"
        preview_html += "</div>"

        st.markdown(preview_html, unsafe_html=True)
        st.write("")

        st.subheader("📥 Export Center")

        # EXPORT 1: Word Document with Custom styles applied natively
        if st.button("Export to Microsoft Word (.DOCX)", use_container_width=True):
            try:
                out_doc = Document()
                align_docx = {
                    "Justified": WD_ALIGN_PARAGRAPH.JUSTIFY,
                    "Left": WD_ALIGN_PARAGRAPH.LEFT,
                    "Center": WD_ALIGN_PARAGRAPH.CENTER,
                    "Right": WD_ALIGN_PARAGRAPH.RIGHT
                }
                
                for p in st.session_state.edited_paragraphs:
                    if p.strip():
                        para_obj = out_doc.add_paragraph()
                        para_obj.alignment = align_docx[alignment]
                        para_obj.paragraph_format.line_spacing = line_spacing
                        
                        run = para_obj.add_run(p)
                        run.font.name = font_style
                        run.font.size = Pt(font_size)
                        
                        # Hex to RGB color conversion for docx
                        h = text_color.lstrip('#')
                        rgb = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
                        run.font.color.rgb = RGBColor(*rgb)

                docx_temp = tempfile.mktemp(suffix=".docx")
                out_doc.save(docx_temp)

                with open(docx_temp, "rb") as f:
                    st.download_button(
                        label="📥 Download Custom Word File (.docx)",
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
                pdf.set_font("Helvetica", size=font_size)
                
                # Extract RGB from text_color color picker
                h = text_color.lstrip('#')
                r, g, b = tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
                pdf.set_text_color(r, g, b)
                
                for p in st.session_state.edited_paragraphs:
                    if p.strip():
                        clean_p = p.encode("latin-1", "ignore").decode("latin-1")
                        # Calculate alignment
                        pdf_align = "J" if alignment == "Justified" else alignment[0]
                        pdf.multi_cell(0, 10 * (line_spacing / 1.15), clean_p, align=pdf_align)
                        pdf.ln(paragraph_spacing / 2)

                pdf_temp = tempfile.mktemp(suffix=".pdf")
                pdf.output(pdf_temp)

                with open(pdf_temp, "rb") as f:
                    st.download_button(
                        label="📥 Download Custom PDF File (.pdf)",
                        data=f.read(),
                        file_name=f"{st.session_state.doc_name}_edited.pdf",
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Failed to compile PDF: {e}")
