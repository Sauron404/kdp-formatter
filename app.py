import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
import tempfile
import os

def format_docx(uploaded_file, formato="cartaceo"):
    doc = Document(uploaded_file)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.87)
        section.right_margin = Inches(0.67)
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Georgia'
            run.font.size = Pt(12)
    doc.paragraphs[0].insert_paragraph_before("Titolo del Libro", style='Title')
    doc.paragraphs[1].insert_paragraph_before("Autore", style='Subtitle')
    return doc

st.title("📘 KDP Formatter 6x9")
st.write("Carica il tuo file Word (.docx) e scarica una versione formattata pronta per KDP.")
uploaded_file = st.file_uploader("Carica un file .docx", type=["docx"])
formato = st.selectbox("Formato desiderato:", ["cartaceo", "ebook"])

if uploaded_file:
    if st.button("📄 Formatta il documento"):
        doc = format_docx(uploaded_file, formato=formato)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            tmp_path = tmp.name
        with open(tmp_path, "rb") as f:
            st.download_button(
                label="📥 Scarica il file formattato (.docx)",
                data=f,
                file_name="kdp_formattato.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        os.remove(tmp_path)
