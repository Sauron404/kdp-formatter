import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
import tempfile
import os

# Funzione per applicare formattazione KDP
def format_docx(uploaded_file, formato="cartaceo"):
    doc = Document(uploaded_file)

    # Margini per formato cartaceo KDP 6x9
    for section in doc.sections:
        section.top_margin = Inches(0.79)     # ≈2 cm
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.87)    # ≈2.2 cm
        section.right_margin = Inches(0.67)   # ≈1.7 cm

    # Font e dimensione uniforme
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Georgia'
            run.font.size = Pt(12)

    # Inserimento frontespizio semplice (senza stile personalizzato)
    doc.paragraphs[0].insert_paragraph_before("Titolo del Libro")
    doc.paragraphs[1].insert_paragraph_before("Autore")

    return doc

# --- Streamlit App ---
st.set_page_config(page_title="KDP Formatter", layout="centered")
st.title("📘 KDP Formatter 6x9")
st.write("Carica un file Word `.docx` e ottieni un file formattato per la stampa su KDP.")

uploaded_file = st.file_uploader("📤 Carica il tuo file Word", type=["docx"])
formato = st.selectbox("🖋️ Formato desiderato:", ["cartaceo", "ebook"])

