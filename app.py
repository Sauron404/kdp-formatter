import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
import tempfile
import os

try:
    from docx2pdf import convert as docx_to_pdf
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

# Funzione per applicare formattazione KDP
def format_docx(uploaded_file, formato="cartaceo", frontespizio=True, numeri_pagina=True):
    doc = Document(uploaded_file)

    # Imposta margini
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.87)
        section.right_margin = Inches(0.67)

    # Applica font
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Georgia'
            run.font.size = Pt(12)

    # Inserisci frontespizio
    if frontespizio:
        doc.paragraphs[0].insert_paragraph_before("Titolo del Libro")
        doc.paragraphs[1].insert_paragraph_before("Autore")

    return doc

st.set_page_config(page_title="KDP Formatter", layout="centered")
st.title("📘 KDP Formatter 6x9")
st.write("Carica un file Word `.docx` e ottieni un file formattato per KDP, pronto per la stampa o l'eBook.")

uploaded_file = st.file_uploader("📤 Carica il tuo file Word", type=["docx"])
formato = st.selectbox("🖋️ Formato desiderato:", ["cartaceo", "ebook"])
add_frontespizio = st.checkbox("Aggiungi frontespizio?", value=True)
add_numeri_pagina = st.checkbox("Aggiungi numeri di pagina?", value=True)

if uploaded_file:
    if st.button("📄 Formatta il documento"):
        with st.spinner("Formattazione in corso..."):
            doc = format_docx(uploaded_file, formato=formato, frontespizio=add_frontespizio, numeri_pagina=add_numeri_pagina)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                docx_path = tmp.name

            with open(docx_path, "rb") as f:
                st.success("✅ Documento Word formattato con successo!")
                st.download_button(
                    label="📥 Scarica .DOCX",
                    data=f,
                    file_name="kdp_formattato.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            # Genera PDF se possibile
            if DOCX2PDF_AVAILABLE:
                with tempfile.TemporaryDirectory() as tmpdir:
                    pdf_path = os.path.join(tmpdir, "output.pdf")
                    docx_to_pdf(docx_path, pdf_path)
                    with open(pdf_path, "rb") as pdf_file:
                        st.download_button(
                            label="📄 Scarica anche in PDF",
                            data=pdf_file,
                            file_name="kdp_formattato.pdf",
                            mime="application/pdf"
                        )

            os.remove(docx_path)
