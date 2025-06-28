def apply_style_safe(paragraph, doc, style_name, fallback='Normal'):
    try:
        paragraph.style = style_name
    except KeyError:
        paragraph.style = fallback

def format_docx(uploaded_file, formato="cartaceo", frontespizio=True, numeri_pagina=True, titolo_libro="Titolo del Libro", autore_libro="Autore", editore="Nome Editore"):
    doc = Document(uploaded_file)

    # Imposta margini
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.87)
        section.right_margin = Inches(0.67)
        if numeri_pagina and formato == "cartaceo":
            add_page_numbers(section)

    # Applica font, giustificazione e formatta titoli
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip().lower()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if text.startswith("capitolo"):
            apply_style_safe(paragraph, doc, 'Heading 1')
        elif text.startswith("sezione"):
            apply_style_safe(paragraph, doc, 'Heading 2')
        else:
            apply_style_safe(paragraph, doc, 'Normal')

        for run in paragraph.runs:
            run.font.name = 'Georgia'
            run.font.size = Pt(12)

    # Frontespizio
    if frontespizio:
        doc._body.clear_content()
        add_centered_page(doc, [titolo_libro, autore_libro])
        add_centered_page(doc, [f"\u00a9 2025 {editore}", "Tutti i diritti riservati"])
        doc.add_page_break()
        doc.add_paragraph("Indice")
        add_table_of_contents(doc)
        doc.add_page_break()
        doc.add_paragraph("Inizio contenuto del libro...")

    return doc
