import io
import math
import os
import tempfile
from typing import List, Tuple

import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

try:
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt
    WORD_AVAILABLE = True
except:
    WORD_AVAILABLE = False


st.set_page_config(page_title="Anexos de videoscopía", layout="wide")

PAGE_W, PAGE_H = letter
COLUMNS = 2
ROWS = 4
MAX_PER_PAGE = 8
DEFAULT_LOGO_PATH = "/mnt/data/image.png"


def mm_to_pt(mm):
    return mm * 72 / 25.4


OUTER_MARGIN = mm_to_pt(12)
GAP_X = mm_to_pt(6)
GAP_Y = mm_to_pt(6)
FOOTER_H = mm_to_pt(12)
DESC_H = mm_to_pt(24)

USABLE_W = PAGE_W - (2 * OUTER_MARGIN)
USABLE_H = PAGE_H - (2 * OUTER_MARGIN) - FOOTER_H - DESC_H

BOX_W = (USABLE_W - GAP_X) / 2
BOX_H = (USABLE_H - (GAP_Y * 3)) / 4


def abrir_imagen(file):
    if hasattr(file, "seek"):
        file.seek(0)
    img = Image.open(file)
    img = ImageOps.exif_transpose(img)
    return img.convert("RGB")


def calcular_ajuste(w, h, box_w, box_h):
    ratio = min(box_w / w, box_h / h)
    return w * ratio, h * ratio


def wrap_text(text, max_chars=110, max_lines=5):
    words = text.split()
    lines = []
    current = ""

    for w in words:
        test = f"{current} {w}".strip()
        if len(test) <= max_chars:
            current = test
        else:
            lines.append(current)
            current = w
            if len(lines) >= max_lines:
                break

    if current and len(lines) < max_lines:
        lines.append(current)

    return lines[:max_lines]


def obtener_logo(logo_file):
    if logo_file:
        return logo_file
    if os.path.exists(DEFAULT_LOGO_PATH):
        return DEFAULT_LOGO_PATH
    return None


def preparar_logo(logo_file):
    src = obtener_logo(logo_file)
    if not src:
        return None, 0, 0

    img = abrir_imagen(src)
    w, h = calcular_ajuste(img.width, img.height, mm_to_pt(24), mm_to_pt(10))
    return ImageReader(img), w, h


def dibujar_descripcion(pdf, cilindro, descripcion):
    x = OUTER_MARGIN
    y = PAGE_H - OUTER_MARGIN - mm_to_pt(2)

    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(x, y, f"CILINDRO {cilindro or 'X'}:")

    pdf.setFont("Helvetica", 9)
    lines = wrap_text(descripcion)

    for i, line in enumerate(lines):
        pdf.drawString(x, y - mm_to_pt(5) - i * mm_to_pt(4), line)


def dibujar_footer(pdf, campo, logo_reader, lw, lh):
    y = OUTER_MARGIN + mm_to_pt(1)

    pdf.setFont("Helvetica", 9)
    pdf.drawString(OUTER_MARGIN, y, "Lubricantes Mobil")

    if campo:
        pdf.drawCentredString(PAGE_W / 2, y, campo)

    if logo_reader:
        pdf.drawImage(logo_reader, PAGE_W - OUTER_MARGIN - lw, OUTER_MARGIN, width=lw, height=lh)


def generar_pdf(registros, campo, logo_file, cilindro, descripcion):
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)

    logo_reader, lw, lh = preparar_logo(logo_file)

    total_paginas = math.ceil(len(registros) / MAX_PER_PAGE)

    for p in range(total_paginas):
        lote = registros[p * 8:(p + 1) * 8]

        dibujar_descripcion(pdf, cilindro, descripcion)

        for i, item in enumerate(lote):
            row = i // 2
            col = i % 2

            x = OUTER_MARGIN + col * (BOX_W + GAP_X)
            y_top = PAGE_H - OUTER_MARGIN - DESC_H - row * (BOX_H + GAP_Y)
            y = y_top - BOX_H

            img = abrir_imagen(item)
            w, h = calcular_ajuste(img.width, img.height, BOX_W, BOX_H)

            pdf.drawImage(ImageReader(img), x + (BOX_W - w)/2, y + (BOX_H - h)/2, width=w, height=h)

        dibujar_footer(pdf, campo, logo_reader, lw, lh)
        pdf.showPage()

    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()


def generar_word(registros, campo, logo_file, cilindro, descripcion):
    if not WORD_AVAILABLE:
        return b""

    doc = Document()
    sec = doc.sections[0]

    sec.top_margin = Inches(0.45)
    sec.bottom_margin = Inches(0.45)
    sec.left_margin = Inches(0.45)
    sec.right_margin = Inches(0.45)

    total_paginas = math.ceil(len(registros) / 8)

    for p in range(total_paginas):
        lote = registros[p * 8:(p + 1) * 8]

        p_desc = doc.add_paragraph()
        r1 = p_desc.add_run(f"CILINDRO {cilindro or 'X'}: ")
        r1.bold = True
        p_desc.add_run(descripcion)

        tabla = doc.add_table(rows=4, cols=2)

        for i, img_file in enumerate(lote):
            r = i // 2
            c = i % 2

            cell = tabla.cell(r, c)
            cell.text = ""
            p_img = cell.paragraphs[0]

            img = abrir_imagen(img_file)

            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            img.save(tmp.name)

            p_img.add_run().add_picture(tmp.name, width=Inches(3.1))

        footer = doc.add_table(rows=1, cols=3)

        footer.cell(0,0).text = "Lubricantes Mobil"
        footer.cell(0,1).text = campo or ""

        if p < total_paginas - 1:
            doc.add_section(WD_SECTION.NEW_PAGE)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# UI
st.title("Generador de anexos")

campo = st.text_input("Campo")
cilindro = st.text_input("Cilindro")
descripcion = st.text_area("Descripción")

logo = st.file_uploader("Logo", type=["png","jpg","jpeg"])
files = st.file_uploader("Imágenes", accept_multiple_files=True)

if files:
    pdf = generar_pdf(files, campo, logo, cilindro, descripcion)

    if WORD_AVAILABLE:
        word = generar_word(files, campo, logo, cilindro, descripcion)

        col1, col2 = st.columns(2)
        col1.download_button("PDF", pdf, "anexos.pdf")
        col2.download_button("Word", word, "anexos.docx")
    else:
        st.download_button("PDF", pdf, "anexos.pdf")
