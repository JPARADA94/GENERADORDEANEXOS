import io
import math
import os
import tempfile
from typing import List, Tuple

import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

try:
    from docx import Document
    from docx.enum.section import WD_SECTION
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt
    WORD_AVAILABLE = True
except Exception:
    WORD_AVAILABLE = False

# =========================================================
# App: Generador de anexos de videoscopía
# Ajustes solicitados:
# - SIN título de hoja
# - Párrafo superior de hallazgos: "CILINDRO X:" en negrilla + texto
# - Máx 5 renglones (controlado por altura)
# - Sin línea negra inferior
# =========================================================

st.set_page_config(page_title="Anexos de videoscopía", page_icon="🖼️", layout="wide")

PAGE_W, PAGE_H = letter
COLUMNS = 2
ROWS = 4
MAX_PER_PAGE = COLUMNS * ROWS
DEFAULT_LOGO_PATH = "/mnt/data/image.png"


def mm_to_pt(mm: float) -> float:
    return mm * 72 / 25.4


OUTER_MARGIN = mm_to_pt(12)
GAP_X = mm_to_pt(6)
GAP_Y = mm_to_pt(8)
FOOTER_H = mm_to_pt(14)
DESC_H = mm_to_pt(32)  # espacio para ~5 renglones
TOP_EXTRA = mm_to_pt(6)

USABLE_W = PAGE_W - (2 * OUTER_MARGIN)
USABLE_H = PAGE_H - (2 * OUTER_MARGIN) - FOOTER_H - DESC_H - TOP_EXTRA
BOX_W = (USABLE_W - GAP_X) / COLUMNS
CELL_H = (USABLE_H - (GAP_Y * (ROWS - 1))) / ROWS
IMAGE_BOX_H = CELL_H


# -----------------------------
# Utilidades
# -----------------------------

def natural_key(name: str):
    parts, current, is_digit = [], "", None
    for ch in name.lower():
        if ch.isdigit():
            if is_digit is False:
                parts.append(current)
                current = ""
            current += ch
            is_digit = True
        else:
            if is_digit is True:
                parts.append(int(current))
                current = ""
            current += ch
            is_digit = False
    if current:
        parts.append(int(current) if is_digit else current)
    return parts


def ordenar_archivos(files) -> List:
    return sorted(files, key=lambda f: natural_key(f.name))


def abrir_imagen(file) -> Image.Image:
    if hasattr(file, "seek"):
        file.seek(0)
    img = Image.open(file)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    elif img.mode == "L":
        img = img.convert("RGB")
    return img


def calcular_ajuste(img_w: int, img_h: int, box_w: float, box_h: float) -> Tuple[float, float]:
    ratio = min(box_w / img_w, box_h / img_h)
    return img_w * ratio, img_h * ratio


def wrap_text(pdf: canvas.Canvas, text: str, max_width: float, font_name="Helvetica", font_size=10):
    words = text.split()
    lines = []
    current = ""
    for w in words:
        test = (current + " " + w).strip()
        if stringWidth(test, font_name, font_size) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = w
    if current:
        lines.append(current)
    return lines


def obtener_logo_fuente(logo_file):
    if logo_file is not None:
        return logo_file
    if os.path.exists(DEFAULT_LOGO_PATH):
        return DEFAULT_LOGO_PATH
    return None


def preparar_logo_pdf(logo_file, max_width=mm_to_pt(24), max_height=mm_to_pt(10)):
    src = obtener_logo_fuente(logo_file)
    if not src:
        return None, 0, 0
    img = abrir_imagen(src)
    w, h = calcular_ajuste(img.width, img.height, max_width, max_height)
    return ImageReader(img), w, h


# -----------------------------
# Encabezado de descripción
# -----------------------------

def dibujar_descripcion(pdf: canvas.Canvas, cilindro: str, descripcion: str):
    x = OUTER_MARGIN
    y_top = PAGE_H - OUTER_MARGIN - mm_to_pt(2)

    etiqueta = f"CILINDRO {cilindro.strip() if cilindro else 'X'}:"

    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(x, y_top, etiqueta)

    # texto a la derecha de la etiqueta
    offset = stringWidth(etiqueta, "Helvetica-Bold", 11) + mm_to_pt(3)
    max_width = USABLE_W - offset

    pdf.setFont("Helvetica", 10)
    lines = wrap_text(pdf, descripcion or "", max_width)

    # limitar a 5 renglones
    lines = lines[:5]

    line_h = mm_to_pt(5)
    y = y_top
    for i, ln in enumerate(lines):
        pdf.drawString(x + offset, y - (i * line_h), ln)


# -----------------------------
# Footer (sin línea negra)
# -----------------------------

def dibujar_footer(pdf: canvas.Canvas, campo: str, logo_reader, logo_w: float, logo_h: float):
    text_y = OUTER_MARGIN + mm_to_pt(2)

    pdf.setFont("Helvetica", 9)
    pdf.drawString(OUTER_MARGIN, text_y, "Lubricantes Mobil")

    campo = (campo or "").strip()
    if campo:
        w = stringWidth(campo, "Helvetica", 9)
        pdf.drawString((PAGE_W - w) / 2, text_y, campo)

    if logo_reader:
        x_logo = PAGE_W - OUTER_MARGIN - logo_w
        y_logo = OUTER_MARGIN + mm_to_pt(1)
        pdf.drawImage(logo_reader, x_logo, y_logo, width=logo_w, height=logo_h, preserveAspectRatio=True, mask="auto")


# -----------------------------
# PDF
# -----------------------------

def generar_pdf(registros, campo: str, logo_file, cilindro: str, descripcion: str) -> bytes:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    logo_reader, logo_w, logo_h = preparar_logo_pdf(logo_file)

    total_paginas = math.ceil(len(registros) / MAX_PER_PAGE)

    for page_idx in range(total_paginas):
        start = page_idx * MAX_PER_PAGE
        end = start + MAX_PER_PAGE
        lote = registros[start:end]

        dibujar_descripcion(pdf, cilindro, descripcion)

        for i, item in enumerate(lote):
            row = i // COLUMNS
            col = i % COLUMNS

            x = OUTER_MARGIN + col * (BOX_W + GAP_X)
            y_top = PAGE_H - OUTER_MARGIN - DESC_H - TOP_EXTRA - row * (CELL_H + GAP_Y)
            y_cell = y_top - CELL_H

            img = abrir_imagen(item["file"])
            draw_w, draw_h = calcular_ajuste(img.width, img.height, BOX_W, IMAGE_BOX_H)
            x_img = x + (BOX_W - draw_w) / 2
            y_img = y_cell + (IMAGE_BOX_H - draw_h) / 2

            pdf.drawImage(ImageReader(img), x_img, y_img, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")

        dibujar_footer(pdf, campo, logo_reader, logo_w, logo_h)
        pdf.showPage()

    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# Word
# -----------------------------

def agregar_imagen_a_word(parrafo, img: Image.Image, max_w_in=3.05, max_h_in=1.85):
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    try:
        img.save(temp.name, format="PNG")
        run = parrafo.add_run()
        run.add_picture(temp.name, width=Inches(max_w_in), height=Inches(max_h_in))
    finally:
        temp.close()
        if os.path.exists(temp.name):
            os.unlink(temp.name)


def agregar_logo_a_word(parrafo, logo_file, width_in=0.9):
    src = obtener_logo_fuente(logo_file)
    if not src:
        return
    img = abrir_imagen(src)
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    try:
        img.save(temp.name, format="PNG")
        parrafo.add_run().add_picture(temp.name, width=Inches(width_in))
    finally:
        temp.close()
        if os.path.exists(temp.name):
            os.unlink(temp.name)


def generar_docx(registros, campo: str, logo_file, cilindro: str, descripcion: str) -> bytes:
    if not WORD_AVAILABLE:
        return b""

    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(0.45)
    sec.bottom_margin = Inches(0.5)
    sec.left_margin = Inches(0.45)
    sec.right_margin = Inches(0.45)

    total_paginas = math.ceil(len(registros) / MAX_PER_PAGE)

    for page_idx in range(total_paginas):
        # párrafo superior
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p.add_run(f"CILINDRO {cilindro.strip() if cilindro else 'X'}: ")
        r1.bold = True
        r2 = p.add_run(descripcion or "")

        start = page_idx * MAX_PER_PAGE
        end = start + MAX_PER_PAGE
        lote = registros[start:end]

        tabla = doc.add_table(rows=ROWS, cols=COLUMNS)
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabla.autofit = False

        for row in tabla.rows:
            row.height = Inches(2.45)
            for cell in row.cells:
                cell.width = Inches(3.75)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        for i, item in enumerate(lote):
            r = i // COLUMNS
            c = i % COLUMNS
            cell = tabla.cell(r, c)
            cell.text = ""
            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            agregar_imagen_a_word(p_img, abrir_imagen(item["file"]))

        doc.add_paragraph()
        footer = doc.add_table(rows=1, cols=3)
        footer.alignment = WD_TABLE_ALIGNMENT.CENTER
        footer.autofit = False

        c0, c1, c2 = footer.rows[0].cells

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p0.add_run("Lubricantes Mobil")

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.add_run((campo or "").strip())

        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        agregar_logo_a_word(p2, logo_file)

        if page_idx < total_paginas - 1:
            doc.add_section(WD_SECTION.NEW_PAGE)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# Estado
# -----------------------------

def inicializar_registros(files):
    ordenados = ordenar_archivos(files)
    regs = []
    for idx, f in enumerate(ordenados, start=1):
        regs.append({"uid": f"{idx}_{f.name}_{getattr(f, 'size', 0)}", "file": f})
    return regs


def mover_arriba(registros, idx):
    if idx > 0:
        registros[idx - 1], registros[idx] = registros[idx], registros[idx - 1]


def mover_abajo(registros, idx):
    if idx < len(registros) - 1:
        registros[idx + 1], registros[idx] = registros[idx], registros[idx + 1]


# -----------------------------
# UI
# -----------------------------

st.title("🖼️ Generador de anexos de videoscopía")

with st.container(border=True):
    st.markdown("#### Datos generales")
    campo = st.text_input("Campo", placeholder="Ej: Estación SANTS")
    cilindro = st.text_input("Cilindro", placeholder="Ej: 1L, 2R")
    descripcion = st.text_area("Descripción de hallazgos (máx 5 renglones)")
    logo_file = st.file_uploader("Logo opcional", type=["png", "jpg", "jpeg"])
    if logo_file is None and os.path.exists(DEFAULT_LOGO_PATH):
        st.caption("Se usará el logo predeterminado de Mobil.")

uploaded_files = st.file_uploader("Sube imágenes", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if uploaded_files:
    current_ids = [f"{i}_{f.name}_{getattr(f, 'size', 0)}" for i, f in enumerate(ordenar_archivos(uploaded_files), start=1)]

    if "registros_imagenes" not in st.session_state:
        st.session_state.registros_imagenes = inicializar_registros(uploaded_files)
    else:
        existing_ids = [r["uid"] for r in st.session_state.registros_imagenes]
        if existing_ids != current_ids:
            st.session_state.registros_imagenes = inicializar_registros(uploaded_files)

    registros = st.session_state.registros_imagenes

    with st.container(border=True):
        st.markdown("#### Reordenar imágenes")
        for idx, item in enumerate(registros):
            c1, c2 = st.columns([3, 1])
            with c1:
                st.image(abrir_imagen(item["file"]), use_container_width=True)
                st.caption(item["file"].name)
            with c2:
                if st.button("⬆️", key=f"up_{item['uid']}"):
                    mover_arriba(registros, idx)
                    st.rerun()
                if st.button("⬇️", key=f"down_{item['uid']}"):
                    mover_abajo(registros, idx)
                    st.rerun()

    pdf_bytes = generar_pdf(registros, campo, logo_file, cilindro, descripcion)

    if WORD_AVAILABLE:
        word_bytes = generar_docx(registros, campo, logo_file, cilindro, descripcion)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button("📄 Descargar PDF", pdf_bytes, "anexos.pdf")
        with col2:
            st.download_button("📝 Descargar Word", word_bytes, "anexos.docx")
    else:
        st.download_button("📄 Descargar PDF", pdf_bytes, "anexos.pdf")
        st.warning("Word no disponible. Instala python-docx.")
else:
    st.info("Sube imágenes para continuar.")

