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
# - 2 columnas
# - máximo 8 imágenes por página
# - mismo tamaño visual para todas las imágenes
# - reordenamiento manual
# - nombre por imagen
# - salida en PDF y Word
# - pie de página con logo Mobil, campo y texto fijo
# =========================================================

st.set_page_config(
    page_title="Anexos de videoscopía",
    page_icon="🖼️",
    layout="wide",
)

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
HEADER_H = mm_to_pt(20)
TOP_EXTRA = mm_to_pt(6)

USABLE_W = PAGE_W - (2 * OUTER_MARGIN)
USABLE_H = PAGE_H - (2 * OUTER_MARGIN) - FOOTER_H - HEADER_H - TOP_EXTRA
BOX_W = (USABLE_W - GAP_X) / COLUMNS
CELL_H = (USABLE_H - (GAP_Y * (ROWS - 1))) / ROWS
IMAGE_BOX_H = CELL_H


# -----------------------------
# Utilidades generales
# -----------------------------
def natural_key(name: str):
    parts = []
    current = ""
    is_digit = None
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
    draw_w = img_w * ratio
    draw_h = img_h * ratio
    return draw_w, draw_h



def truncar_texto(pdf: canvas.Canvas, text: str, max_width: float, font_name="Helvetica-Bold", font_size=10) -> str:
    if stringWidth(text, font_name, font_size) <= max_width:
        return text
    suffix = "..."
    out = text
    while out and stringWidth(out + suffix, font_name, font_size) > max_width:
        out = out[:-1]
    return (out + suffix) if out else suffix


def construir_titulo_documento(cilindro: str) -> str:
    base = "ANEXOS FOTOGRÁFICO VIDEOSCOPIA CILINDRO"
    cilindro = cilindro.strip()
    if cilindro:
        return f"{base} {cilindro}"
    return f"{base} X"



def obtener_logo_fuente(logo_file):
    if logo_file is not None:
        return logo_file
    if os.path.exists(DEFAULT_LOGO_PATH):
        return DEFAULT_LOGO_PATH
    return None



def preparar_logo_pdf(logo_file, max_width=mm_to_pt(24), max_height=mm_to_pt(10)):
    img_source = obtener_logo_fuente(logo_file)
    if not img_source:
        return None, 0, 0

    img = abrir_imagen(img_source)
    draw_w, draw_h = calcular_ajuste(img.width, img.height, max_width, max_height)
    return ImageReader(img), draw_w, draw_h


# -----------------------------
# PDF
# -----------------------------
def dibujar_header(pdf: canvas.Canvas, cilindro: str):
    titulo = construir_titulo_documento(cilindro)
    titulo = truncar_texto(pdf, titulo, USABLE_W, font_name="Helvetica-Bold", font_size=13)
    y = PAGE_H - OUTER_MARGIN - mm_to_pt(4)
    pdf.setFont("Helvetica-Bold", 13)
    pdf.drawCentredString(PAGE_W / 2, y, titulo)
    line_y = y - mm_to_pt(5)
    pdf.setLineWidth(0.8)
    pdf.line(OUTER_MARGIN, line_y, PAGE_W - OUTER_MARGIN, line_y)


def dibujar_footer(pdf: canvas.Canvas, campo: str, logo_reader, logo_w: float, logo_h: float):
    line_y = OUTER_MARGIN + FOOTER_H - mm_to_pt(2)
    pdf.setLineWidth(0.5)
    pdf.line(OUTER_MARGIN, line_y, PAGE_W - OUTER_MARGIN, line_y)

    text_y = OUTER_MARGIN + mm_to_pt(2)

    pdf.setFont("Helvetica", 9)
    pdf.drawString(OUTER_MARGIN, text_y, "Lubricantes Mobil")

    campo = campo.strip() if campo else ""
    if campo:
        campo_width = stringWidth(campo, "Helvetica", 9)
        pdf.drawString((PAGE_W - campo_width) / 2, text_y, campo)

    if logo_reader:
        x_logo = PAGE_W - OUTER_MARGIN - logo_w
        y_logo = OUTER_MARGIN + mm_to_pt(1)
        pdf.drawImage(
            logo_reader,
            x_logo,
            y_logo,
            width=logo_w,
            height=logo_h,
            preserveAspectRatio=True,
            mask="auto",
        )



def generar_pdf(registros, campo: str, logo_file, cilindro: str) -> bytes:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    logo_reader, logo_w, logo_h = preparar_logo_pdf(logo_file)

    total_paginas = math.ceil(len(registros) / MAX_PER_PAGE)

    for page_idx in range(total_paginas):
        start = page_idx * MAX_PER_PAGE
        end = start + MAX_PER_PAGE
        lote = registros[start:end]

        dibujar_header(pdf, cilindro)

        for i, item in enumerate(lote):
            row = i // COLUMNS
            col = i % COLUMNS

            x = OUTER_MARGIN + col * (BOX_W + GAP_X)
            y_top = PAGE_H - OUTER_MARGIN - HEADER_H - TOP_EXTRA - row * (CELL_H + GAP_Y)
            y_cell = y_top - CELL_H

            img = abrir_imagen(item["file"])
            draw_w, draw_h = calcular_ajuste(img.width, img.height, BOX_W, IMAGE_BOX_H)
            x_img = x + (BOX_W - draw_w) / 2
            y_img = y_cell + (IMAGE_BOX_H - draw_h) / 2

            pdf.drawImage(
                ImageReader(img),
                x_img,
                y_img,
                width=draw_w,
                height=draw_h,
                preserveAspectRatio=True,
                mask="auto",
            )

        dibujar_footer(pdf, campo, logo_reader, logo_w, logo_h)
        pdf.showPage()

    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# Word
# -----------------------------
def agregar_imagen_a_word(parrafo, img: Image.Image, max_width_in=3.05, max_height_in=1.85):
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    try:
        img.save(temp_file.name, format="PNG")

        width_px, height_px = img.size
        ratio = min((max_width_in * 96) / width_px, (max_height_in * 96) / height_px)
        final_width = max_width_in if ratio == (max_width_in * 96) / width_px else width_px * ratio / 96
        final_height = max_height_in if ratio == (max_height_in * 96) / height_px else height_px * ratio / 96

        run = parrafo.add_run()
        run.add_picture(temp_file.name, width=Inches(final_width), height=Inches(final_height))
    finally:
        temp_file.close()
        if os.path.exists(temp_file.name):
            os.unlink(temp_file.name)



def agregar_logo_a_word(parrafo, logo_file, width_in=0.9):
    logo_source = obtener_logo_fuente(logo_file)
    if not logo_source:
        return

    img_logo = abrir_imagen(logo_source)
    temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    try:
        img_logo.save(temp_logo.name, format="PNG")
        parrafo.add_run().add_picture(temp_logo.name, width=Inches(width_in))
    finally:
        temp_logo.close()
        if os.path.exists(temp_logo.name):
            os.unlink(temp_logo.name)



def generar_docx(registros, campo: str, logo_file, cilindro: str) -> bytes:
    if not WORD_AVAILABLE:
        return b""

    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(9)

    total_paginas = math.ceil(len(registros) / MAX_PER_PAGE)
    titulo_doc = construir_titulo_documento(cilindro)

    for page_idx in range(total_paginas):
        p_header = doc.add_paragraph()
        p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_header.space_after = Pt(8)
        r_header = p_header.add_run(titulo_doc)
        r_header.bold = True
        r_header.font.size = Pt(13)

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
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for i, item in enumerate(lote):
            row_idx = i // COLUMNS
            col_idx = i % COLUMNS
            cell = tabla.cell(row_idx, col_idx)
            cell.text = ""

            p_img = cell.paragraphs[0]
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            agregar_imagen_a_word(p_img, abrir_imagen(item["file"]))

        doc.add_paragraph()
        footer = doc.add_table(rows=1, cols=3)
        footer.alignment = WD_TABLE_ALIGNMENT.CENTER
        footer.autofit = False

        footer.rows[0].cells[0].width = Inches(2.3)
        footer.rows[0].cells[1].width = Inches(3.2)
        footer.rows[0].cells[2].width = Inches(1.5)

        c0, c1, c2 = footer.rows[0].cells

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r0 = p0.add_run("Lubricantes Mobil")
        r0.font.size = Pt(9)

        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(campo.strip() if campo else "")
        r1.font.size = Pt(9)

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
# Nombres de salida
# -----------------------------
def nombre_salida_pdf() -> str:
    return "anexos_videoscopia.pdf"



def nombre_salida_word() -> str:
    return "anexos_videoscopia.docx"


# -----------------------------
# Estado y orden manual
# -----------------------------
def inicializar_registros(files):
    ordenados = ordenar_archivos(files)
    registros = []
    for idx, f in enumerate(ordenados, start=1):
        base = f.name.rsplit('.', 1)[0]
        registros.append({
            "uid": f"{idx}_{f.name}_{getattr(f, 'size', 0)}",
            "file": f,
            "nombre_base": base,
        })
    return registros



def mover_arriba(registros, idx):
    if idx > 0:
        registros[idx - 1], registros[idx] = registros[idx], registros[idx - 1]



def mover_abajo(registros, idx):
    if idx < len(registros) - 1:
        registros[idx + 1], registros[idx] = registros[idx], registros[idx + 1]


# -----------------------------
# Interfaz
# -----------------------------
st.title("🖼️ Generador de anexos de videoscopía")
st.write(
    "Sube tus imágenes y genera anexos en tamaño carta con 2 columnas, máximo 8 imágenes por página, "
    "mismo tamaño visual, nombre por imagen y salida en PDF o Word."
)

with st.container(border=True):
    st.markdown("#### Datos generales del anexo")
    campo = st.text_input(
        "Campo donde se realizó la videoscopía",
        placeholder="Ejemplo: Campo Tibú, Estación SANTS, GRB"
    )
    cilindro = st.text_input(
        "Cilindro",
        placeholder="Ejemplo: 1L, 2R, 3, A"
    )
    logo_file = st.file_uploader(
        "Logo para el pie de página, opcional",
        type=["png", "jpg", "jpeg"],
        key="logo_uploader"
    )
    if logo_file is None and os.path.exists(DEFAULT_LOGO_PATH):
        st.caption("Se usará automáticamente el logo predeterminado de Mobil en el pie de página.")

uploaded_files = st.file_uploader(
    "Sube imágenes JPG, JPEG o PNG",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="images_uploader",
)

if uploaded_files:
    current_ids = [f"{i}_{f.name}_{getattr(f, 'size', 0)}" for i, f in enumerate(ordenar_archivos(uploaded_files), start=1)]

    if "registros_imagenes" not in st.session_state:
        st.session_state.registros_imagenes = inicializar_registros(uploaded_files)
    else:
        existing_ids = [r["uid"] for r in st.session_state.registros_imagenes]
        if existing_ids != current_ids:
            st.session_state.registros_imagenes = inicializar_registros(uploaded_files)

    registros = st.session_state.registros_imagenes

    st.success(f"Se cargaron {len(registros)} imagen(es).")

    with st.container(border=True):
        st.markdown("#### Reordenar imágenes")
        st.caption("Puedes cambiar el orden manualmente. Las imágenes no llevarán nombre ni título individual.")

        for idx, item in enumerate(registros):
            col1, col2, col3 = st.columns([1.3, 2.8, 1.2])

            with col1:
                img = abrir_imagen(item["file"])
                st.image(img, use_container_width=True)

            with col2:
                st.markdown(f"**Imagen {idx + 1}**")
                st.caption(f"Archivo original: {item['file'].name}")

            with col3:
                st.write("")
                if st.button("⬆️ Subir", key=f"up_{item['uid']}", use_container_width=True):
                    mover_arriba(registros, idx)
                    st.rerun()
                if st.button("⬇️ Bajar", key=f"down_{item['uid']}", use_container_width=True):
                    mover_abajo(registros, idx)
                    st.rerun()

            st.divider()

    total_paginas = math.ceil(len(registros) / MAX_PER_PAGE)
    st.info(f"El documento tendrá {total_paginas} página(s). Cada página admite hasta {MAX_PER_PAGE} imágenes.")

    with st.expander("Ver orden final"):
        for idx, item in enumerate(registros, start=1):
            st.write(f"{idx}. {item['file'].name}")

    pdf_bytes = generar_pdf(registros, campo, logo_file, cilindro)

    if WORD_AVAILABLE:
        word_bytes = generar_docx(registros, campo, logo_file, cilindro)
        col_pdf, col_word = st.columns(2)
        with col_pdf:
            st.download_button(
                label="📄 Descargar PDF de anexos",
                data=pdf_bytes,
                file_name=nombre_salida_pdf(),
                mime="application/pdf",
                use_container_width=True,
            )
        with col_word:
            st.download_button(
                label="📝 Descargar Word de anexos",
                data=word_bytes,
                file_name=nombre_salida_word(),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
    else:
        st.download_button(
            label="📄 Descargar PDF de anexos",
            data=pdf_bytes,
            file_name=nombre_salida_pdf(),
            mime="application/pdf",
            use_container_width=True,
        )
        st.warning("La exportación a Word no está disponible porque falta instalar python-docx en el entorno.")
else:
    st.warning("Sube al menos una imagen para generar el documento.")

with st.container(border=True):
    st.markdown("#### Formato del documento")
    st.markdown(
        f"""
- Tamaño carta.
- 2 columnas por página.
- Máximo {MAX_PER_PAGE} imágenes por página.
- Todas las imágenes conservan el mismo tamaño de presentación.
- Las imágenes no llevan nombre ni título individual.
- Cada hoja lleva en la parte superior el título: ANEXOS FOTOGRÁFICO VIDEOSCOPIA CILINDRO X.
- El valor de X lo ingresa el usuario.
- Si no subes logo, la app usa el logo predeterminado de Mobil.
- Pie de página con texto fijo a la izquierda, campo al centro y logo a la derecha.
- Descarga disponible en PDF y también en Word si python-docx está instalado.
        """
    )

with st.container(border=True):
    st.markdown("#### requirements.txt")
    st.code(
        """streamlit
pillow
reportlab
python-docx""",
        language="text",
    )

    st.markdown("#### Ejecución local")
    st.code("streamlit run app_organizador_imagenes_pdf.py", language="bash")
