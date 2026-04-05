import io
import math
from typing import List, Tuple

import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

# =========================================================
# App: Organizador de imágenes a PDF tamaño carta
# Uso previsto: anexos de videoscopía
#
# Características:
# - 2 columnas por página
# - Máximo 8 imágenes por página
# - Todas las imágenes con el mismo tamaño de contenedor
# - Reordenamiento manual de imágenes
# - Título por imagen (ej. cilindro inspeccionado)
# - Pie de página fijo:
#   izquierda: Lubricantes Mobil
#   centro: campo ingresado por el usuario
#   derecha: logo cargado en la app
# =========================================================

st.set_page_config(
    page_title="Anexos de videoscopía a PDF",
    page_icon="🖼️",
    layout="wide",
)

# -----------------------------
# Configuración PDF
# -----------------------------
PAGE_W, PAGE_H = letter
COLUMNS = 2
ROWS = 4
MAX_PER_PAGE = COLUMNS * ROWS


def mm_to_pt(mm: float) -> float:
    return mm * 72 / 25.4


OUTER_MARGIN = mm_to_pt(12)
GAP_X = mm_to_pt(6)
GAP_Y = mm_to_pt(8)
FOOTER_H = mm_to_pt(14)
TITLE_H = mm_to_pt(11)
TOP_EXTRA = mm_to_pt(6)

USABLE_W = PAGE_W - (2 * OUTER_MARGIN)
USABLE_H = PAGE_H - (2 * OUTER_MARGIN) - FOOTER_H - TOP_EXTRA
BOX_W = (USABLE_W - GAP_X) / COLUMNS
CELL_H = (USABLE_H - (GAP_Y * (ROWS - 1))) / ROWS
IMAGE_BOX_H = CELL_H - TITLE_H


# -----------------------------
# Utilidades
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



def preparar_logo(logo_file, max_width=mm_to_pt(24), max_height=mm_to_pt(10)):
    if not logo_file:
        return None, 0, 0
    img = Image.open(logo_file)
    img = ImageOps.exif_transpose(img)
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    elif img.mode == "L":
        img = img.convert("RGB")

    draw_w, draw_h = calcular_ajuste(img.width, img.height, max_width, max_height)
    return ImageReader(img), draw_w, draw_h



def dibujar_footer(pdf: canvas.Canvas, campo: str, logo_reader, logo_w: float, logo_h: float):
    line_y = OUTER_MARGIN + FOOTER_H - mm_to_pt(2)
    pdf.setLineWidth(0.5)
    pdf.line(OUTER_MARGIN, line_y, PAGE_W - OUTER_MARGIN, line_y)

    text_y = OUTER_MARGIN + mm_to_pt(2)

    pdf.setFont("Helvetica", 9)
    pdf.drawString(OUTER_MARGIN, text_y, "Lubricantes Mobil")

    campo = campo.strip() if campo else ""
    if campo:
        pdf.setFont("Helvetica", 9)
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
            mask='auto',
        )



def generar_pdf(registros, campo: str, logo_file) -> bytes:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    logo_reader, logo_w, logo_h = preparar_logo(logo_file)

    total_paginas = math.ceil(len(registros) / MAX_PER_PAGE)

    for page_idx in range(total_paginas):
        start = page_idx * MAX_PER_PAGE
        end = start + MAX_PER_PAGE
        lote = registros[start:end]

        for i, item in enumerate(lote):
            row = i // COLUMNS
            col = i % COLUMNS

            x = OUTER_MARGIN + col * (BOX_W + GAP_X)
            y_top = PAGE_H - OUTER_MARGIN - TOP_EXTRA - row * (CELL_H + GAP_Y)
            y_cell = y_top - CELL_H

            titulo = item["titulo"].strip() if item["titulo"].strip() else item["file"].name
            titulo = truncar_texto(pdf, titulo, BOX_W, font_name="Helvetica-Bold", font_size=10)

            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawString(x, y_top - mm_to_pt(4), titulo)

            image_y = y_cell
            img = abrir_imagen(item["file"])
            draw_w, draw_h = calcular_ajuste(img.width, img.height, BOX_W, IMAGE_BOX_H)
            x_img = x + (BOX_W - draw_w) / 2
            y_img = image_y + (IMAGE_BOX_H - draw_h) / 2

            pdf.drawImage(
                ImageReader(img),
                x_img,
                y_img,
                width=draw_w,
                height=draw_h,
                preserveAspectRatio=True,
                mask='auto',
            )

        dibujar_footer(pdf, campo, logo_reader, logo_w, logo_h)
        pdf.showPage()

    pdf.save()
    buffer.seek(0)
    return buffer.getvalue()



def nombre_salida() -> str:
    return "anexos_videoscopia.pdf"


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
            "titulo": base,
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
    "Sube tus imágenes y genera un PDF tamaño carta con 2 columnas, máximo 8 imágenes por página, "
    "todas del mismo tamaño y con pie de página personalizado."
)

with st.container(border=True):
    st.markdown("#### Datos generales del anexo")
    campo = st.text_input("Campo donde se realizó la videoscopía", placeholder="Ejemplo: Campo Tibú, Estación SANTS, GRB")
    logo_file = st.file_uploader("Logo para el pie de página", type=["png", "jpg", "jpeg"], key="logo_uploader")

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
        st.markdown("#### Reordenar imágenes y editar títulos")
        st.caption(
            "Puedes cambiar el orden manualmente y escribir el título de cada imagen, por ejemplo: Cilindro 1L, "
            "Cilindro 2R, Culata, Válvula de escape, etc."
        )

        for idx, item in enumerate(registros):
            col1, col2, col3 = st.columns([1.3, 2.8, 1.2])

            with col1:
                img = abrir_imagen(item["file"])
                st.image(img, use_container_width=True)

            with col2:
                nuevo_titulo = st.text_input(
                    f"Título {idx + 1}",
                    value=item["titulo"],
                    key=f"titulo_{item['uid']}"
                )
                item["titulo"] = nuevo_titulo
                st.caption(f"Archivo: {item['file'].name}")

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
    st.info(f"El PDF tendrá {total_paginas} página(s). Cada página admite hasta {MAX_PER_PAGE} imágenes.")

    with st.expander("Ver orden final"):
        for idx, item in enumerate(registros, start=1):
            st.write(f"{idx}. {item['titulo']}")

    pdf_bytes = generar_pdf(registros, campo, logo_file)

    st.download_button(
        label="📄 Descargar PDF de anexos",
        data=pdf_bytes,
        file_name=nombre_salida(),
        mime="application/pdf",
        use_container_width=True,
    )
else:
    st.warning("Sube al menos una imagen para generar el PDF.")

with st.container(border=True):
    st.markdown("#### Formato del documento")
    st.markdown(
        f"""
- Tamaño carta.
- 2 columnas por página.
- Máximo {MAX_PER_PAGE} imágenes por página.
- Todas las imágenes conservan el mismo tamaño de presentación.
- Cada imagen lleva su título encima.
- Pie de página con texto fijo a la izquierda, campo al centro y logo a la derecha.
        """
    )

with st.container(border=True):
    st.markdown("#### requirements.txt")
    st.code(
        """streamlit
pillow
reportlab""",
        language="text",
    )

    st.markdown("#### Ejecución local")
    st.code("streamlit run app_organizador_imagenes_pdf.py", language="bash")
