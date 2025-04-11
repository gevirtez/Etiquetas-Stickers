import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from tkinter import filedialog
import os

# Tamaño etiqueta
LABEL_WIDTH = 55 * mm
LABEL_HEIGHT = 25 * mm
MARGIN = 1.0 * mm  # Reducido para maximizar espacio
INNER_WIDTH = LABEL_WIDTH - 2 * MARGIN
INNER_HEIGHT = LABEL_HEIGHT - 2 * MARGIN

# Cortar texto por ancho real medido
def wrap_text_by_width(text, max_width, font_name, font_size, canvas_obj):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        test_line = f"{current_line} {word}".strip()
        text_width = canvas_obj.stringWidth(test_line, font_name, font_size)
        if text_width <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    return lines

# Selección del archivo
excel_path = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Excel Files", "*.xlsx *.xls")])
df = pd.read_excel(excel_path, skiprows=5)
codigos = df.iloc[:, 3].dropna().astype(str).tolist()
nombres = df.iloc[:, 5].dropna().astype(str).tolist()

# Crear PDF
pdf_path = os.path.join(os.path.dirname(excel_path), "etiquetas_nombre_muestra.pdf")
c = canvas.Canvas(pdf_path, pagesize=(LABEL_WIDTH, LABEL_HEIGHT))

for codigo, nombre in zip(codigos, nombres):
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.rect(MARGIN, MARGIN, INNER_WIDTH, INNER_HEIGHT)

    # Dibujar código (más cerca del borde superior)
    c.setFont("Helvetica-Bold", 6.5)
    codigo_y = LABEL_HEIGHT - MARGIN - 1.8 * mm
    c.drawString(MARGIN + 1 * mm, codigo_y, f"Código: {codigo}")

    # Espacio disponible para el nombre de la muestra
    nombre_top_y = codigo_y - 1.2 * mm
    nombre_bottom_y = MARGIN + 1.2 * mm
    available_height = nombre_top_y - nombre_bottom_y
    available_width = INNER_WIDTH - 2 * mm

    # Encontrar el tamaño de fuente más grande posible que encaje en ambos ejes
    best_font_size = None
    best_lines = []
    max_font_size = 8.0
    min_font_size = 2.0
    step = 0.1

    font_size = max_font_size
    while font_size >= min_font_size:
        line_height = font_size * 1.15
        lines = wrap_text_by_width(nombre, available_width, "Helvetica", font_size, c)
        total_height = len(lines) * line_height

        # Comprobar si cabe tanto en alto como en ancho
        if total_height <= available_height and all(c.stringWidth(line, "Helvetica", font_size) <= available_width for line in lines):
            best_font_size = font_size
            best_lines = lines
            break

        font_size -= step

    if not best_lines:
        best_font_size = min_font_size
        best_lines = wrap_text_by_width(nombre, available_width, "Helvetica", best_font_size, c)

    # Dibujar el nombre centrado verticalmente en el espacio disponible
    c.setFont("Helvetica", best_font_size)
    line_height = best_font_size * 1.15
    total_text_height = len(best_lines) * line_height
    start_y = nombre_top_y - ((available_height - total_text_height) / 2)

    text = c.beginText()
    text.setTextOrigin(MARGIN + 1.5 * mm, start_y)

    for line in best_lines:
        text.textLine(line)

    c.drawText(text)
    c.showPage()

c.save()
print(f"✅ PDF generado con nombres lo más grandes posible: {pdf_path}")
