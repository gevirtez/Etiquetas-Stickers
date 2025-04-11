import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from tkinter import filedialog
from tkinter import messagebox
import os
import sys

# Tamaño de etiqueta
LABEL_WIDTH = 55 * mm
LABEL_HEIGHT = 25 * mm
MARGIN = 2 * mm
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
excel_path = filedialog.askopenfilename(
    title="Selecciona el archivo Excel",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not excel_path:
    messagebox.showinfo("Cancelado", "No se seleccionó ningún archivo. El programa se cerrará.")
    sys.exit()

df = pd.read_excel(excel_path, skiprows=5)

codigos = df.iloc[:, 3].dropna().astype(str).tolist()
nombres = df.iloc[:, 5].dropna().astype(str).tolist()

# Crear PDF
pdf_path = os.path.join(os.path.dirname(excel_path), "etiquetas_nombre_muestra.pdf")
c = canvas.Canvas(pdf_path, pagesize=(LABEL_WIDTH, LABEL_HEIGHT))

# Mostrar mensaje de éxito
messagebox.showinfo("Éxito", f"PDF de etiquetas creado con éxito:\n\n{pdf_path}")

for codigo, nombre in zip(codigos, nombres):
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.rect(MARGIN, MARGIN, INNER_WIDTH, INNER_HEIGHT)

    # Dibujar código
    c.setFont("Helvetica-Bold", 6.5)
    codigo_y = LABEL_HEIGHT - MARGIN - 3.5 * mm
    c.drawString(MARGIN + 1 * mm, codigo_y, f"Código: {codigo}")

    # Espacio para el nombre
    nombre_top_y = codigo_y - 1.5 * mm
    nombre_bottom_y = MARGIN + 1.5 * mm
    available_height = nombre_top_y - nombre_bottom_y

    # Encontrar el tamaño de fuente más grande posible
    best_font_size = None
    best_lines = []
    max_font_size = 7.5
    min_font_size = 2.5
    step = 0.1

    font_size = max_font_size
    while font_size >= min_font_size:
        line_height = font_size * 1.2  # más realista
        lines = wrap_text_by_width(nombre, INNER_WIDTH - 2 * mm, "Helvetica", font_size, c)
        total_height = len(lines) * line_height

        if total_height <= available_height:
            best_font_size = font_size
            best_lines = lines
            break

        font_size -= step

    if not best_lines:
        best_font_size = min_font_size
        best_lines = wrap_text_by_width(nombre, INNER_WIDTH - 2 * mm, "Helvetica", best_font_size, c)

    # Dibujar texto desde la parte superior disponible hacia abajo
    c.setFont("Helvetica", best_font_size)
    line_height = best_font_size * 1.2
    start_y = nombre_top_y - line_height + 1.2  # ajustado para pegar más arriba

    text = c.beginText()
    text.setTextOrigin(MARGIN + 1.5 * mm, start_y)

    for line in best_lines:
        if start_y < nombre_bottom_y:
            break
        text.textLine(line)
        start_y -= line_height

    c.drawText(text)
    c.showPage()

c.save()
print(f"✅ PDF generado sin desbordes: {pdf_path}")
