import sys
sys.path.insert(0, '/Users/javiermartinamador/Documents/Agencia Marketing')

from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os

BASE_DIR = '/Users/javiermartinamador/Documents/Agencia Marketing'
BG_DIR = f'{BASE_DIR}/assets/pdf_backgrounds_salesianas'

# Slide dimensions from PDF (1440x811 at 2x = landscape 16:9-ish)
# Actually PDF is A4 landscape: 29.7cm x 21cm
SLIDE_WIDTH = Cm(29.7)
SLIDE_HEIGHT = Cm(21.0)

prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT

blank_layout = prs.slide_layouts[6]

def add_text_box(slide, left, top, width, height, text, font_size=12, bold=False, color=RGBColor(0x00,0x00,0x00), align=PP_ALIGN.LEFT, font_name='Calibri'):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = align
    return txBox

# Colors from the actual design
GREEN = RGBColor(0x45, 0xC2, 0x41)  # Lynks green
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)

# ============================================
# SLIDE 1: Cover
# ============================================
slide1 = prs.slides.add_slide(blank_layout)
slide1.shapes.add_picture(f'{BG_DIR}/salesianas_slide_1.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# Editable fields on slide 1:
# - Client name in black box: "HIJAS DE MARÍA AUXILIADORA"
add_text_box(slide1, Cm(1.8), Cm(9.0), Cm(18), Cm(1.2), "HIJAS DE MARÍA AUXILIADORA", 20, True, WHITE, PP_ALIGN.LEFT)
# - Client suffix: "INSTITUTO SALESIANAS CIR"
add_text_box(slide1, Cm(1.8), Cm(10.2), Cm(18), Cm(0.6), "INSTITUTO SALESIANAS CIR", 11, False, GREEN, PP_ALIGN.LEFT)
# - NIF: "NIF R0800057B"
add_text_box(slide1, Cm(1.8), Cm(10.8), Cm(10), Cm(0.6), "NIF R0800057B", 11, False, GREEN, PP_ALIGN.LEFT)
# - Feature bullets
add_text_box(slide1, Cm(1.8), Cm(12.5), Cm(20), Cm(0.8), "✓ 730 líneas móviles  ·  B2 Pack Ilimitada", 11, False, DARK_GRAY, PP_ALIGN.LEFT)
add_text_box(slide1, Cm(1.8), Cm(13.3), Cm(20), Cm(0.8), "✓ Cobertura nacional MasOrange + Movistar", 11, False, DARK_GRAY, PP_ALIGN.LEFT)
add_text_box(slide1, Cm(1.8), Cm(14.1), Cm(20), Cm(0.8), "✓ Bolsa de terminales: 8.000€ (36 meses)", 11, False, DARK_GRAY, PP_ALIGN.LEFT)
# - Date
add_text_box(slide1, Cm(24), Cm(19), Cm(4), Cm(0.6), "Junio 2026", 10, False, DARK_GRAY, PP_ALIGN.RIGHT)

# ============================================
# SLIDE 2: Why Lynks TIC
# ============================================
slide2 = prs.slides.add_slide(blank_layout)
slide2.shapes.add_picture(f'{BG_DIR}/salesianas_slide_2.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# Footer commercial info
add_text_box(slide2, Cm(1), Cm(19.5), Cm(27), Cm(0.4), "lynks-tic.com · 911 08 98 77 · info@lynks-tic.com · Comercial: Javier Martín · 664 25 89 77", 8, False, DARK_GRAY, PP_ALIGN.CENTER)

# ============================================
# SLIDE 3: B2Mobile Detail
# ============================================
slide3 = prs.slides.add_slide(blank_layout)
slide3.shapes.add_picture(f'{BG_DIR}/salesianas_slide_3.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# Big number
add_text_box(slide3, Cm(1.5), Cm(5.5), Cm(8), Cm(2), "730", 72, True, GREEN, PP_ALIGN.CENTER)
add_text_box(slide3, Cm(1.5), Cm(7.8), Cm(8), Cm(0.6), "líneas móviles ilimitadas", 12, False, DARK_GRAY, PP_ALIGN.CENTER)
# Price
add_text_box(slide3, Cm(1), Cm(9.5), Cm(27), Cm(0.8), "Precio negociado para 730 líneas: 7,50€/línea/mes   (Catálogo público: 9,45€  ·  21% de descuento)", 11, True, DARK_GRAY, PP_ALIGN.CENTER)

add_text_box(slide3, Cm(1), Cm(19.5), Cm(27), Cm(0.4), "lynks-tic.com · 911 08 98 77 · info@lynks-tic.com · Comercial: Javier Martín · 664 25 89 77", 8, False, DARK_GRAY, PP_ALIGN.CENTER)

# ============================================
# SLIDE 4: What's included
# ============================================
slide4 = prs.slides.add_slide(blank_layout)
slide4.shapes.add_picture(f'{BG_DIR}/salesianas_slide_4.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# Total price in green
add_text_box(slide4, Cm(1), Cm(5.2), Cm(8), Cm(1.2), "5.475€/mes", 32, True, GREEN, PP_ALIGN.CENTER)
# Before price
add_text_box(slide4, Cm(1), Cm(6.5), Cm(8), Cm(0.5), "antes: 6.898,50€/mes", 10, False, DARK_GRAY, PP_ALIGN.CENTER)
# Discount
add_text_box(slide4, Cm(1), Cm(7.2), Cm(8), Cm(0.5), "21% de descuento", 10, True, GREEN, PP_ALIGN.CENTER)

# Card details
add_text_box(slide4, Cm(10.5), Cm(7.0), Cm(7), Cm(0.5), "730 líneas móviles", 10, False, DARK_GRAY, PP_ALIGN.LEFT)
add_text_box(slide4, Cm(10.5), Cm(7.6), Cm(7), Cm(0.5), "Voz y datos ilimitados", 10, False, DARK_GRAY, PP_ALIGN.LEFT)
add_text_box(slide4, Cm(10.5), Cm(8.2), Cm(7), Cm(0.5), "7,50€/línea/mes", 10, False, DARK_GRAY, PP_ALIGN.LEFT)

# Coverage card
add_text_box(slide4, Cm(18.5), Cm(5.8), Cm(7), Cm(0.8), "Nacional", 20, True, GREEN, PP_ALIGN.CENTER)
add_text_box(slide4, Cm(18.5), Cm(6.8), Cm(7), Cm(0.5), "cobertura garantizada", 10, False, DARK_GRAY, PP_ALIGN.CENTER)
add_text_box(slide4, Cm(18.5), Cm(7.6), Cm(7), Cm(0.5), "Doble red", 10, True, GREEN, PP_ALIGN.CENTER)

# Terminal bag
add_text_box(slide4, Cm(1), Cm(14.0), Cm(8), Cm(1.2), "8.000€", 32, True, GREEN, PP_ALIGN.CENTER)

add_text_box(slide4, Cm(1), Cm(19.5), Cm(27), Cm(0.4), "lynks-tic.com · 911 08 98 77 · info@lynks-tic.com · Comercial: Javier Martín · 664 25 89 77", 8, False, DARK_GRAY, PP_ALIGN.CENTER)

# ============================================
# SLIDE 5: Economic detail
# ============================================
slide5 = prs.slides.add_slide(blank_layout)
slide5.shapes.add_picture(f'{BG_DIR}/salesianas_slide_5.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# Total
add_text_box(slide5, Cm(1.5), Cm(4.5), Cm(10), Cm(1.5), "5.475 €/mes", 36, True, GREEN, PP_ALIGN.LEFT)
add_text_box(slide5, Cm(1.5), Cm(6.2), Cm(10), Cm(0.5), "730 líneas · 7,50€/línea", 11, False, DARK_GRAY, PP_ALIGN.LEFT)

# Client info
add_text_box(slide5, Cm(1.5), Cm(7.8), Cm(12), Cm(0.5), "HIJAS DE MARÍA AUXILIADORA · INSTITUTO SALESIANAS CIR", 9, False, DARK_GRAY, PP_ALIGN.LEFT)
add_text_box(slide5, Cm(1.5), Cm(8.4), Cm(8), Cm(0.4), "R0800057B", 9, False, DARK_GRAY, PP_ALIGN.LEFT)
add_text_box(slide5, Cm(1.5), Cm(9.0), Cm(8), Cm(0.4), "36 meses", 9, False, DARK_GRAY, PP_ALIGN.LEFT)
add_text_box(slide5, Cm(1.5), Cm(9.6), Cm(8), Cm(0.4), "Sin cuota de alta", 9, False, DARK_GRAY, PP_ALIGN.LEFT)

# Table totals
add_text_box(slide5, Cm(19), Cm(11.5), Cm(5), Cm(0.5), "5.475,00€", 10, True, DARK_GRAY, PP_ALIGN.RIGHT)
add_text_box(slide5, Cm(24), Cm(11.5), Cm(4), Cm(0.5), "0,00€", 10, True, DARK_GRAY, PP_ALIGN.RIGHT)
add_text_box(slide5, Cm(24), Cm(12.2), Cm(4), Cm(0.5), "8.000,00€", 10, True, DARK_GRAY, PP_ALIGN.RIGHT)

# Total row
add_text_box(slide5, Cm(19), Cm(13.0), Cm(4), Cm(0.6), "730 uds.", 10, True, WHITE, PP_ALIGN.RIGHT)
add_text_box(slide5, Cm(24), Cm(13.0), Cm(4), Cm(0.6), "5.475,00€/mes", 10, True, WHITE, PP_ALIGN.RIGHT)

add_text_box(slide5, Cm(1), Cm(19.5), Cm(27), Cm(0.4), "lynks-tic.com · 911 08 98 77 · info@lynks-tic.com · Comercial: Javier Martín · 664 25 89 77", 8, False, DARK_GRAY, PP_ALIGN.CENTER)

# ============================================
# SLIDE 6: Connectivity
# ============================================
slide6 = prs.slides.add_slide(blank_layout)
slide6.shapes.add_picture(f'{BG_DIR}/salesianas_slide_6.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# FTTO prices
add_text_box(slide6, Cm(1.5), Cm(11.0), Cm(8), Cm(1.0), "69€/mes", 24, True, GREEN, PP_ALIGN.CENTER)
add_text_box(slide6, Cm(10.5), Cm(11.0), Cm(8), Cm(1.0), "49€/mes", 24, True, GREEN, PP_ALIGN.CENTER)
add_text_box(slide6, Cm(19.5), Cm(11.0), Cm(8), Cm(1.0), "34€/mes", 24, True, GREEN, PP_ALIGN.CENTER)

add_text_box(slide6, Cm(1), Cm(19.5), Cm(27), Cm(0.4), "lynks-tic.com · 911 08 98 77 · info@lynks-tic.com · Comercial: Javier Martín · 664 25 89 77", 8, False, DARK_GRAY, PP_ALIGN.CENTER)

# ============================================
# SLIDE 7: Cybersecurity
# ============================================
slide7 = prs.slides.add_slide(blank_layout)
slide7.shapes.add_picture(f'{BG_DIR}/salesianas_slide_7.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# Prices
add_text_box(slide7, Cm(14.5), Cm(7.5), Cm(8), Cm(1.0), "70€/mes", 28, True, GREEN, PP_ALIGN.CENTER)
add_text_box(slide7, Cm(14.5), Cm(11.5), Cm(8), Cm(1.0), "35€/mes", 28, True, GREEN, PP_ALIGN.CENTER)

add_text_box(slide7, Cm(1), Cm(19.5), Cm(27), Cm(0.4), "lynks-tic.com · 911 08 98 77 · info@lynks-tic.com · Comercial: Javier Martín · 664 25 89 77", 8, False, DARK_GRAY, PP_ALIGN.CENTER)

# ============================================
# SLIDE 8: Next steps
# ============================================
slide8 = prs.slides.add_slide(blank_layout)
slide8.shapes.add_picture(f'{BG_DIR}/salesianas_slide_8.png', Cm(0), Cm(0), width=SLIDE_WIDTH, height=SLIDE_HEIGHT)

# Client name
add_text_box(slide8, Cm(1), Cm(3.5), Cm(20), Cm(0.8), "HIJAS DE MARÍA AUXILIADORA", 14, True, WHITE, PP_ALIGN.LEFT)
# Lines in steps
add_text_box(slide8, Cm(13.5), Cm(7.2), Cm(5), Cm(0.6), "730 líneas móviles", 11, False, DARK_GRAY, PP_ALIGN.LEFT)

# Commercial contact
add_text_box(slide8, Cm(1), Cm(14.5), Cm(27), Cm(0.8), "Javier Martín Amador  ·  664 25 89 77  ·  javier.martin@lynks-tic.com", 14, True, WHITE, PP_ALIGN.CENTER)

# Save
output_path = f"{BASE_DIR}/templates/lynks-tic_salesianas_base.pptx"
prs.save(output_path)
print(f"Template saved: {output_path}")
print(f"Slides: {len(prs.slides)}")
print(f"Size: {os.path.getsize(output_path)} bytes")