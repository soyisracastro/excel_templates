"""
Genera guia_rapida.pdf — Guía Rápida de la plantilla Deducción de Inversiones LISR.
Marca: Columna 13 | Tamaño: Letter (Carta) | Sin logos.

Ejecutar:
    python3 investment_deductions/generate_guia_pdf.py

Genera:
    investment_deductions/guia_rapida.pdf
"""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ── Brand colours ───────────────────────────────────────────────
NAVY = colors.HexColor("#1B2A4A")       # primary dark
BLUE = colors.HexColor("#3498DB")       # accent
GREEN_DARK = colors.HexColor("#1E6B3A") # header green (matches Excel)
WHITE = colors.white
LIGHT_GRAY = colors.HexColor("#F4F6F8")
MID_GRAY = colors.HexColor("#7F8C8D")
TEXT_COLOR = colors.HexColor("#2C3E50")
NOTE_BG = colors.HexColor("#EBF5FB")
NOTE_BORDER = colors.HexColor("#3498DB")

BRAND = "Columna 13"
TITLE = "Deducción de Inversiones LISR"
SUBTITLE = "Guía Rápida — Artículos 31 al 38 de la Ley del ISR"

OUT_DIR = Path(__file__).parent
OUT_FILE = OUT_DIR / "guia_rapida.pdf"


# ── Styles ──────────────────────────────────────────────────────
def build_styles():
    ss = getSampleStyleSheet()

    s = {
        "title": ParagraphStyle(
            "c13_title", parent=ss["Title"],
            fontName="Helvetica-Bold", fontSize=22, leading=26,
            textColor=NAVY, alignment=TA_LEFT, spaceAfter=2 * mm,
        ),
        "subtitle": ParagraphStyle(
            "c13_subtitle", parent=ss["Normal"],
            fontName="Helvetica", fontSize=11, leading=14,
            textColor=BLUE, spaceAfter=6 * mm,
        ),
        "h2": ParagraphStyle(
            "c13_h2", parent=ss["Heading2"],
            fontName="Helvetica-Bold", fontSize=14, leading=18,
            textColor=NAVY, spaceBefore=8 * mm, spaceAfter=3 * mm,
        ),
        "h3": ParagraphStyle(
            "c13_h3", parent=ss["Heading3"],
            fontName="Helvetica-Bold", fontSize=11, leading=14,
            textColor=GREEN_DARK, spaceBefore=5 * mm, spaceAfter=2 * mm,
        ),
        "body": ParagraphStyle(
            "c13_body", parent=ss["Normal"],
            fontName="Helvetica", fontSize=9.5, leading=13,
            textColor=TEXT_COLOR, spaceAfter=2 * mm,
        ),
        "bold": ParagraphStyle(
            "c13_bold", parent=ss["Normal"],
            fontName="Helvetica-Bold", fontSize=9.5, leading=13,
            textColor=TEXT_COLOR, spaceAfter=2 * mm,
        ),
        "bullet": ParagraphStyle(
            "c13_bullet", parent=ss["Normal"],
            fontName="Helvetica", fontSize=9.5, leading=13,
            textColor=TEXT_COLOR, leftIndent=12 * mm,
            bulletIndent=5 * mm, spaceAfter=1.5 * mm,
        ),
        "note": ParagraphStyle(
            "c13_note", parent=ss["Normal"],
            fontName="Helvetica-Oblique", fontSize=9, leading=12,
            textColor=MID_GRAY, leftIndent=5 * mm, spaceAfter=3 * mm,
        ),
        "footer": ParagraphStyle(
            "c13_footer", parent=ss["Normal"],
            fontName="Helvetica", fontSize=7.5, leading=10,
            textColor=MID_GRAY, alignment=TA_CENTER,
        ),
        "header_r": ParagraphStyle(
            "c13_header_r", parent=ss["Normal"],
            fontName="Helvetica", fontSize=7.5, leading=10,
            textColor=MID_GRAY, alignment=TA_RIGHT,
        ),
        "cell": ParagraphStyle(
            "c13_cell", parent=ss["Normal"],
            fontName="Helvetica", fontSize=8.5, leading=11,
            textColor=TEXT_COLOR,
        ),
        "cell_bold": ParagraphStyle(
            "c13_cell_bold", parent=ss["Normal"],
            fontName="Helvetica-Bold", fontSize=8.5, leading=11,
            textColor=TEXT_COLOR,
        ),
        "cell_header": ParagraphStyle(
            "c13_cell_header", parent=ss["Normal"],
            fontName="Helvetica-Bold", fontSize=8.5, leading=11,
            textColor=WHITE,
        ),
        "num_list": ParagraphStyle(
            "c13_num", parent=ss["Normal"],
            fontName="Helvetica", fontSize=9.5, leading=13,
            textColor=TEXT_COLOR, leftIndent=12 * mm,
            bulletIndent=5 * mm, spaceAfter=1.5 * mm,
        ),
    }
    return s


# ── Table helpers ───────────────────────────────────────────────
def make_table(headers, rows, col_widths=None, first_col_bold=True):
    """Build a styled Table flowable."""
    s = build_styles()
    data = [[Paragraph(h, s["cell_header"]) for h in headers]]
    for row in rows:
        cells = []
        for i, cell in enumerate(row):
            style = s["cell_bold"] if (i == 0 and first_col_bold) else s["cell"]
            cells.append(Paragraph(str(cell), style))
        data.append(cells)

    t = Table(data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        # Header row
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8.5),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
        # Body
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        # Grid
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDC3C7")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    # Alternating row colors
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, i), (-1, i), LIGHT_GRAY))

    t.setStyle(TableStyle(style_cmds))
    return t


def make_result_table(rows, col_widths=None):
    """Two-column result table (Columna | Valor calculado)."""
    return make_table(["Columna", "Valor calculado"], rows, col_widths)


def note_box(text):
    """A light-blue note box."""
    s = build_styles()
    data = [[Paragraph(text, s["note"])]]
    t = Table(data, colWidths=[None])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NOTE_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LINEBEFORELENGTH", (0, 0), (0, -1), 3),
        ("LINEBEFORE", (0, 0), (0, -1), 3, NOTE_BORDER),
    ]))
    return t


def hr():
    return HRFlowable(
        width="100%", thickness=0.5, color=colors.HexColor("#BDC3C7"),
        spaceBefore=4 * mm, spaceAfter=4 * mm,
    )


# ── Header / Footer ────────────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    w, h = letter
    # Top accent line
    canvas.setStrokeColor(NAVY)
    canvas.setLineWidth(2)
    canvas.line(2 * cm, h - 1.4 * cm, w - 2 * cm, h - 1.4 * cm)
    # Header text
    canvas.setFont("Helvetica-Bold", 8)
    canvas.setFillColor(NAVY)
    canvas.drawString(2 * cm, h - 1.25 * cm, BRAND)
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MID_GRAY)
    canvas.drawRightString(w - 2 * cm, h - 1.25 * cm, TITLE)
    # Footer
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MID_GRAY)
    canvas.drawCentredString(w / 2, 1.2 * cm, f"{BRAND}  •  {TITLE}  •  Página {doc.page}")
    # Bottom accent line
    canvas.setStrokeColor(colors.HexColor("#BDC3C7"))
    canvas.setLineWidth(0.5)
    canvas.line(2 * cm, 1.5 * cm, w - 2 * cm, 1.5 * cm)
    canvas.restoreState()


def on_first_page(canvas, doc):
    """First page: just footer, no header (title block handles the top)."""
    canvas.saveState()
    w, h = letter
    # Top accent bar — thicker on first page
    canvas.setStrokeColor(NAVY)
    canvas.setLineWidth(3)
    canvas.line(2 * cm, h - 1.4 * cm, w - 2 * cm, h - 1.4 * cm)
    # Footer
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MID_GRAY)
    canvas.drawCentredString(w / 2, 1.2 * cm, f"{BRAND}  •  {TITLE}  •  Página {doc.page}")
    canvas.setStrokeColor(colors.HexColor("#BDC3C7"))
    canvas.setLineWidth(0.5)
    canvas.line(2 * cm, 1.5 * cm, w - 2 * cm, 1.5 * cm)
    canvas.restoreState()


# ── Build document ──────────────────────────────────────────────
def build_pdf():
    doc = SimpleDocTemplate(
        str(OUT_FILE),
        pagesize=letter,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        title=f"{TITLE} — Guía Rápida",
        author=BRAND,
    )

    s = build_styles()
    story = []

    avail_w = doc.width  # usable width between margins

    # ── Cover / Title ───────────────────────────────────────────
    story.append(Spacer(1, 8 * mm))
    story.append(Paragraph(BRAND, s["subtitle"]))
    story.append(Paragraph(TITLE, s["title"]))
    story.append(Paragraph(SUBTITLE, s["subtitle"]))
    story.append(Paragraph(
        "Plantilla de Cálculo Fiscal | Versión 1.0 | Marzo 2026",
        s["note"],
    ))
    story.append(hr())

    # ── Contenido de la plantilla ───────────────────────────────
    story.append(Paragraph("Contenido de la plantilla", s["h2"]))
    story.append(Paragraph(
        "Tu archivo contiene <b>7 hojas</b> de trabajo:",
        s["body"],
    ))
    story.append(make_table(
        ["Hoja", "Función"],
        [
            ["Instrucciones", "Guía de uso integrada en la plantilla"],
            ["Config", "Ejercicio fiscal y topes de automóviles"],
            ["Catálogo", "Porcentajes de deducción (Art. 33, 34 y 35)"],
            ["Inversiones", "Hoja principal de cálculo"],
            ["Resumen", "Totales agrupados por tipo de bien"],
            ["Baja_Activos", "Ganancia o pérdida por venta de activos"],
            ["INPC", "Índices Nacionales de Precios al Consumidor (1984–2025)"],
        ],
        col_widths=[3.5 * cm, avail_w - 3.5 * cm],
    ))
    story.append(hr())

    # ── Paso 1 — Config ────────────────────────────────────────
    story.append(Paragraph("Paso 1 — Configura el ejercicio fiscal", s["h2"]))
    story.append(Paragraph(
        "Abre la hoja <b>Config</b> y verifica que el campo <b>Ejercicio</b> "
        "tenga el año correcto (por defecto: 2025).",
        s["body"],
    ))
    story.append(make_table(
        ["Parámetro", "Valor", "Descripción"],
        [
            ["Ejercicio", "2025", "Año fiscal para todos los cálculos"],
            ["TOPE_AUTO_COMBUSTION", "175,000", "Tope MOI deducible — autos de combustión"],
            ["TOPE_AUTO_ELECTRICO", "250,000", "Tope MOI deducible — autos eléctricos/híbridos"],
        ],
        col_widths=[4.5 * cm, 2.5 * cm, avail_w - 7 * cm],
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(note_box(
        "Los topes se aplican automáticamente en la hoja Inversiones. "
        "Si el SAT modifica estos valores, solo necesitas cambiarlos aquí."
    ))
    story.append(hr())

    # ── Paso 2 — Inversiones ───────────────────────────────────
    story.append(Paragraph("Paso 2 — Registra tus inversiones", s["h2"]))
    story.append(Paragraph(
        "En la hoja <b>Inversiones</b>, llena únicamente estas columnas:",
        s["body"],
    ))
    story.append(make_table(
        ["Columna", "Qué capturar"],
        [
            ["A — No.", "Número de control o cuenta contable"],
            ["B — Cuenta Contable", "Código contable del activo"],
            ["C — Concepto", "Descripción del bien"],
            ["D — Fecha de Adquisición", "Fecha en formato DD/MM/AAAA"],
            ["E — M.O.I.", "Monto Original de la Inversión (sin IVA)"],
            ["G — Tipo de Bien", "Selecciona del menú desplegable"],
        ],
        col_widths=[4.5 * cm, avail_w - 4.5 * cm],
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(Paragraph(
        "<b>Todo lo demás se calcula automáticamente</b>, incluyendo:",
        s["body"],
    ))
    bullets = [
        "MOI Deducible (aplica topes de automóviles)",
        "Porcentaje de deducción (del catálogo)",
        "Meses de uso en el ejercicio",
        "Deducción del ejercicio",
        "INPC y factor de actualización (a 4 decimales)",
        "Deducción actualizada",
        "Depreciación acumulada de ejercicios anteriores",
        "Saldo pendiente de deducir",
    ]
    for b in bullets:
        story.append(Paragraph(b, s["bullet"], bulletText="•"))
    story.append(hr())

    # ── Ejemplo práctico ────────────────────────────────────────
    story.append(Paragraph("Ejemplo práctico — 5 activos", s["h2"]))
    story.append(Paragraph(
        "A continuación, 5 activos que muestran los diferentes escenarios "
        "que la plantilla maneja:",
        s["body"],
    ))

    # Activo 1
    story.append(KeepTogether([
        Paragraph("Activo 1: Laptop Dell Latitude 5540", s["h3"]),
        make_table(
            ["Campo", "Valor"],
            [
                ["Fecha de Adquisición", "15/03/2023"],
                ["M.O.I.", "$32,000"],
                ["Tipo de Bien", "Equipo de cómputo"],
            ],
            col_widths=[4.5 * cm, avail_w - 4.5 * cm],
        ),
        Spacer(1, 2 * mm),
        Paragraph(
            "<b>Resultado:</b> 30% de deducción, 12 meses de uso (tercer año completo), "
            "depreciación acumulada de $16,800 por 21 meses en 2023–2024.",
            s["body"],
        ),
        make_result_table([
            ["MOI Deducible", "$32,000"],
            ["% Deducción", "30%"],
            ["Meses de Uso", "12"],
            ["Deducción del Ejercicio", "$9,600"],
            ["Dep. Acumulada", "$16,800"],
            ["Saldo Pendiente", "$5,600"],
        ], col_widths=[4.5 * cm, avail_w - 4.5 * cm]),
    ]))
    story.append(hr())

    # Activo 2
    story.append(KeepTogether([
        Paragraph("Activo 2: Escritorio y silla ejecutiva", s["h3"]),
        make_table(
            ["Campo", "Valor"],
            [
                ["Fecha de Adquisición", "01/06/2024"],
                ["M.O.I.", "$18,500"],
                ["Tipo de Bien", "Mobiliario y equipo de oficina"],
            ],
            col_widths=[4.5 * cm, avail_w - 4.5 * cm],
        ),
        Spacer(1, 2 * mm),
        Paragraph(
            "<b>Resultado:</b> 10% de deducción (vida útil 10 años). "
            "Dep. acumulada de $925 por 6 meses de uso en 2024.",
            s["body"],
        ),
        make_result_table([
            ["MOI Deducible", "$18,500"],
            ["% Deducción", "10%"],
            ["Meses de Uso", "12"],
            ["Deducción del Ejercicio", "$1,850"],
            ["Dep. Acumulada", "$925"],
            ["Saldo Pendiente", "$15,725"],
        ], col_widths=[4.5 * cm, avail_w - 4.5 * cm]),
    ]))
    story.append(hr())

    # Activo 3
    story.append(KeepTogether([
        Paragraph("Activo 3: Toyota Corolla 2024 (combustión)", s["h3"]),
        make_table(
            ["Campo", "Valor"],
            [
                ["Fecha de Adquisición", "10/01/2024"],
                ["M.O.I.", "$420,000"],
                ["Tipo de Bien", "Automóvil (combustión)"],
            ],
            col_widths=[4.5 * cm, avail_w - 4.5 * cm],
        ),
        Spacer(1, 2 * mm),
        Paragraph(
            "<b>Resultado:</b> MOI de $420,000 excede el tope de $175,000. "
            "La plantilla recorta automáticamente. El contribuyente \"pierde\" $245,000.",
            s["body"],
        ),
        make_result_table([
            ["MOI Deducible", "$175,000 (tope aplicado)"],
            ["% Deducción", "25%"],
            ["Meses de Uso", "12"],
            ["Deducción del Ejercicio", "$43,750"],
            ["Dep. Acumulada", "$40,104"],
            ["Saldo Pendiente", "$91,146"],
        ], col_widths=[4.5 * cm, avail_w - 4.5 * cm]),
        Spacer(1, 2 * mm),
        note_box(
            "Si el vehículo fuera una pick-up (camión de carga), no tendría tope "
            "y el MOI Deducible sería $420,000 completos (Criterio 27/ISR/N)."
        ),
    ]))
    story.append(hr())

    # Activo 4
    story.append(KeepTogether([
        Paragraph("Activo 4: BYD Dolphin Mini 2025 (eléctrico)", s["h3"]),
        make_table(
            ["Campo", "Valor"],
            [
                ["Fecha de Adquisición", "20/02/2025"],
                ["M.O.I.", "$299,000"],
                ["Tipo de Bien", "Automóvil (eléctrico/híbrido)"],
            ],
            col_widths=[4.5 * cm, avail_w - 4.5 * cm],
        ),
        Spacer(1, 2 * mm),
        Paragraph(
            "<b>Resultado:</b> Tope eléctrico de $250,000 aplicado. "
            "Primer año parcial: 10 meses (marzo a diciembre). Dep. acumulada = $0.",
            s["body"],
        ),
        make_result_table([
            ["MOI Deducible", "$250,000 (tope eléctrico)"],
            ["% Deducción", "25%"],
            ["Meses de Uso", "10 (primer año parcial)"],
            ["Deducción del Ejercicio", "$52,083"],
            ["Dep. Acumulada", "$0"],
            ["Saldo Pendiente", "$197,917"],
        ], col_widths=[4.5 * cm, avail_w - 4.5 * cm]),
    ]))
    story.append(hr())

    # Activo 5
    story.append(KeepTogether([
        Paragraph("Activo 5: Servidor HPE ProLiant DL380", s["h3"]),
        make_table(
            ["Campo", "Valor"],
            [
                ["Fecha de Adquisición", "01/07/2022"],
                ["M.O.I.", "$95,000"],
                ["Tipo de Bien", "Equipo de cómputo"],
            ],
            col_widths=[4.5 * cm, avail_w - 4.5 * cm],
        ),
        Spacer(1, 2 * mm),
        Paragraph(
            "<b>Resultado:</b> Vida útil de 40 meses (30%). Ya acumula 29 meses en ejercicios "
            "anteriores, le quedan 11 meses. Queda totalmente depreciado al cierre de 2025.",
            s["body"],
        ),
        make_result_table([
            ["MOI Deducible", "$95,000"],
            ["% Deducción", "30%"],
            ["Meses de Uso", "11 (últimos meses de vida útil)"],
            ["Deducción del Ejercicio", "$26,125"],
            ["Dep. Acumulada", "$68,875"],
            ["Saldo Pendiente", "$0"],
        ], col_widths=[4.5 * cm, avail_w - 4.5 * cm]),
    ]))
    story.append(hr())

    # ── Paso 3 — Resumen ────────────────────────────────────────
    story.append(Paragraph("Paso 3 — Consulta el resumen", s["h2"]))
    story.append(Paragraph(
        "La hoja <b>Resumen</b> agrupa automáticamente los totales por tipo de bien. "
        "Con los 5 activos de ejemplo:",
        s["body"],
    ))

    # Build summary table manually for the TOTAL bold row
    sum_headers = ["Tipo de Bien", "Cant.", "MOI Deducible", "Deducción Ejercicio", "Saldo Pendiente"]
    sum_rows = [
        ["Equipo de cómputo", "2", "$127,000", "$35,725", "$5,600"],
        ["Mobiliario y eq. oficina", "1", "$18,500", "$1,850", "$15,725"],
        ["Automóvil (combustión)", "1", "$175,000", "$43,750", "$91,146"],
        ["Automóvil (eléctrico)", "1", "$250,000", "$52,083", "$197,917"],
    ]

    # Build data with Paragraph objects
    hdr_cells = [Paragraph(h, s["cell_header"]) for h in sum_headers]
    body_cells = []
    for row in sum_rows:
        body_cells.append([
            Paragraph(row[0], s["cell_bold"]),
            *[Paragraph(c, s["cell"]) for c in row[1:]],
        ])
    # Total row
    total_row = [
        Paragraph("<b>TOTAL</b>", s["cell_bold"]),
        Paragraph("<b>5</b>", s["cell_bold"]),
        Paragraph("<b>$570,500</b>", s["cell_bold"]),
        Paragraph("<b>$133,408</b>", s["cell_bold"]),
        Paragraph("<b>$310,388</b>", s["cell_bold"]),
    ]
    all_data = [hdr_cells] + body_cells + [total_row]

    sum_cw = [4.5 * cm, 1.5 * cm, 3.2 * cm, 3.5 * cm, 3.2 * cm]
    sum_t = Table(all_data, colWidths=sum_cw, repeatRows=1)
    n_rows = len(all_data)
    sum_style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDC3C7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        # Total row highlight
        ("BACKGROUND", (0, n_rows - 1), (-1, n_rows - 1), colors.HexColor("#D5E8D4")),
        ("LINEABOVE", (0, n_rows - 1), (-1, n_rows - 1), 1.2, NAVY),
    ]
    for i in range(1, n_rows - 1):
        if i % 2 == 0:
            sum_style.append(("BACKGROUND", (0, i), (-1, i), LIGHT_GRAY))
    sum_t.setStyle(TableStyle(sum_style))
    story.append(sum_t)

    story.append(Spacer(1, 2 * mm))
    story.append(note_box(
        "Estos totales son útiles para tu declaración anual y reportes financieros."
    ))
    story.append(hr())

    # ── Paso 4 — Baja de activos ────────────────────────────────
    story.append(Paragraph("Paso 4 — Baja de activos (si aplica)", s["h2"]))
    story.append(Paragraph(
        "Si vendes o das de baja un activo, usa la hoja <b>Baja_Activos</b>. Ejemplo:",
        s["body"],
    ))
    story.append(Paragraph(
        "<b>Impresora Multifuncional HP — vendida en junio 2025</b>",
        s["bold"],
    ))
    story.append(make_table(
        ["Campo", "Valor"],
        [
            ["Concepto", "Impresora Multifuncional HP LaserJet Pro"],
            ["MOI Deducible", "$15,000"],
            ["Deducciones Acumuladas", "$15,000"],
            ["INPC Enajenación (Jun 2025)", "140.405"],
            ["INPC Adquisición (Mar 2021)", "111.824"],
            ["Precio de Venta (sin IVA)", "$3,500"],
        ],
        col_widths=[5 * cm, avail_w - 5 * cm],
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(make_table(
        ["Cálculo", "Resultado"],
        [
            ["Saldo Pendiente", "$0 (totalmente depreciado)"],
            ["Factor de Actualización", "1.2556"],
            ["Saldo Actualizado", "$0"],
            ["Ganancia / Pérdida", "$3,500"],
            ["Resultado", "Ganancia Acumulable"],
        ],
        col_widths=[5 * cm, avail_w - 5 * cm],
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(note_box(
        "La ganancia es ingreso acumulable para ISR. "
        "Si el resultado fuera negativo, sería una pérdida deducible."
    ))
    story.append(hr())

    # ── Paso 5 — INPC ──────────────────────────────────────────
    story.append(Paragraph("Paso 5 — Actualizar el INPC", s["h2"]))
    story.append(Paragraph(
        "La hoja <b>INPC</b> contiene datos desde 1984 hasta 2025. "
        "Para agregar un año nuevo:",
        s["body"],
    ))
    steps = [
        "Ve a la última fila (2025)",
        "Inserta una fila debajo",
        "Escribe el año (2026) en la columna A",
        "Llena los valores mensuales conforme el INEGI los publique",
    ]
    for i, step in enumerate(steps, 1):
        story.append(Paragraph(f"{i}. {step}", s["num_list"]))
    story.append(Spacer(1, 2 * mm))
    story.append(note_box(
        "Los valores de INPC se publican quincenalmente en inegi.org.mx/temas/inpc/"
    ))
    story.append(hr())

    # ── Topes de automóviles ────────────────────────────────────
    story.append(Paragraph("Topes de automóviles (Art. 36 LISR)", s["h2"]))
    story.append(make_table(
        ["Tipo de vehículo", "Tope deducible"],
        [
            ["Combustión interna", "$175,000 MXN"],
            ["Eléctrico o híbrido", "$250,000 MXN"],
            ["Pick-up (camión de carga)", "Sin tope (100% deducible)"],
        ],
        col_widths=[5.5 * cm, avail_w - 5.5 * cm],
    ))
    story.append(Spacer(1, 2 * mm))
    story.append(Paragraph(
        "Las pick-up se clasifican como camiones de carga conforme al "
        "Criterio 27/ISR/N del SAT.",
        s["note"],
    ))
    story.append(hr())

    # ── Notas importantes ───────────────────────────────────────
    story.append(Paragraph("Notas importantes", s["h2"]))
    notes = [
        "El <b>IVA no forma parte del MOI</b> (es acreditable), salvo que "
        "no tengas derecho al acreditamiento.",
        "Si <b>no deduces en el ejercicio de inicio de uso</b> ni en el "
        "siguiente, pierdes esos montos de forma permanente.",
        "Puedes aplicar un <b>porcentaje menor al máximo</b>, pero queda "
        "fijo por 5 años (Art. 66 RLISR).",
        "El <b>Factor de Actualización</b> se calcula a 4 decimales conforme "
        "al Art. 9 del Reglamento de la LISR.",
        "La <b>Dep. Acumulada se calcula automáticamente</b>. Si tu "
        "depreciación real difiere por ajustes o porcentajes menores, "
        "puedes sobreescribir la fórmula directamente en la celda.",
        "Para bienes de <b>energía renovable</b> (100% deducible), el "
        "sistema debe operar al menos 5 años continuos.",
    ]
    for n in notes:
        story.append(Paragraph(n, s["bullet"], bulletText="•"))
    story.append(hr())

    # ── Soporte ─────────────────────────────────────────────────
    story.append(Paragraph("Soporte", s["h2"]))
    story.append(Paragraph(
        "Si tienes dudas sobre cómo usar esta plantilla, revisa el video "
        "tutorial disponible en el blog o responde al correo de confirmación "
        "de compra.",
        s["body"],
    ))
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph(
        f"<b>{BRAND}</b>  •  Versión 1.0  •  Marzo 2026",
        s["footer"],
    ))

    # ── Build ───────────────────────────────────────────────────
    doc.build(story, onFirstPage=on_first_page, onLaterPages=on_page)
    print(f"✓ PDF generado: {OUT_FILE}")
    print(f"  Tamaño: Letter (Carta) — {OUT_FILE.stat().st_size / 1024:.0f} KB")


if __name__ == "__main__":
    build_pdf()
