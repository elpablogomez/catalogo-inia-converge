import json
import csv
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# 1. Cargar datos base
with open('soluciones_converge_fechas_completas.json', 'r', encoding='utf-8') as f:
    items = json.load(f)

# Convocatoria / Postulación extraída de la solapa de Google Sheets
convocatorias_map = {
    "DeepAgro": "Septiembre 2023",
    "Nettra_Riego": "Septiembre 2023",
    "AgriTechnovation": "Abril 2024",
    "Agroyeitos": "Abril 2024",
    "Auravant": "Abril 2024",
    "Baqueano": "Abril 2024",
    "BioGuardians": "Abril 2024",
    "CompassForest": "Abril 2024",
    "HerdVision": "Abril 2024",
    "metaBix": "Abril 2024",
    "metaBix Avícola": "Abril 2024",
    "Oryzativa": "Abril 2024",
    "SmaXtec": "Abril 2024",
    "Smartway": "Abril 2024",
    "Ucrop.it": "Abril 2024",
    "Wisflow": "Abril 2024",
    "Autolink": "Octubre 2024",
    "Cattler": "Octubre 2024",
    "Insights": "Octubre 2024",
    "Nettra": "Octubre 2024",
    "Pastech": "Octubre 2024",
    "SensorData": "Mayo 2025",
    "TGA": "Octubre 2025"
}

periodos_curados = {
    "AgriTechnovation": "Setiembre 2024 a Febrero 2025",
    "Agroyeitos": "Septiembre 2024 a Febrero 2025 (Zafra 2023-2024)",
    "Auravant": "Septiembre 2024 a Febrero 2025",
    "Smartway": "Octubre 2024 a Enero 2025",
    "Ucrop.it": "Enero a Mayo 2025 (Zafra 2024)",
    "CompassForest": "Noviembre 2024 (Ensayos en Cerro Colorado)",
    "Oryzativa": "Zafras de Arroz 2022 a 2024",
    "BioGuardians": "Campañas agrícolas 2024–2025",
    "DeepAgro": "Febrero a Julio 2025 (Barbecho en 5 chacras)",
    "Baqueano": "Diciembre 2024 (Ensayos campo set-oct 2024)",
    "Cattler": "Abril 2025 (Jornada técnica en feedlot)",
    "Insights": "Abril a Agosto 2025 (Sede INIA Glencoe)",
    "Pastech": "Mayo a Octubre 2025",
    "HerdVision": "Julio a Octubre 2025",
    "SmaXtec": "Años 2020 a 2021 (Ensayos corral y pH)",
    "metaBix": "Agosto a Noviembre 2025 (9 sem. recría/preengorde)",
    "metaBix Avícola": "Octubre a Diciembre 2025 (Engorde y postura)",
    "Autolink": "Febrero a Abril 2025",
    "Nettra_Riego": "Junio 2024 a Junio 2025 (INIA Las Brujas)",
    "Wisflow": "Enero a Febrero 2025 (INIA La Estanzuela)",
    "SensorData": "Ciclos de aplicación 2025–2026 (Vid y frutales)",
    "Nettra": "Setiembre a Diciembre 2024 (Verdeagua)",
    "TGA": "Febrero a Julio 2026 (Pruebas operativas en campo)"
}

for item in items:
    name = item["solucion"]
    item["convocatoria_postulacion"] = convocatorias_map.get(name, "-")
    if name in periodos_curados:
        item["periodo_evaluacion"] = periodos_curados[name]

# 2. Guardar JSON
with open('soluciones_inia_converge.json', 'w', encoding='utf-8') as f:
    json.dump(items, f, ensure_ascii=False, indent=2)

# 3. Guardar CSV
with open('soluciones_inia_converge.csv', 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow([
        "Categoría",
        "Subcategoría",
        "Solución / Empresa",
        "Convocatoria Postulación",
        "Período Verificación Técnica",
        "Fecha Publicación Web",
        "Fecha Testimonio Validación",
        "URL Ficha INIA",
        "URL Reporte Verificación (PDF)",
        "URL Informe Detallado Extendido (PDF)",
        "URL Testimonio Validación",
        "Tipo Testimonio",
        "URL Ficha Gráfica",
        "URL Sitio Web Oficial"
    ])
    for item in items:
        writer.writerow([
            item["categoria"],
            item["subcategoria"],
            item["solucion"],
            item["convocatoria_postulacion"],
            item["periodo_evaluacion"],
            item["fecha_publicacion_web"],
            item["fecha_testimonio"],
            item["url_inia"],
            item["reporte_verificacion"],
            item["informe_detallado"],
            item["testimonio_validacion"],
            item["tipo_testimonio"],
            item["ficha_grafica"],
            item["pagina_web"]
        ])

# 4. Guardar Excel profesional
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Soluciones INIA Converge"
ws.views.sheetView[0].showGridLines = True

headers = [
    "Categoría",
    "Subcategoría",
    "Solución / Empresa",
    "Postulación",
    "Período Verificación",
    "Publicación Web",
    "Ficha INIA",
    "Reporte Verificación",
    "Informe Detallado",
    "Testimonio",
    "Ficha Gráfica",
    "Web Oficial"
]

header_fill = PatternFill(start_color="1A472A", end_color="1A472A", fill_type="solid")
header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
data_font = Font(name="Segoe UI", size=10)
link_font = Font(name="Segoe UI", size=10, color="0B5394", underline="single")
zebra_fill = PatternFill(start_color="F7FAF8", end_color="F7FAF8", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="E0E0E0"),
    right=Side(style="thin", color="E0E0E0"),
    top=Side(style="thin", color="E0E0E0"),
    bottom=Side(style="thin", color="E0E0E0")
)

ws.append(headers)
for col_idx in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border
ws.row_dimensions[1].height = 28

for row_idx, item in enumerate(items, start=2):
    ws.cell(row=row_idx, column=1, value=item["categoria"]).alignment = Alignment(vertical="center")
    ws.cell(row=row_idx, column=2, value=item["subcategoria"]).alignment = Alignment(vertical="center")
    ws.cell(row=row_idx, column=3, value=item["solucion"]).alignment = Alignment(vertical="center")
    
    c_conv = ws.cell(row=row_idx, column=4, value=item["convocatoria_postulacion"])
    c_conv.alignment = Alignment(horizontal="center", vertical="center")
    c_conv.font = data_font
    
    c_per = ws.cell(row=row_idx, column=5, value=item["periodo_evaluacion"])
    c_per.alignment = Alignment(vertical="center", wrap_text=True)
    c_per.font = data_font
    
    c_pub = ws.cell(row=row_idx, column=6, value=item["fecha_publicacion_web"])
    c_pub.alignment = Alignment(horizontal="center", vertical="center")
    c_pub.font = data_font
    
    def set_link(col, url, label):
        c = ws.cell(row=row_idx, column=col)
        if url:
            c.value = f'=HYPERLINK("{url}", "{label}")'
            c.font = link_font
        else:
            c.value = "-"
            c.font = data_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    set_link(7, item["url_inia"], "Ver Ficha")
    set_link(8, item["reporte_verificacion"], "Reporte (PDF)")
    set_link(9, item["informe_detallado"], "Extendido (PDF)")
    set_link(10, item["testimonio_validacion"], f"Testimonio ({item['tipo_testimonio']})")
    set_link(11, item["ficha_grafica"], "Ficha Gráfica")
    set_link(12, item["pagina_web"], "Web Oficial")
    
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        c.border = thin_border
        if col_idx <= 3:
            c.font = data_font
        if row_idx % 2 == 1:
            c.fill = zebra_fill
            
    ws.row_dimensions[row_idx].height = 24

col_widths = {
    1: 18, # Categoría
    2: 24, # Subcategoría
    3: 20, # Solución
    4: 18, # Postulación
    5: 32, # Período Verificación
    6: 16, # Publicación Web
    7: 15, # Ficha INIA
    8: 18, # Reporte
    9: 18, # Informe Detallado
    10: 18,# Testimonio
    11: 15,# Ficha Gráfica
    12: 16 # Web
}
for col_idx, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col_idx)].width = width

ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(items) + 1}"
wb.save("soluciones_inia_converge.xlsx")

# 5. Generar PDF Profesional Independiente con ReportLab (Horizontal A4, nunca cortado)
def generar_pdf(dataset, filename="soluciones_inia_converge.pdf"):
    doc = SimpleDocTemplate(
        filename,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=15,
        textColor=colors.HexColor('#134e2c'),
        spaceAfter=4,
        spaceBefore=0
    )
    
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        textColor=colors.HexColor('#475569'),
        spaceAfter=10
    )
    
    cell_style = ParagraphStyle(
        'Cell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor('#1e293b')
    )
    
    cell_bold = ParagraphStyle(
        'CellBold',
        parent=cell_style,
        fontName='Helvetica-Bold'
    )
    
    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9,
        textColor=colors.white,
        alignment=1
    )
    
    link_style = ParagraphStyle(
        'LinkStyle',
        parent=cell_style,
        textColor=colors.HexColor('#0b5394'),
        alignment=1
    )
    
    story = []
    story.append(Paragraph("INIA Converge — Catálogo de Soluciones y Materiales", title_style))
    story.append(Paragraph(f"Plataforma de validación técnica y difusión de tecnologías AgTech • Generado: {datetime.now().strftime('%d/%m/%Y')}", subtitle_style))
    
    # Tabla de datos
    # Columnas: Solución, Categoría, Subcategoría, Postulación, Verificación, Publicación, Reporte, Testimonio, Web
    col_widths_pdf = [75, 75, 85, 65, 160, 65, 75, 75, 75, 52]
    
    table_data = [[
        Paragraph("Solución", header_style),
        Paragraph("Categoría", header_style),
        Paragraph("Subcategoría", header_style),
        Paragraph("Postulación", header_style),
        Paragraph("Período Verificación", header_style),
        Paragraph("Publicación", header_style),
        Paragraph("Reporte", header_style),
        Paragraph("Detallado", header_style),
        Paragraph("Testimonio", header_style),
        Paragraph("Web", header_style),
    ]]
    
    for item in dataset:
        rep_p = Paragraph(f'<a href="{item["reporte_verificacion"]}">PDF</a>', link_style) if item["reporte_verificacion"] else Paragraph("-", cell_style)
        det_p = Paragraph(f'<a href="{item["informe_detallado"]}">Extendido</a>', link_style) if item["informe_detallado"] else Paragraph("-", cell_style)
        test_p = Paragraph(f'<a href="{item["testimonio_validacion"]}">{item["tipo_testimonio"]}</a>', link_style) if item["testimonio_validacion"] else Paragraph("-", cell_style)
        web_p = Paragraph(f'<a href="{item["pagina_web"]}">Sitio</a>', link_style) if item["pagina_web"] else Paragraph("-", cell_style)
        
        table_data.append([
            Paragraph(item["solucion"], cell_bold),
            Paragraph(item["categoria"], cell_style),
            Paragraph(item["subcategoria"], cell_style),
            Paragraph(item["convocatoria_postulacion"], cell_style),
            Paragraph(item["periodo_evaluacion"], cell_style),
            Paragraph(item["fecha_publicacion_web"], cell_style),
            rep_p,
            det_p,
            test_p,
            web_p
        ])
        
    t = Table(table_data, colWidths=col_widths_pdf, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a472a')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8fafc')])
    ]))
    story.append(t)
    doc.build(story)

generar_pdf(items)

# 6. Guardar HTML Rediseñado (3 tarjetas sin verificación técnica, columna postulación, descarga PDF directa)
now_str = datetime.now().strftime("%d/%m/%Y")

html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Catálogo de Soluciones - INIA Converge</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
  <style>
    :root {{
      --inia-green: #1a5e38;
      --inia-light-green: #eaf5ee;
      --inia-accent: #2e7d32;
    }}
    body {{
      background-color: #f8fafc;
      font-family: system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      color: #1e293b;
    }}
    .header-banner {{
      background: linear-gradient(135deg, #134e2c 0%, #1e6b3f 100%);
      color: white;
      padding: 1.75rem 1rem;
      border-bottom: 4px solid #103d22;
      margin-bottom: 1.5rem;
      box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    }}
    .stats-card {{
      background: white;
      border-radius: 10px;
      padding: 1.1rem;
      border: 1px solid #e2e8f0;
      box-shadow: 0 2px 5px rgba(0,0,0,0.02);
      text-align: center;
    }}
    .stats-num {{
      font-size: 2rem;
      font-weight: 700;
      color: var(--inia-green);
    }}
    .table-container {{
      background: white;
      border-radius: 12px;
      border: 1px solid #e2e8f0;
      box-shadow: 0 4px 14px rgba(0,0,0,0.03);
      padding: 1rem;
      width: 100%;
    }}
    .table {{
      font-size: 0.82rem;
      margin-bottom: 0;
      width: 100%;
    }}
    .table thead th {{
      background-color: var(--inia-light-green);
      color: var(--inia-green);
      font-weight: 600;
      border-bottom: 2px solid #cce8d5;
      padding: 0.55rem 0.35rem;
      vertical-align: middle;
      font-size: 0.8rem;
    }}
    .table tbody td {{
      padding: 0.5rem 0.35rem;
      vertical-align: middle;
    }}
    .btn-action {{
      padding: 0.22rem 0.45rem;
      font-size: 0.74rem;
      border-radius: 5px;
      font-weight: 500;
      display: inline-flex;
      align-items: center;
      gap: 0.25rem;
      text-decoration: none;
      transition: all 0.15s ease;
      margin: 1px;
      white-space: nowrap;
    }}
    .btn-pdf {{
      background-color: #fef2f2;
      color: #b91c1c;
      border: 1px solid #fecaca;
    }}
    .btn-pdf:hover {{
      background-color: #dc2626;
      color: white;
    }}
    .btn-detail {{
      background-color: #fff7ed;
      color: #c2410c;
      border: 1px solid #ffedd5;
    }}
    .btn-detail:hover {{
      background-color: #ea580c;
      color: white;
    }}
    .btn-video {{
      background-color: #eff6ff;
      color: #1d4ed8;
      border: 1px solid #bfdbfe;
    }}
    .btn-video:hover {{
      background-color: #2563eb;
      color: white;
    }}
    .btn-ficha {{
      background-color: #f0fdf4;
      color: #15803d;
      border: 1px solid #bbf7d0;
    }}
    .btn-ficha:hover {{
      background-color: #16a34a;
      color: white;
    }}
    .btn-web {{
      background-color: #faf5ff;
      color: #7e22ce;
      border: 1px solid #e9d5ff;
    }}
    .btn-web:hover {{
      background-color: #9333ea;
      color: white;
    }}
    .badge-cat {{
      font-size: 0.74rem;
      padding: 0.22rem 0.45rem;
      border-radius: 12px;
      font-weight: 600;
      display: inline-block;
      white-space: nowrap;
    }}
    .cat-vegetal-extensivo {{
      background-color: #e0f2fe;
      color: #0369a1;
    }}
    .cat-produccin-animal {{
      background-color: #fef3c7;
      color: #92400e;
    }}
    .cat-vegetal-intensivo {{
      background-color: #dcfce7;
      color: #15803d;
    }}
    .date-pub {{
      font-size: 0.78rem;
      font-weight: 600;
      color: #0f5132;
      background: #d1e7dd;
      padding: 0.2rem 0.45rem;
      border-radius: 4px;
      display: inline-block;
      white-space: nowrap;
    }}
    .badge-conv {{
      font-size: 0.74rem;
      font-weight: 600;
      color: #4338ca;
      background: #e0e7ff;
      padding: 0.2rem 0.45rem;
      border-radius: 4px;
      display: inline-block;
      white-space: nowrap;
    }}
    .periodo-box {{
      font-size: 0.75rem;
      line-height: 1.25;
      color: #334155;
      max-width: 160px;
      word-break: normal;
      overflow-wrap: break-word;
      display: inline-block;
    }}
    .search-input {{
      border-radius: 8px;
      padding: 0.55rem 0.85rem;
      border: 1px solid #cbd5e1;
      font-size: 0.9rem;
    }}
    .search-input:focus {{
      border-color: var(--inia-green);
      box-shadow: 0 0 0 0.2rem rgba(26, 94, 56, 0.15);
    }}

    /* Modo impresión en PDF sin cortes */
    @page {{
      size: A4 landscape;
      margin: 8mm 6mm;
    }}
    @media print {{
      .header-banner .btn, #searchInput, #categoryFilter, .card, footer, .stats-row {{
        display: none !important;
      }}
      .header-banner {{
        padding: 0.5rem 0 !important;
        background: none !important;
        color: #000 !important;
        border-bottom: 2px solid #1a5e38 !important;
        margin-bottom: 0.5rem !important;
      }}
      .header-banner h1, .header-banner p {{
        color: #000 !important;
      }}
      .table-container {{
        box-shadow: none !important;
        border: none !important;
        padding: 0 !important;
      }}
      .table {{
        font-size: 6.8pt !important;
        width: 100% !important;
      }}
      .table th, .table td {{
        padding: 2px 3px !important;
      }}
      .btn-action {{
        padding: 1px 2px !important;
        font-size: 6.2pt !important;
      }}
      .badge-cat, .badge-conv, .date-pub {{
        font-size: 6.2pt !important;
        padding: 1px 3px !important;
      }}
      body {{
        background: #fff !important;
      }}
    }}
  </style>
</head>
<body>

  <div class="header-banner">
    <div class="container-fluid px-lg-4">
      <div class="d-flex flex-wrap justify-content-between align-items-center gap-3">
        <div>
          <span class="badge bg-light text-success mb-1 px-3 py-1 fw-bold">Sello INIA Converge</span>
          <h1 class="h3 fw-bold mb-1"><i class="fa-solid fa-leaf me-2"></i>Catálogo de Soluciones y Materiales</h1>
          <p class="mb-0 text-white-50 small">Plataforma de validación técnica y difusión de tecnologías AgTech</p>
        </div>
        <div class="d-flex gap-2">
          <a href="soluciones_inia_converge.xlsx" class="btn btn-light fw-semibold text-success shadow-sm" download>
            <i class="fa-solid fa-file-excel text-success me-1"></i> Descargar Excel
          </a>
          <a href="soluciones_inia_converge.pdf" class="btn btn-light fw-semibold text-danger shadow-sm" download>
            <i class="fa-solid fa-file-pdf text-danger me-1"></i> Descargar PDF
          </a>
        </div>
      </div>
    </div>
  </div>

  <div class="container-fluid px-lg-4 pb-5">

    <!-- 3 Tarjetas de Resumen -->
    <div class="row g-3 mb-3 stats-row">
      <div class="col-md-4 col-4">
        <div class="stats-card">
          <div class="stats-num" id="count-total">{len(items)}</div>
          <div class="text-muted small fw-semibold">Soluciones Publicadas</div>
        </div>
      </div>
      <div class="col-md-4 col-4">
        <div class="stats-card">
          <div class="stats-num">3</div>
          <div class="text-muted small fw-semibold">Áreas Productivas</div>
        </div>
      </div>
      <div class="col-md-4 col-4">
        <div class="stats-card">
          <div class="stats-num">13</div>
          <div class="text-muted small fw-semibold">Subcategorías</div>
        </div>
      </div>
    </div>

    <!-- Filtros y Búsqueda -->
    <div class="card border-0 shadow-sm p-3 mb-3 rounded-3">
      <div class="row g-2 align-items-center">
        <div class="col-md-5">
          <div class="input-group">
            <span class="input-group-text bg-white border-end-0 text-muted"><i class="fa-solid fa-magnifying-glass"></i></span>
            <input type="text" id="searchInput" class="form-control border-start-0 search-input" placeholder="Buscar por solución, categoría, convocatoria o período...">
          </div>
        </div>
        <div class="col-md-4">
          <select id="categoryFilter" class="form-select search-input">
            <option value="">Todas las categorías</option>
            <option value="Vegetal Extensivo">Vegetal Extensivo</option>
            <option value="Producción Animal">Producción Animal</option>
            <option value="Vegetal Intensivo">Vegetal Intensivo</option>
          </select>
        </div>
        <div class="col-md-3 text-md-end text-muted small">
          Mostrando <strong id="visibleCount">{len(items)}</strong> de {len(items)} soluciones
        </div>
      </div>
    </div>

    <!-- Tabla Adaptada -->
    <div class="table-container">
      <div class="table-responsive">
        <table class="table table-hover align-middle mb-0" id="solutionsTable">
          <thead>
            <tr>
              <th style="min-width: 105px;">Solución / Empresa</th>
              <th style="min-width: 95px;">Categoría</th>
              <th style="min-width: 105px;">Subcategoría</th>
              <th style="min-width: 85px; text-align: center;">Postulación</th>
              <th style="min-width: 140px;">Período Verificación</th>
              <th style="min-width: 80px; text-align: center;">Publicación</th>
              <th style="min-width: 65px; text-align: center;">Ficha INIA</th>
              <th style="min-width: 125px; text-align: center;">Reportes Técnicos</th>
              <th style="min-width: 70px; text-align: center;">Testimonio</th>
              <th style="min-width: 60px; text-align: center;">Ficha</th>
              <th style="min-width: 60px; text-align: center;">Web</th>
            </tr>
          </thead>
          <tbody>
"""

for item in items:
    cat_class = "cat-" + re.sub(r'[^a-zA-Z0-9]', '-', item["categoria"].lower())
    
    # Convocatoria / Postulación
    conv_html = f'<span class="badge-conv">{item["convocatoria_postulacion"]}</span>'
    
    # Fecha de publicación
    fecha_html = f'<span class="date-pub">{item["fecha_publicacion_web"]}</span>'
    
    # Período de verificación en 2 líneas
    periodo_html = f'<div class="periodo-box"><i class="fa-regular fa-clock text-muted me-1"></i>{item["periodo_evaluacion"]}</div>'
    
    # Enlace ficha INIA
    inia_btn = f'<a href="{item["url_inia"]}" target="_blank" class="btn btn-action btn-outline-secondary" title="Ficha en portal INIA"><i class="fa-solid fa-arrow-up-right-from-square"></i> INIA</a>'
    
    # Reportes
    reportes_html = ""
    if item["reporte_verificacion"]:
        reportes_html += f'<a href="{item["reporte_verificacion"]}" target="_blank" class="btn btn-action btn-pdf" title="Reporte de Verificación"><i class="fa-solid fa-file-pdf"></i> Reporte</a>'
    if item["informe_detallado"]:
        reportes_html += f'<a href="{item["informe_detallado"]}" target="_blank" class="btn btn-action btn-detail" title="Informe Técnico Detallado"><i class="fa-solid fa-book-open"></i> Detallado</a>'
    if not reportes_html:
        reportes_html = '<span class="text-muted small">-</span>'
        
    # Testimonio
    if item["testimonio_validacion"]:
        icon = "fa-video" if item["tipo_testimonio"] == "Video" else "fa-file-pdf"
        btn_cls = "btn-video" if item["tipo_testimonio"] == "Video" else "btn-pdf"
        valid_btn = f'<a href="{item["testimonio_validacion"]}" target="_blank" class="btn btn-action {btn_cls}" title="Testimonio de Validación"><i class="fa-solid {icon}"></i> {item["tipo_testimonio"]}</a>'
    else:
        valid_btn = '<span class="text-muted small">-</span>'
        
    # Ficha gráfica
    if item["ficha_grafica"]:
        ficha_btn = f'<a href="{item["ficha_grafica"]}" target="_blank" class="btn btn-action btn-ficha" title="Ficha Gráfica"><i class="fa-solid fa-image"></i> Ficha</a>'
    else:
        ficha_btn = '<span class="text-muted small">-</span>'
        
    # Web oficial
    if item["pagina_web"]:
        web_btn = f'<a href="{item["pagina_web"]}" target="_blank" class="btn btn-action btn-web" title="Sitio Web Oficial"><i class="fa-solid fa-globe"></i> Web</a>'
    else:
        web_btn = '<span class="text-muted small">-</span>'

    html_content += f"""          <tr data-cat="{item['categoria']}">
              <td class="fw-bold text-dark">{item['solucion']}</td>
              <td><span class="badge-cat {cat_class}">{item['categoria']}</span></td>
              <td class="text-secondary small">{item['subcategoria']}</td>
              <td class="text-center">{conv_html}</td>
              <td>{periodo_html}</td>
              <td class="text-center">{fecha_html}</td>
              <td class="text-center">{inia_btn}</td>
              <td class="text-center">{reportes_html}</td>
              <td class="text-center">{valid_btn}</td>
              <td class="text-center">{ficha_btn}</td>
              <td class="text-center">{web_btn}</td>
            </tr>
"""

html_content += f"""          </tbody>
        </table>
      </div>
    </div>
    
    <footer class="text-center text-muted small mt-4">
      Catálogo INIA Converge | Actualizado al {now_str}.
    </footer>
  </div>

  <script>
    const searchInput = document.getElementById('searchInput');
    const categoryFilter = document.getElementById('categoryFilter');
    const tableBody = document.querySelector('#solutionsTable tbody');
    const visibleCount = document.getElementById('visibleCount');
    const rows = Array.from(tableBody.querySelectorAll('tr'));

    function filterTable() {{
      const query = searchInput.value.toLowerCase().trim();
      const selectedCat = categoryFilter.value;
      let count = 0;

      rows.forEach(row => {{
        const text = row.textContent.toLowerCase();
        const cat = row.getAttribute('data-cat');
        
        const matchesQuery = !query || text.includes(query);
        const matchesCat = !selectedCat || cat === selectedCat;

        if (matchesQuery && matchesCat) {{
          row.style.display = '';
          count++;
        }} else {{
          row.style.display = 'none';
        }}
      }});

      visibleCount.textContent = count;
    }}

    searchInput.addEventListener('input', filterTable);
    categoryFilter.addEventListener('change', filterTable);
  </script>
</body>
</html>
"""

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print("PDF generado, Excel actualizado, CSV y HTML regenerados exitosamente!")
