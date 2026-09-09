#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scraper y Generador del Catálogo INIA Converge
=============================================
Integra la extracción web, cruce con convocatorias de postulación (Google Sheets),
fechas reales de publicación (confirmaciones de correo cschiavi), períodos técnicos
de verificación a campo y genera:
  - soluciones_inia_converge.xlsx (Excel profesional)
  - soluciones_inia_converge.pdf  (PDF horizontal A4 con ReportLab)
  - soluciones_inia_converge.csv  (CSV compatible con Excel/Sheets UTF-8 BOM)
  - soluciones_inia_converge.json (JSON estructurado)
  - index.html                   (Dashboard web interactivo compacto)
"""

import os
import sys
import subprocess

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    target_script = os.path.join(script_dir, "generar_catalogo_completo_con_pdf.py")
    if os.path.exists(target_script):
        subprocess.run([sys.executable, target_script], check=True)
    else:
        print("Ejecutando generación del catálogo...")

if __name__ == "__main__":
    main()
