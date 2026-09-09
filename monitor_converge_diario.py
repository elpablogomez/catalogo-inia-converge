#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitor y Trigger Diario de INIA Converge
=========================================
1. Revisa https://inia.uy/Converge para detectar si hay nuevas soluciones publicadas.
2. Si no hay cambios, finaliza rápidamente registrando la verificación.
3. Si detecta una nueva solución:
   a) Dispara la sincronización de correos (sync_emails.py).
   b) Busca la confirmación de Carlos Schiavi en emails.db para obtener la fecha real.
   c) Analiza los reportes PDF de la nueva solución (período de evaluación y reporte detallado).
   d) Regenera todos los archivos (soluciones_inia_converge.xlsx, .pdf, .csv, .json, index.html).
   e) Guarda un registro de la novedad en monitor_converge.log.
"""

import os
import sys
import json
import subprocess
from datetime import datetime
import requests
from bs4 import BeautifulSoup

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "soluciones_inia_converge.json")
LOG_FILE = os.path.join(BASE_DIR, "monitor_converge.log")
PORTAL_URL = "https://inia.uy/Converge"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

def log(msg: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def get_known_solutions():
    if not os.path.exists(JSON_FILE):
        return []
    try:
        with open(JSON_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return [item["solucion"].lower().strip() for item in data]
    except Exception as e:
        log(f"Error leyendo {JSON_FILE}: {e}")
        return []

def check_new_solutions():
    log("Iniciando verificación programada de INIA Converge...")
    known_solutions = get_known_solutions()
    log(f"Soluciones registradas actualmente en la base local: {len(known_solutions)}")

    try:
        resp = requests.get(PORTAL_URL, headers=HEADERS, timeout=20)
        resp.raise_for_status()
    except Exception as e:
        log(f"Error al conectar con {PORTAL_URL}: {e}")
        return

    # Usar el script unificado para inspeccionar y reconstruir si hay cambios
    # Ejecutamos una verificación completa invocando el generador
    log("Chequeando cambios en el catálogo...")
    try:
        res = subprocess.run(
            [sys.executable, os.path.join(BASE_DIR, "generar_catalogo_completo_con_pdf.py")],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        if res.returncode == 0:
            # Comprobar si aumentó la cantidad de soluciones
            with open(JSON_FILE, "r", encoding="utf-8") as f:
                new_data = json.load(f)
            if len(new_data) > len(known_solutions):
                nuevas = [item["solucion"] for item in new_data if item["solucion"].lower().strip() not in known_solutions]
                log(f"¡ALERTA! Se detectaron {len(nuevas)} soluciones nuevas: {', '.join(nuevas)}")
                
                # Ejecutar sincronización de correos para buscar el mail de Carlos Schiavi
                sync_script = os.path.join(BASE_DIR, "sync_emails.py")
                if os.path.exists(sync_script):
                    log("Sincronizando correos con Carlos Schiavi...")
                    subprocess.run([sys.executable, sync_script])
                    
                # Re-ejecutar generador con la base de correos actualizada
                subprocess.run([sys.executable, os.path.join(BASE_DIR, "generar_catalogo_completo_con_pdf.py")])
                
                # Copiar y desplegar a GitHub Pages
                scraper_dir = r"c:\Users\pmg19\OneDrive\Documentos\scraper-web-converge"
                if os.path.exists(scraper_dir):
                    log("Desplegando actualización a GitHub Pages...")
                    import shutil
                    for f in ["index.html", "soluciones_inia_converge.xlsx", "soluciones_inia_converge.pdf", "soluciones_inia_converge.csv", "soluciones_inia_converge.json"]:
                        src = os.path.join(BASE_DIR, f)
                        dst = os.path.join(scraper_dir, f)
                        if os.path.exists(src):
                            shutil.copy2(src, dst)
                    
                    commit_msg = f"Auto-actualización: {len(nuevas)} nueva(s) solución(es) ({', '.join(nuevas)})"
                    subprocess.run(["git", "add", "."], cwd=scraper_dir)
                    subprocess.run(["git", "commit", "-m", commit_msg], cwd=scraper_dir)
                    push_res = subprocess.run(["git", "push", "origin", "main"], cwd=scraper_dir, capture_output=True, text=True)
                    if push_res.returncode == 0:
                        log("¡GitHub Pages actualizado con éxito!")
                    else:
                        log(f"Aviso al hacer push a GitHub: {push_res.stderr[:200]}")

                log("Archivos regenerados con éxito (Excel, PDF, CSV, JSON, HTML).")
            else:
                log(f"Verificación finalizada sin novedades. Total soluciones: {len(new_data)}.")
        else:
            log(f"Error durante la ejecución del generador: {res.stderr[:200]}")
    except Exception as e:
        log(f"Excepción durante el monitoreo: {e}")

if __name__ == "__main__":
    check_new_solutions()
