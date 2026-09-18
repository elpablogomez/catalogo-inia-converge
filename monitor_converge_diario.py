#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitor Diario Automatizado de INIA Converge
============================================
1. Rastrea el portal https://inia.uy/Converge y todas sus categorías/subcategorías.
2. Compara las soluciones online con las registradas localmente en el catálogo.
3. Si detecta una o más soluciones nuevas:
   - Extrae enlaces a reporte PDF, informe técnico detallado, testimonio (video/PDF), ficha gráfica y web oficial.
   - Infiere la fecha real de publicación mediante la cabecera HTTP 'Last-Modified'.
   - Extrae el período de evaluación analizando el contenido del reporte PDF mediante Gemini 2.5 Flash (con fallback a regex).
   - Agrega las soluciones a la base y regenera todos los artefactos (Excel, PDF, CSV, JSON, index.html).
4. Compatible para ejecución en GitHub Actions y de forma local en Windows/Linux.
"""

import os
import sys
import json
import re
import unicodedata
import subprocess
from datetime import datetime
from urllib.parse import urljoin, unquote, urlparse
from email.utils import parsedate_to_datetime
import requests
from bs4 import BeautifulSoup

try:
    import fitz  # PyMuPDF para analizar texto en PDF
except ImportError:
    fitz = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, "soluciones_inia_converge.json")
BASE_JSON_FILE = os.path.join(BASE_DIR, "soluciones_converge_fechas_completas.json")
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

CATEGORY_MAP = {
    "vegetal-extensivo": "Vegetal Extensivo",
    "vegetal-intensivo": "Vegetal Intensivo",
    "producción-animal": "Producción Animal",
    "produccion-animal": "Producción Animal",
    "producci%c3%b3n-animal": "Producción Animal",
}

SUBCATEGORY_MAP = {
    "agricultura-de-precisión": "Agricultura De Precisión",
    "agricultura-de-precision": "Agricultura De Precisión",
    "agricultura-de-precisi%c3%b3n": "Agricultura De Precisión",
    "arroz": "Arroz",
    "control-de-aplicaciones": "Control De Aplicaciones",
    "control-de-aplicación": "Control De Aplicación",
    "control-de-aplicacion": "Control De Aplicación",
    "control-de-aplicaci%c3%b3n": "Control De Aplicación",
    "forestal": "Forestal",
    "monitoreo-fitosanitario": "Monitoreo Fitosanitario",
    "avicola": "Avicola",
    "avícola": "Avicola",
    "av%c3%adcola": "Avicola",
    "ganadería": "Ganadería",
    "ganaderia": "Ganadería",
    "ganader%c3%ada": "Ganadería",
    "lechería": "Lechería",
    "lecheria": "Lechería",
    "lecher%c3%ada": "Lechería",
    "porcinos": "Porcinos",
    "gestión-de-operaciones": "Gestión De Operaciones",
    "gestion-de-operaciones": "Gestión De Operaciones",
    "gesti%c3%b3n-de-operaciones": "Gestión De Operaciones",
    "hidroponia": "Hidroponia",
    "riego": "Riego",
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

def normalize_name(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode()
    return re.sub(r"[\s\-_]+", "", s).lower()

def get_known_data():
    target = BASE_JSON_FILE if os.path.exists(BASE_JSON_FILE) else JSON_FILE
    if not os.path.exists(target):
        return []
    try:
        with open(target, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log(f"Error leyendo base conocida {target}: {e}")
        return []

def format_category(raw: str) -> str:
    cleaned = unquote(raw).lower().strip().replace(" ", "-")
    return CATEGORY_MAP.get(cleaned, unquote(raw).replace("-", " ").title())

def format_subcategory(raw: str) -> str:
    cleaned = unquote(raw).lower().strip().replace(" ", "-")
    return SUBCATEGORY_MAP.get(cleaned, unquote(raw).replace("-", " ").title())

def get_gemini_api_key() -> str:
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not key:
        key_path = os.path.expanduser(r"~\.gemini\credentials\gemini_api_key.txt")
        if os.path.exists(key_path):
            try:
                with open(key_path, "r", encoding="utf-8") as f:
                    key = f.read().strip()
            except Exception:
                pass
    return key

def extract_period_with_gemini(pdf_text: str, session: requests.Session) -> str:
    api_key = get_gemini_api_key()
    if not api_key or not pdf_text or len(pdf_text.strip()) < 50:
        return ""

    prompt = (
        "Analiza el siguiente texto extraído del reporte técnico de verificación de INIA Converge "
        "y extrae de forma concisa el 'Período de Verificación Técnica' (cuándo y dónde se realizaron las "
        "pruebas a campo, por ejemplo: 'Octubre 2025 (INIA La Estanzuela)', 'Setiembre 2024 a Febrero 2025', "
        "'Junio 2024 a Junio 2025 (INIA Las Brujas)'). "
        "Responde ÚNICAMENTE con una línea con el período formateado, sin asteriscos, sin comillas y sin texto adicional.\n\n"
        f"Texto del reporte:\n{pdf_text[:4000]}"
    )

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={api_key}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1
        }
    }

    try:
        resp = session.post(url, json=payload, timeout=25)
        if resp.status_code == 200:
            data = resp.json()
            candidate = data["candidates"][0]["content"]["parts"][0]["text"].strip()
            clean_period = candidate.strip().strip('"').strip("'")
            if clean_period and len(clean_period) < 100:
                log(f"  [Gemini 2.5 Flash] Período de verificación extraído: {clean_period}")
                return clean_period
        else:
            log(f"  [Gemini] Respuesta HTTP {resp.status_code}: {resp.text[:150]}")
    except Exception as e:
        log(f"  [Gemini] Error al consultar API de Gemini: {e}")

    return ""

def crawl_portal(session: requests.Session):
    log("Rastreando categorías y soluciones en INIA Converge...")
    r = session.get(PORTAL_URL, timeout=25)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    # 1. Encontrar categorías principales
    cat_urls = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if any(c in href.lower() for c in ["producci", "vegetal-extensivo", "vegetal-intensivo"]):
            url = urljoin(PORTAL_URL, href)
            if url not in cat_urls:
                cat_urls.append(url)

    # 2. Encontrar subcategorías
    subcat_urls = []
    for cat in cat_urls:
        try:
            r_cat = session.get(cat, timeout=25)
            soup_cat = BeautifulSoup(r_cat.text, "html.parser")
            for a in soup_cat.find_all("a", href=True):
                href = a["href"]
                if "/converge/" in href.lower():
                    url = urljoin(cat, href)
                    parts = [p for p in urlparse(url).path.strip("/").split("/") if p]
                    if len(parts) == 3 and url not in subcat_urls and url not in cat_urls:
                        subcat_urls.append(url)
        except Exception as e:
            log(f"Aviso al acceder a categoría {cat}: {e}")

    # 3. Encontrar páginas de soluciones individuales
    solutions = {}
    for sub in subcat_urls:
        try:
            r_sub = session.get(sub, timeout=25)
            soup_sub = BeautifulSoup(r_sub.text, "html.parser")
            for a in soup_sub.find_all("a", href=True):
                href = a["href"]
                if "/converge/" in href.lower():
                    url = urljoin(sub, href)
                    parts = [p for p in urlparse(url).path.strip("/").split("/") if p]
                    if len(parts) == 4:
                        sol_name = unquote(parts[3])
                        raw_cat = parts[1]
                        raw_subcat = parts[2]
                        url_key = url.lower().rstrip("/")
                        if url_key not in solutions:
                            solutions[url_key] = {
                                "name": sol_name,
                                "url": url,
                                "raw_cat": raw_cat,
                                "raw_subcat": raw_subcat
                            }
        except Exception as e:
            log(f"Aviso al acceder a subcategoría {sub}: {e}")

    return solutions

def extract_solution_details(session: requests.Session, sol_info: dict) -> dict:
    url = sol_info["url"]
    name = sol_info["name"]
    log(f"Extrayendo información detallada para nueva solución: {name} ({url})")

    resp = session.get(url, timeout=25)
    soup = BeautifulSoup(resp.text, "html.parser")

    item = {
        "categoria": format_category(sol_info["raw_cat"]),
        "subcategoria": format_subcategory(sol_info["raw_subcat"]),
        "solucion": name,
        "fecha_publicacion": "-",
        "periodo_evaluacion": "-",
        "url_inia": url,
        "reporte_verificacion": "-",
        "informe_detallado": "-",
        "testimonio_validacion": "-",
        "tipo_testimonio": "-",
        "ficha_grafica": "-",
        "pagina_web": "-",
        "fecha_publicacion_web": "-",
        "evidencia_correo": "-",
        "fecha_testimonio": "-",
        "convocatoria_postulacion": "-"
    }

    # Buscar enlaces a recursos
    for a in soup.find_all("a", href=True):
        href = urljoin(url, a["href"])
        text = a.get_text(strip=True).lower()
        href_lower = href.lower()
        file_name = os.path.basename(urlparse(href).path).lower()

        if any(ign in href_lower for ign in ["facebook", "twitter", "youtube.com", "instagram", "linkedin"]):
            continue

        if ("testimonio" in file_name or "video" in file_name or "testimonio" in text) and any(ext in file_name for ext in [".pdf", ".mp4", "youtu"]):
            item["testimonio_validacion"] = href
            item["tipo_testimonio"] = "Video" if (".mp4" in file_name or "youtu" in href_lower) else "PDF"
        elif ("informe" in file_name or "detallado" in text) and file_name.endswith(".pdf"):
            item["informe_detallado"] = href
        elif ("reporte" in file_name or "reporte" in text) and file_name.endswith(".pdf"):
            item["reporte_verificacion"] = href
        elif "pagina web" in text or "sitio web" in text or (name.lower() in href_lower and "inia.uy" not in href_lower):
            item["pagina_web"] = href

    # Buscar ficha gráfica en imágenes
    for img in soup.find_all("img"):
        src = img.get("src", "")
        if any(term in src.lower() for term in ["ficha", name.lower(), "paginasproductos"]):
            item["ficha_grafica"] = urljoin(url, src)
            break

    # Si no se encontró informe detallado explícito, probar patrón estándar en servidor
    if item["informe_detallado"] == "-":
        for cand_name in [f"Informe_{name}.pdf", f"Informe%20{name}.pdf"]:
            candidate = f"https://inia.uy/sites/default/files/gras/contenidogras/Reportes/Converge/PaginasProductos/{name}/{cand_name}"
            try:
                head_cand = session.head(candidate, timeout=10)
                if head_cand.status_code == 200:
                    item["informe_detallado"] = candidate
                    break
            except Exception:
                pass

    # Inferencia de fechas mediante HTTP Last-Modified
    if item["reporte_verificacion"] != "-":
        try:
            r_rep = session.head(item["reporte_verificacion"], timeout=10)
            if r_rep.status_code == 200 and "Last-Modified" in r_rep.headers:
                dt = parsedate_to_datetime(r_rep.headers["Last-Modified"])
                item["fecha_publicacion_web"] = dt.strftime("%Y-%m-%d")
                item["fecha_publicacion"] = dt.strftime("%Y-%m-%d")
        except Exception as e:
            log(f"Error obteniendo fecha reporte: {e}")

    if item["testimonio_validacion"] != "-":
        try:
            r_test = session.head(item["testimonio_validacion"], timeout=10)
            if r_test.status_code == 200 and "Last-Modified" in r_test.headers:
                dt = parsedate_to_datetime(r_test.headers["Last-Modified"])
                item["fecha_testimonio"] = dt.strftime("%Y-%m-%d")
        except Exception as e:
            log(f"Error obteniendo fecha testimonio: {e}")

    # Extraer período de evaluación desde el texto del PDF
    if fitz and item["reporte_verificacion"] != "-":
        try:
            r_pdf = session.get(item["reporte_verificacion"], timeout=20)
            if r_pdf.status_code == 200:
                doc = fitz.open(stream=r_pdf.content, filetype="pdf")
                txt = " ".join(p.get_text() for p in doc)

                # 1. Intentar con Gemini si la API key está disponible
                gemini_period = extract_period_with_gemini(txt, session)
                if gemini_period:
                    item["periodo_evaluacion"] = gemini_period
                else:
                    # 2. Fallback a expresiones regulares
                    m = re.search(r"(?:entre|durante|desde|en|con siembra realizada el)\s+([0-9a-záéíóúñ\s\–\-]+(?:de\s+)?202[0-9])", txt, re.IGNORECASE)
                    if m:
                        item["periodo_evaluacion"] = m.group(0).strip().capitalize()
                
                # Revisar enlaces internos del PDF si aún no hay informe detallado
                if item["informe_detallado"] == "-":
                    for p in doc:
                        for l in p.get_links():
                            uri = l.get("uri", "")
                            if "informe" in uri.lower() and uri.lower().endswith(".pdf"):
                                if name.lower() in uri.lower():
                                    item["informe_detallado"] = uri
                                    break
        except Exception as e:
            log(f"Aviso al analizar texto PDF: {e}")

    return item

def check_new_solutions():
    log("=== Iniciando Verificación Automática de INIA Converge ===")
    known_data = get_known_data()
    known_names = set(normalize_name(d["solucion"]) for d in known_data)
    known_urls = set(d["url_inia"].lower().rstrip("/") for d in known_data)
    log(f"Soluciones registradas localmente en la base: {len(known_data)}")

    session = requests.Session()
    session.headers.update(HEADERS)

    try:
        online_solutions = crawl_portal(session)
    except Exception as e:
        log(f"ERROR: Fallo al rastrear portal INIA Converge: {e}")
        return

    log(f"Soluciones detectadas online en portal INIA: {len(online_solutions)}")

    # Detectar cuáles no están en known_data
    nuevas = []
    for url_key, info in online_solutions.items():
        if normalize_name(info["name"]) not in known_names and url_key not in known_urls:
            nuevas.append(info)

    if not nuevas:
        log(f"Verificación finalizada: no hay soluciones nuevas en la web (Total: {len(known_data)}).")
        subprocess.run([sys.executable, os.path.join(BASE_DIR, "generar_catalogo_completo_con_pdf.py")], cwd=BASE_DIR)
        return

    log(f"¡ALERTA! Se detectaron {len(nuevas)} solución(es) NUEVA(S) en el portal:")
    for n in nuevas:
        log(f"  -> {n['name']} en {n['url']}")

    # Extraer y agregar cada nueva solución
    for n in nuevas:
        new_sol = extract_solution_details(session, n)
        known_data.append(new_sol)

    # Guardar en BASE_JSON_FILE y JSON_FILE
    with open(BASE_JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(known_data, f, ensure_ascii=False, indent=2)

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(known_data, f, ensure_ascii=False, indent=2)

    log(f"Base de datos actualizada con éxito ({len(known_data)} soluciones totales).")

    # Ejecutar regeneración de Excel, PDF, CSV, JSON e index.html
    log("Regenerando todos los artefactos de catálogo (Excel, PDF, CSV, JSON, index.html)...")
    res = subprocess.run([sys.executable, os.path.join(BASE_DIR, "generar_catalogo_completo_con_pdf.py")], cwd=BASE_DIR, capture_output=True, text=True, encoding="utf-8")
    if res.returncode == 0:
        log("¡Todos los artefactos del catálogo se regeneraron exitosamente!")
    else:
        log(f"Aviso durante la generación del catálogo: {res.stderr[:300]}")

if __name__ == "__main__":
    check_new_solutions()
