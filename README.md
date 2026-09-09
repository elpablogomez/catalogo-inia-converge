# Scraper y Catálogo de Soluciones - INIA Converge

Este proyecto releva, consolida y estructura todas las soluciones tecnológicas y empresas AgTech publicadas en el portal de [INIA Converge](https://inia.uy/Converge), junto con los enlaces directos a todos sus materiales técnicos de verificación, informes detallados, testimonios y fechas de publicación inferidas.

---

## 📂 Archivos Generados

| Archivo | Descripción |
| :--- | :--- |
| `soluciones_inia_converge.xlsx` | **Planilla Excel profesional** con diseño visual, filtros automáticos, fechas de publicación, períodos de evaluación e **hipervínculos activos** directos a los reportes (PDF), informes detallados (PDF), testimonios (MP4/PDF), fichas gráficas y webs oficiales. |
| `index.html` | **Dashboard / Cuadro web interactivo** con buscador en tiempo real, filtros por categoría, insignias de fecha de publicación y botones de acceso directo a cada material. Se abre con un doble clic en cualquier navegador. |
| `soluciones_inia_converge.csv` | Archivo CSV codificado en UTF-8 con BOM (`utf-8-sig`) para compatibilidad directa con Microsoft Excel y Google Sheets. |
| `soluciones_inia_converge.json` | Datos en formato JSON estructurado con todos los campos y enlaces. |
| `scraper_converge.py` | Script automatizado en Python con control de tasa, reintentos anti-bloqueo e inferencia de fechas. |

---

## 📊 Estructura del Cuadro

El cuadro consolida los siguientes campos para cada una de las soluciones:

1. **Categoría**: Área productiva (*Vegetal Extensivo*, *Producción Animal*, *Vegetal Intensivo*).
2. **Subcategoría**: Rubro o tecnología (*Agricultura de Precisión*, *Monitoreo Fitosanitario*, *Ganadería*, *Riego*, etc.).
3. **Solución / Empresa**: Nombre comercial del producto o empresa AgTech.
4. **Fecha de Publicación**: Fecha inferida a partir de la cabecera HTTP `Last-Modified` del servidor oficial de INIA.
5. **Período de Evaluación**: Fechas en las que INIA realizó las pruebas a campo (extraídas mediante análisis de texto del documento técnico).
6. **Ficha INIA Converge**: Enlace directo a la ficha de la solución en `inia.uy`.
7. **Reporte de Verificación**: Informe técnico de verificación de INIA (PDF resumen).
8. **Informe Detallado**: Informe técnico extendido completo de 15 a 25 páginas (descubierto en los enlaces internos del PDF).
9. **Testimonio de Validación**: Evaluación en campo con usuarios reales (Video MP4 o PDF).
10. **Ficha Gráfica**: Infografía descriptiva del producto en alta resolución.
11. **Sitio Web Oficial**: Enlace al sitio web de la empresa proveedora.

---

## 🚀 Cómo actualizar los datos en el futuro

Para volver a correr el scraper y actualizar la base de datos ante nuevas publicaciones de INIA:

```bash
python scraper_converge.py
```
