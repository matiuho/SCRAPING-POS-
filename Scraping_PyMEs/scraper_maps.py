import time
import re
import pandas as pd
from playwright.sync_api import sync_playwright

# Lista de palabras clave para descartar grandes cadenas, supermercados y malls
CADENAS_Y_MALLS = [
    "falabella", "paris", "ripley", "sodimac", "easy", "lider", "walmart", 
    "santa isabel", "unimarc", "alvi", "jumbo", "express", "tottus", "oxxo", 
    "ok market", "mall", "plaza", "costanera", "outlet", "copec", "shell", 
    "penta", "spid", "cruz verde", "ahumada", "salcobrand"
]

def es_cadena_grande(nombre):
    nombre_lower = nombre.lower()
    return any(cadena in nombre_lower for cadena in CADENAS_Y_MALLS)

def limpiar_texto(texto):
    if not texto:
        return ""
    # Eliminar emojis y caracteres no imprimibles o raros
    texto_limpio = re.sub(r'[^\w\s\.,#\-\(\)/áéíóúÁÉÍÓÚñÑa-zA-Z0-9]', '', texto)
    # Limpiar espacios múltiples o saltos de línea
    return re.sub(r'\s+', ' ', texto_limpio).strip()

def extraer_pymes_google_maps(busqueda, comuna, max_resultados=10):
    print(f"🚀 Iniciando búsqueda PyME: '{busqueda}' en {comuna}...")
    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(locale="es-CL")
        page = context.new_page()

        url = f"https://www.google.com/maps/search/{busqueda.replace(' ', '+')}+{comuna.replace(' ', '+')}"
        page.goto(url)
        page.wait_for_timeout(4000)

        # Scroll para cargar opciones
        try:
            scrollable = page.locator('div[role="feed"]')
            for _ in range(4):
                scrollable.evaluate('el => el.scrollTop += 1200')
                page.wait_for_timeout(1000)
        except Exception:
            pass

        links = page.locator('a[href*="/maps/place/"]').all()
        print(f"🔍 Locales detectados: {len(links)}")

        urls_visitadas = set()

        for link in links:
            href = link.get_attribute("href")
            if not href or href in urls_visitadas:
                continue
            urls_visitadas.add(href)

            nombre_local = link.get_attribute("aria-label")
            if not nombre_local or nombre_local.strip() == "":
                continue

            # FILTRO 1: Si es una gran cadena o supermercado, se ignora
            if es_cadena_grande(nombre_local):
                print(f"⛔ Descartado por Gran Cadena/Mall: {nombre_local}")
                continue

            if len(resultados) >= max_resultados:
                break

            try:
                link.click()
                page.wait_for_timeout(2000)

                # 1. Extraer Teléfono
                telefono = "Sin Teléfono"
                tel_locator = page.locator('button[data-tooltip*="teléfono"], button[aria-label*="Teléfono"], button[data-item-id*="phone"]')
                
                if tel_locator.count() > 0:
                    telefono = tel_locator.first.inner_text().strip()
                else:
                    body_text = page.locator('div[role="main"]').inner_text()
                    match = re.search(r'(\+?56\s?9?\s?\d{4}\s?\d{4}|\b2\s?\d{4}\s?\d{4}\b)', body_text)
                    if match:
                        telefono = match.group(0)

                # 2. Extraer Dirección
                direccion = "Sin Dirección"
                dir_locator = page.locator('button[data-item-id="address"], button[aria-label*="Dirección"]')
                if dir_locator.count() > 0:
                    direccion = dir_locator.first.inner_text().strip()

                # Limpieza estricta de emojis e íconos en textos
                nombre_clean = limpiar_texto(nombre_local)
                tel_clean = re.sub(r'[^\d\+\s\(\)\-]', '', telefono).strip()
                dir_clean = limpiar_texto(direccion)

                if tel_clean and len(tel_clean) > 5:
                    resultados.append([
                        nombre_clean,
                        busqueda.capitalize(),
                        tel_clean,
                        comuna.capitalize(),
                        dir_clean,
                        "Sí",  # Indicador limpio de POS
                        "Pendiente",
                        ""     # Columna 'Observaciones' totalmente VACÍA
                    ])
                    print(f"✅ [{len(resultados)}] PyME: {nombre_clean} | Tel: {tel_clean}")

            except Exception as e:
                continue

        browser.close()

    return resultados

# --- CONFIGURACIÓN DE TU BÚSQUEDA ---
rubros_a_buscar = ["Minimarket", "Almacen", "Botilleria", "Cafeteria"]
comuna_objetivo = "Las Condes"

todas_las_pymes = []

for rubro in rubros_a_buscar:
    datos = extraer_pymes_google_maps(rubro, comuna_objetivo, max_resultados=10)
    todas_las_pymes.extend(datos)

# --- GUARDAR EN EL EXCEL ---
columnas = [
    "Nombre_Empresa", 
    "Rubro", 
    "Telefono", 
    "Comuna", 
    "Direccion", 
    "Tiene_POS_Estimado",
    "Estado_Llamada", 
    "Observaciones"
]

df = pd.DataFrame(todas_las_pymes, columns=columnas)

# Filtrar duplicados
df.drop_duplicates(subset=["Nombre_Empresa", "Telefono"], inplace=True)

nombre_excel = "Base_PyMEs_POS.xlsx"

try:
    df.to_excel(nombre_excel, index=False)
    print(f"\n🎉 ¡Proceso finalizado! Se guardaron {len(df)} PyMEs reales (sin grandes cadenas ni emojis) en '{nombre_excel}'.")
except PermissionError:
    print(f"\n⚠️ ERROR: Por favor CIERRA el archivo '{nombre_excel}' en Microsoft Excel y vuelve a ejecutar el programa.")