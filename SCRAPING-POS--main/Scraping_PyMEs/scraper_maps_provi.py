import time
import re
import pandas as pd
from playwright.sync_api import sync_playwright

CADENAS_Y_MALLS = [
    # Malls, retail y supermercados
    "falabella", "paris", "ripley", "sodimac", "easy", "lider", "walmart", 
    "santa isabel", "unimarc", "alvi", "jumbo", "express", "tottus", "oxxo", 
    "ok market", "mall", "plaza", "costanera", "outlet", "copec", "shell", 
    "penta", "spid", "cruz verde", "ahumada", "salcobrand",
    # Cadenas ferreteras, automotrices y salud
    "imperial", "construmart", "chilemat", "mts", "autoplanet", "redsalud", 
    "integramedica", "clinica alemana", "clinica las condes", "clinica indisa", 
    "bupa", "uno salud"
]

def es_cadena_grande(nombre):
    nombre_lower = nombre.lower()
    return any(cadena in nombre_lower for cadena in CADENAS_Y_MALLS)

def limpiar_texto(texto):
    if not texto:
        return ""
    texto_limpio = re.sub(r'[^\w\s\.,#\-\(\)/áéíóúÁÉÍÓÚñÑa-zA-Z0-9]', '', texto)
    return re.sub(r'\s+', ' ', texto_limpio).strip()

def es_telefono_fijo_chile(telefono_str):
    # Extraer dígitos
    solo_digitos = re.sub(r'\D', '', telefono_str)
    
    if solo_digitos.startswith('56'):
        solo_digitos = solo_digitos[2:]
        
    # Descartar números de red fija (comienzan con 2)
    return solo_digitos.startswith('2')

def extraer_pymes_google_maps(busqueda, comuna, max_resultados=10):
    print(f"\n🚀 Iniciando búsqueda: '{busqueda}' en {comuna}...")
    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(locale="es-CL")
        page = context.new_page()

        query_url = f"https://www.google.com/maps/search/{busqueda.replace(' ', '+')}+{comuna.replace(' ', '+')}"
        page.goto(query_url)
        page.wait_for_timeout(5000)

        # Hacer scroll para cargar resultados en la lista lateral
        for _ in range(6):
            try:
                feed = page.locator('div[role="feed"]')
                if feed.count() > 0:
                    feed.first.evaluate('el => el.scrollTop += 1500')
                else:
                    page.mouse.wheel(0, 1500)
            except Exception:
                page.mouse.wheel(0, 1500)
            page.wait_for_timeout(1500)

        # Localizar tarjetas de resultados
        links = page.locator('a[href*="/maps/place/"]').all()
        print(f"🔍 Locales detectados en pantalla: {len(links)}")

        urls_visitadas = set()

        for link in links:
            href = link.get_attribute("href")
            if not href or href in urls_visitadas:
                continue
            urls_visitadas.add(href)

            nombre_local = link.get_attribute("aria-label")
            if not nombre_local or nombre_local.strip() == "":
                continue

            if es_cadena_grande(nombre_local):
                print(f"⛔ Descartado (Gran Cadena): {nombre_local}")
                continue

            if len(resultados) >= max_resultados:
                break

            try:
                link.click()
                page.wait_for_timeout(2500)

                # Extraer Teléfono
                telefono = "Sin Teléfono"
                tel_locator = page.locator('button[data-tooltip*="teléfono"], button[aria-label*="Teléfono"], button[data-item-id*="phone"]')
                
                if tel_locator.count() > 0:
                    telefono = tel_locator.first.inner_text().strip()
                else:
                    main_container = page.locator('div[role="main"]')
                    if main_container.count() > 0:
                        body_text = main_container.inner_text()
                        match = re.search(r'(\+?56\s?9\s?\d{4}\s?\d{4}|\b9\s?\d{4}\s?\d{4}\b)', body_text)
                        if match:
                            telefono = match.group(0)

                # Validar teléfono celular
                if telefono == "Sin Teléfono":
                    print(f"⚠️ Omitido (Sin número visible): {nombre_local}")
                    continue

                if es_telefono_fijo_chile(telefono):
                    print(f"📞 Omitido (Teléfono fijo descartado): {nombre_local} -> {telefono}")
                    continue

                # Extraer Dirección
                direccion = "Sin Dirección"
                dir_locator = page.locator('button[data-item-id="address"], button[aria-label*="Dirección"]')
                if dir_locator.count() > 0:
                    direccion = dir_locator.first.inner_text().strip()

                nombre_clean = limpiar_texto(nombre_local)
                tel_clean = re.sub(r'[^\d\+\s\(\)\-]', '', telefono).strip()
                dir_clean = limpiar_texto(direccion)

                if tel_clean and len(tel_clean) > 6:
                    resultados.append([
                        nombre_clean,
                        busqueda.capitalize(),
                        tel_clean,
                        comuna.capitalize(),
                        dir_clean,
                        "Sí",
                        "Pendiente",
                        ""
                    ])
                    print(f"✅ [{len(resultados)}] Guardado: {nombre_clean} | Celular: {tel_clean}")

            except Exception as e:
                print(f"❌ Error al procesar local: {e}")
                continue

        browser.close()

    return resultados

if __name__ == "__main__":
    # Nuevos rubros enfocados en ticket promedio > $20.000
    rubros = [
        "Ferreteria",
        "Materiales de construccion",
        "Taller mecanico",
        "Veterinaria",
        "Boutique de ropa",
        "Peluqueria",
        "Barberia",
        "Centro de estetica",
        "Distribuidora",
        "Clinica dental"
    ]
    comuna = "Providencia"

    todas_las_pymes = []

    for rubro in rubros:
        datos = extraer_pymes_google_maps(rubro, comuna, max_resultados=10)
        todas_las_pymes.extend(datos)

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

    if not df.empty:
        df.drop_duplicates(subset=["Nombre_Empresa", "Telefono"], inplace=True)
        archivo_salida = "Base_PyMEs_POS_Providencia.xlsx"
        try:
            df.to_excel(archivo_salida, index=False)
            print(f"\n🎉 ¡Proceso completado! Se guardaron {len(df)} registros en '{archivo_salida}'.")
        except PermissionError:
            print(f"\n⚠️ Cierra el archivo '{archivo_salida}' en Excel y vuelve a ejecutar.")
    else:
        print("\n⚠️ No se encontraron locales con números celulares bajo los criterios establecidos.")