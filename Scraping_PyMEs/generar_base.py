import pandas as pd

# 1. Definir la estructura de la base de datos de PyMEs
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

# 2. Ejemplo de datos iniciales extraídos
datos_ejemplo = [
    ["Minimarket Don Pedro", "Minimarket", "+56912345678", "Providencia", "Av. Italia 1234", "Sí", "Pendiente", ""],
    ["Botillería El Sol", "Botillería", "+56987654321", "Las Condes", "Av. Cristóbal Colón 3770", "Sí", "Pendiente", ""],
    ["Café Central", "Cafetería", "+56955554444", "Conchalí", "Av. Independencia 567", "Sí", "Pendiente", ""]
]

# 3. Crear el DataFrame de Pandas
df = pd.DataFrame(datos_ejemplo, columns=columnas)

# 4. Exportar a un archivo Excel local
nombre_archivo = "Base_PyMEs_POS.xlsx"
df.to_excel(nombre_archivo, index=False)

print(f"✅ Archivo '{nombre_archivo}' generado con éxito.")