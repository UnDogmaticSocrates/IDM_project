import pandas as pd
import openpyxl
import os
#ruta de carpeta
carpeta_cotizaciones = "/workspaces/IDM_project/cotizaciones"
archivo_base = "control_de_facturacion_2025.xlsx"

if os.path.exists(archivo_base):
    df_base = pd.read_excel(archivo_base, engine="openpyxl")
else:
    df_base = pd.DataFrame(columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", "Cantidad", "Descripción"])

# Iterar sobre todos los archivos en la carpeta
for archivo in os.listdir(carpeta_cotizaciones):
    if archivo.endswith(".xlsx") or archivo.endswith(".xlsm"):  # Solo procesar archivos Excel
        ruta_archivo = os.path.join(carpeta_cotizaciones, archivo)

        wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
        sheet = wb["Cotización"]  # Ajusta el nombre de la hoja si es diferente

        # 🔎 Extraer datos según las coordenadas que mencionaste
        empresa = sheet["H2"].value if sheet["H2"].value else "No encontrado"
        requisitor = sheet["I9"].value if sheet["I9"].value else "No encontrado"
        no_cotizacion = sheet["E9"].value if sheet["E9"].value else "No encontrado"
        cantidad = sheet["H14"].value if sheet["H14"].value else "No encontrado"
        descripcion = sheet["A14"].value if sheet["A14"].value else "No encontrado"

        nueva_fila = pd.DataFrame([[archivo, empresa, requisitor, no_cotizacion, cantidad, descripcion]], 
                                  columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", "Cantidad", "Descripción"])
        
         # Agregar a la base de datos
        df_base = pd.concat([df_base, nueva_fila], ignore_index=True)

# Guardar la base de datos actualizada
df_base.to_excel(archivo_base, index=False, engine="openpyxl")
print("✅ Se han extraído los datos de todas las cotizaciones y la base de datos ha sido actualizada.")