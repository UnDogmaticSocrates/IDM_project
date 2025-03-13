import pandas as pd
import openpyxl
import os
#ruta de carpeta
carpeta_cotizaciones = "/workspaces/IDM_project/folder_test"
archivo_base = "control_de_facturacion_2025.xlsx"

if os.path.exists(archivo_base):
    df_base = pd.read_excel(archivo_base, engine="openpyxl")
else:
    df_base = pd.DataFrame(columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", "Cantidad", "Descripción", "Po", "Fecha de Po", "Precio Unitario", "Subtotal", "IVA", "Total", "Tipo de moneda"])

# Iterar sobre todos los archivos en la carpeta
for archivo in os.listdir(carpeta_cotizaciones):
    if archivo.endswith(".xlsx") or archivo.endswith(".xlsm"):  # Solo procesar archivos Excel
        ruta_archivo = os.path.join(carpeta_cotizaciones, archivo)

        wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
        sheet = wb["Cotizacion"]  # Ajusta el nombre de la hoja si es diferente

        # 🔎 Extraer datos según las coordenadas que mencionaste
        po = sheet["K3"].value if sheet["K3"].value else "No encontrado"
        fecha_po = sheet["K4"].value if sheet["K4"].value else "No encontrado"
        no_cotizacion = sheet["K5"].value if sheet["K5"].value else "No encontrado"
        empresa = sheet["B3"].value if sheet["B3"].value else "No encontrado"
        requisitor = sheet["B5"].value if sheet["B5"].value else "No encontrado"
        descripcion = sheet["B15"].value if sheet["B15"].value else "No encontrado"
        cantidad = sheet["H15"].value if sheet["H15"].value else "No encontrado"
        tipo_moneda = sheet["K15"].value if sheet["K15"].value else "No encontrado"
        precio_unidad = sheet["J15"].value if sheet["J15"].value else "No encontrado"
        subtotal = sheet["L36"].value if sheet["L36"].value else "No encontrado"
        iva = sheet["L37"].value if sheet["L37"].value else "No encontrado"
        total = sheet["L38"].value if sheet["L38"].value else "No encontrado"
    
        nueva_fila = pd.DataFrame([[archivo, empresa, requisitor, no_cotizacion, cantidad, descripcion, po, fecha_po, precio_unidad, subtotal, iva, total, tipo_moneda]], 
                                  columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", "Cantidad", "Descripción", "Po", "Fecha de Po", "Precio Unitario", "Subtotal", "IVA", "Total", "Tipo de moneda"])
        
         # Agregar a la base de datos
        df_base = pd.concat([df_base, nueva_fila], ignore_index=True)

# Guardar la base de datos actualizada
df_base.to_excel(archivo_base, index=False, engine="openpyxl")
print("✅ Se han extraído los datos de todas las cotizaciones y la base de datos ha sido actualizada.")