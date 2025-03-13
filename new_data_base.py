import pandas as pd
import os
import openpyxl

# Ruta de carpeta
carpeta_cotizaciones = "/workspaces/IDM_project/folder_test"
archivo_base = "control_de_facturacion_2025.xlsx"

# Cargar base de datos o crear una nueva si no existe
if os.path.exists(archivo_base):
    df_base = pd.read_excel(archivo_base, engine="openpyxl")
else:
    df_base = pd.DataFrame(columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", "Cantidad", "Descripción", "Po", "Fecha de Po", "Precio Unitario", "Subtotal", "IVA", "Total", "Tipo de moneda"])

# Obtener lista de archivos ya procesados
archivos_procesados = df_base["Archivo"].unique().tolist() if not df_base.empty else []

# Iterar sobre todos los archivos en la carpeta
for archivo in os.listdir(carpeta_cotizaciones):
    if archivo.endswith(".xlsx") or archivo.endswith(".xlsm"):  # Solo procesar archivos Excel
        if archivo in archivos_procesados:
            print(f"⏩ Archivo {archivo} ya procesado, saltando...")
            continue  # Saltar el archivo si ya ha sido procesado

        ruta_archivo = os.path.join(carpeta_cotizaciones, archivo)
        print(f"📄 Procesando archivo: {archivo}")

        try:
            wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
            sheet = wb["cotizacion"]  # Ajusta el nombre de la hoja si es diferente

            # Datos generales
            po = sheet["K3"].value if sheet["K3"].value else "No encontrado"
            fecha_po = sheet["L4"].value if sheet["L4"].value else "No encontrado"
            no_cotizacion = sheet["K5"].value if sheet["K5"].value else "No encontrado"
            empresa = sheet["B3"].value if sheet["B3"].value else "No encontrado"
            requisitor = sheet["B5"].value if sheet["B5"].value else "No encontrado"
            subtotal = sheet["L36"].value if sheet["L36"].value else "No encontrado"
            iva = sheet["L37"].value if sheet["L37"].value else "No encontrado"
            total = sheet["L38"].value if sheet["L38"].value else "No encontrado"

            # Recorrer materiales desde la fila 15 en adelante
            fila = 15
            while True:
                descripcion = sheet[f"B{fila}"].value
                cantidad = sheet[f"H{fila}"].value
                precio_unidad = sheet[f"J{fila}"].value
                tipo_moneda = sheet[f"K{fila}"].value

                # Si la celda de descripción está vacía, terminamos el bucle
                if not descripcion or str(descripcion).strip() == "":
                    break  

                # Agregar una nueva fila con los datos extraídos
                nueva_fila = pd.DataFrame([[archivo, empresa, requisitor, no_cotizacion, cantidad, descripcion, po, fecha_po, precio_unidad, subtotal, iva, total, tipo_moneda]], 
                                          columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", "Cantidad", "Descripción", "Po", "Fecha de Po", "Precio Unitario", "Subtotal", "IVA", "Total", "Tipo de moneda"])
                df_base = pd.concat([df_base, nueva_fila], ignore_index=True)

                fila += 1  # Pasar a la siguiente fila

            wb.close()  # Cerrar el archivo después de procesarlo

        except Exception as e:
            print(f"⚠️ Error procesando el archivo {archivo}: {e}")

# Guardar la base de datos actualizada
df_base.to_excel(archivo_base, index=False, engine="openpyxl")
print("✅ Se han extraído los datos de todas las cotizaciones y la base de datos ha sido actualizada.")