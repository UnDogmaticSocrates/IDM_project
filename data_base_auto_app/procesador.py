import pandas as pd
import os
import openpyxl

def procesar_archivos(carpeta_cotizaciones, archivo_base):
    # Cargar base de datos o crear nueva
    if os.path.exists(archivo_base):
        df_base = pd.read_excel(archivo_base, engine="openpyxl")
    else:
        df_base = pd.DataFrame(columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", 
                                        "Cantidad", "Descripción", "Po", "Fecha de Po", 
                                        "Precio Unitario", "Subtotal", "IVA", "Total", "Tipo de moneda"])

    # Iterar sobre archivos en la carpeta
    for archivo in os.listdir(carpeta_cotizaciones):
        if archivo.endswith(".xlsx") or archivo.endswith(".xlsm"):
            ruta_archivo = os.path.join(carpeta_cotizaciones, archivo)
            print(f"📂 Procesando: {archivo}")
            try:
                wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
                sheet = wb["cotizacion"]

                # Leer datos
                po = sheet["K3"].value or "No encontrado"
                fecha_po = sheet["L4"].value or "No encontrado"
                no_cotizacion = sheet["K5"].value or "No encontrado"
                empresa = sheet["B3"].value or "No encontrado"
                requisitor = sheet["B5"].value or "No encontrado"
                subtotal = sheet["L36"].value or "No encontrado"
                iva = sheet["L37"].value or "No encontrado"
                total = sheet["L38"].value or "No encontrado"

                # Recorrer materiales
                fila = 15
                while True:
                    descripcion = sheet[f"B{fila}"].value
                    cantidad = sheet[f"H{fila}"].value
                    precio_unidad = sheet[f"J{fila}"].value
                    tipo_moneda = sheet[f"K{fila}"].value

                    if not descripcion or str(descripcion).strip() == "":
                        break  

                    nueva_fila = pd.DataFrame([[archivo, empresa, requisitor, no_cotizacion, cantidad, 
                                                descripcion, po, fecha_po, precio_unidad, 
                                                subtotal, iva, total, tipo_moneda]], 
                                              columns=df_base.columns)
                    if not ((df_base["Archivo"] == archivo) & (df_base["No. Cotización"] == no_cotizacion) &  (df_base["Descripción"] == descripcion) & (df_base["Cantidad"] == cantidad)).any():df_base = pd.concat([df_base, nueva_fila], ignore_index=True)
                    fila += 1
                wb.close()
            except Exception as e:
                print(f"⚠️ Error en {archivo}: {e}")

    # Guardar archivo actualizado
    df_base.to_excel(archivo_base, index=False, engine="openpyxl")
    print("✅ Base de datos actualizada.")