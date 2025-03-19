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
        if archivo.endswith((".xlsx", ".xlsm")):
            ruta_archivo = os.path.join(carpeta_cotizaciones, archivo)
            print(f"📂 Procesando: {archivo}")

            try:
                    wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
                    if "cotizacion" not in wb.sheetnames:
                        print(f"⚠️ {archivo} no contiene la hoja 'cotizacion'.")
                        wb.close()
                        continue

                    sheet = wb["cotizacion"]

                    # Leer datos generales
                    po = sheet["B6"].value or "No encontrado"
                    fecha_po = sheet["A6"].value or "No encontrado"
                    no_cotizacion = sheet["G6"].value or "No encontrado"
                    empresa = sheet["C6"].value or "No encontrado"
                    requisitor = sheet["H6"].value or "No encontrado"
                    subtotal = sheet["H33"].value or "No encontrado"
                    iva = sheet["H36"].value or "No encontrado"
                    total = sheet["H37"].value or "No encontrado"

                    # Recorrer materiales
                    fila = 9
                    nuevas_filas = []

                    while fila <= 32:  # Se mantiene el límite de filas
                        descripcion = sheet[f"B{fila}"].value
                        cantidad = sheet[f"H{fila}"].value
                        precio_unidad = sheet[f"J{fila}"].value
                        tipo_moneda = sheet[f"K{fila}"].value

                        # Si la descripción está vacía, se ignora, pero no se detiene
                        if descripcion:
                            nuevas_filas.append([archivo, empresa, requisitor, no_cotizacion, cantidad, 
                                                 descripcion, po, fecha_po, precio_unidad, 
                                                 subtotal, iva, total, tipo_moneda])
                        
                        fila += 1  # Seguir iterando hasta la fila 32

                    # Crear DataFrame con nuevas filas
                    df_nuevo = pd.DataFrame(nuevas_filas, columns=df_base.columns)

                    # Evitar duplicados antes de concatenar
                    df_base = pd.concat([df_base, df_nuevo]).drop_duplicates(
                        subset=["Archivo", "No. Cotización", "Descripción", "Cantidad"],
                        keep="first"
                    ).reset_index(drop=True)

            except Exception as e:
                print(f"⚠️ Error en {archivo}: {e}")

    # Guardar archivo actualizado
    df_base.to_excel(archivo_base, index=False, engine="openpyxl")
    print("✅ Base de datos actualizada.")