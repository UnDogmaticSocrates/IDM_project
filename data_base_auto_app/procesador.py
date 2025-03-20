import pandas as pd
import os
import openpyxl
from openpyxl.utils.datetime import from_excel
from datetime import datetime, timedelta


def obtener_fecha(sheet, celda):
    valor = sheet[celda].value

    if isinstance(valor, (int, float)):  # Si es un número, convertirlo
        fecha_convertida = datetime(1899, 12, 30) + timedelta(days=valor)
        return fecha_convertida.strftime("%Y-%m-%d")  # Devuelve la fecha en formato legible

    elif isinstance(valor, datetime):  # Si ya es datetime, devolverlo formateado
        return valor.strftime("%Y-%m-%d")

    elif isinstance(valor, str):  # Si está guardado como texto, intenta parsearlo
        try:
            fecha_parseada = datetime.strptime(valor, "%m-%d-%Y")  # Ajusta formato si es necesario
            return fecha_parseada.strftime("%Y-%m-%d")
        except ValueError:
            return f"⚠️ No se pudo convertir: {valor}"

    return "No encontrado"  # Si está vacío

def procesar_archivos(carpeta_cotizaciones, archivo_base):
    # Cargar base de datos o crear nueva
    if os.path.exists(archivo_base):
        df_base = pd.read_excel(archivo_base, engine="openpyxl")
    else:
        df_base = pd.DataFrame(columns=["Archivo", "Empresa", "Requisitor", "No. Cotización", "Po", "Fecha de Po", "Descripción", 
                                        "Cantidad", "Precio Unitario", "Importe", "Subtotal", "IVA", "Total", "Tipo de moneda"])

    # Iterar sobre archivos en la carpeta
    for archivo in os.listdir(carpeta_cotizaciones):
        if archivo.endswith((".xlsx", ".xlsm")):
            ruta_archivo = os.path.join(carpeta_cotizaciones, archivo)
            print(f"📂 Procesando: {archivo}")

            try:
                    wb = openpyxl.load_workbook(ruta_archivo, data_only= True)
                    if "cotizacion" not in wb.sheetnames:
                        print(f"⚠️ {archivo} no contiene la hoja 'cotizacion'.")
                       

                        continue
                    df_excel = pd.read_excel(ruta_archivo, sheet_name="cotizacion", engine="openpyxl")
                    print(df_excel.head(30))

                    sheet = wb["cotizacion"]
                    print(f"🔎 PO: {sheet['B6'].value}")
                    print(f"🔎 Fecha PO: {sheet['A6'].value}")
                    print(f"🔎 No. Cotización: {sheet['G6'].value}")
                    print(f"🔎 Empresa: {sheet['C6'].value}")
                    print(f"🔎 Requisitor: {sheet['H6'].value}")
                    print(f"🔎 Subtotal: {sheet['H33'].value}")
                    print(f"🔎 IVA: {sheet['H36'].value}")
                    print(f"🔎 Total: {sheet['H37'].value}")
                    print(f"🔎 Importe: {sheet['H9'].value}")
                    # Leer datos generales
                    po = sheet["B6"].value or "No encontrado"
                    fecha_po = obtener_fecha(sheet, "A6") or "No encontrado"
                    no_cotizacion = sheet["G6"].value or "No encontrado"
                    empresa = sheet["C6"].value or "No encontrado"
                    requisitor = sheet["H6"].value or "No encontrado"
                    importe = sheet["H9"].value or "No encontrado"
                    subtotal = sheet["H33"].value or "No encontrado"
                    iva = sheet["H36"].value or "No encontrado"
                    total = sheet["H37"].value or "No encontrado"
                    print (f"La fecha de po deberia ser: {fecha_po}")
                    # Recorrer materiales
                    fila = 9
                    nuevas_filas = []

                    while fila <= 32:  # Se mantiene el límite de filas
                        descripcion = sheet[f"A{fila}"].value
                        cantidad = sheet[f"F{fila}"].value
                        precio_unidad = sheet[f"G{fila}"].value
                        importe = sheet[f"H{fila}"].value
                        tipo_moneda = sheet[f"I{fila}"].value

                        print(f"🔎 Fila {fila}: {descripcion}, {cantidad}, {precio_unidad}, {tipo_moneda}")  # Debug

                        # Si la descripción está vacía, se ignora, pero no se detiene
                        if descripcion:
                            nuevas_filas.append([archivo, empresa, requisitor, no_cotizacion, po, fecha_po, descripcion, 
                                        cantidad, precio_unidad, importe, subtotal, iva, total, tipo_moneda])
                        fila += 1  # Seguir iterando hasta la fila 32

                    # Crear DataFrame con nuevas filas
                    print(f"📊 Nuevas filas extraídas: {len(nuevas_filas)}")
                    df_nuevo = pd.DataFrame(nuevas_filas, columns=df_base.columns)
                    if df_nuevo.empty:
                        print("⚠️ Advertencia: No se encontraron nuevas filas en este archivo.")
                    # Evitar duplicados antes de concatenar
                    df_base = pd.concat([df_base, df_nuevo]).drop_duplicates(subset=["Archivo", "No. Cotización", "Descripción"],keep="first").reset_index(drop=True)

            except Exception as e:
                print(f"⚠️ Error en {archivo}: {e}")

    print(f"📊 Tamaño del DataFrame antes de guardar: {df_base.shape}")

    # Guardar archivo actualizado
    df_base.to_excel(archivo_base, index=False, engine="openpyxl")
    print("✅ Base de datos actualizada.")