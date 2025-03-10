import pandas as pd
import openpyxl

import openpyxl

# Cargar el archivo Excel
file_path = "prueba.xlsm"  # Cambia esto por el nombre de tu archivo
wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = wb["Cotización"]

# Función para obtener el valor de una celda específica
def obtener_valor(celda):
    return sheet[celda].value if sheet[celda].value else "No encontrado"

# Extraer los datos según las coordenadas proporcionadas
datos_extraidos = {
    "Empresa": obtener_valor("H3"),  # Empresa (H-L, 2-3) -> Tomo H2 como referencia
    "Requisitor": obtener_valor("I9"),  # Requisitor (H-L, 9) -> Tomo I9
    "Número de Cotización": obtener_valor("E9"),  # Número de Cotización (E9)
    "Cantidad": obtener_valor("H14"),  # Cantidad (H14)
    "Descripción": obtener_valor("A14"),  # Descripción (A14 como inicio de la tabla)
}

# Mostrar los datos extraídos
print(datos_extraidos)