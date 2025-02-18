import pandas as pd
# Cargar el archivo una sola vez
archivo = "tu_archivo.xlsx"
xls = pd.ExcelFile(archivo)

# Ver nombres de las hojas
print(xls.sheet_names)  # Lista con los nombres de las hojas

# Leer todas las hojas en un diccionario
dfs = {hoja: xls.parse(hoja) for hoja in xls.sheet_names}

# Acceder a una hoja específica
df_hoja1 = dfs["NombreDeLaHoja"]