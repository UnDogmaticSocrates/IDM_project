import pandas as pd
# Cargar el archivo una sola vez
archivo = "control_facturation_2025.xlsm"
xls = pd.ExcelFile(archivo)

# Ver nombres de las hojas
print(xls.sheet_names)  # Lista con los nombres de las hojas

# Leer todas las hojas en un diccionario
dfs = {hoja: xls.parse(hoja) for hoja in xls.sheet_names}

# Acceder a una hoja específica
df_hoja1 = dfs["2024"]
print(df_hoja1.info())
print()
print(df_hoja1.head())
print()
print(df_hoja1.describe())