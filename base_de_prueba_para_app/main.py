import json
from procesador import procesar_archivos

# Cargar configuración
with open("/workspaces/IDM_project/data_base_auto_app/config.json", "r") as file:
    config = json.load(file)

# Ejecutar procesamiento
procesar_archivos(config["carpeta_cotizaciones"], config["archivo_base"])
print("✅ Proceso completado.")