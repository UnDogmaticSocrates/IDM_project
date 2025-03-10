import pandas as pd
import pdfplumber 
import re

pdf_path = "pdf_test.pdf"

def eliminar_urls(texto):
    # Elimina URLs del texto
    texto_limpio = re.sub(r'http[s]?://\S+|www\.\S+', '', texto)
    return texto_limpio


def extraer_texto(texto):
    patron = r"[A-Z]\d+"
    asiento = re.findall(patron, texto)
    return asiento

columnas = ["Página", "Asiento"]
df = pd.DataFrame(columns=columnas)

with pdfplumber.open(pdf_path) as pdf:
    for page_num, page in enumerate(pdf.pages, start = 1):
        texto = page.extract_text()
        
        texto_limpio = eliminar_urls(texto)
        asientos = extraer_texto(texto_limpio)
        if asientos:
            df_nueva_fila = pd.DataFrame({"Página": [page_num] * len(asientos), "Asiento": asientos})
            df = pd.concat([df, df_nueva_fila], ignore_index=True)
        

excel_path = "asientos_extraidos.xlsx"
df.to_excel(excel_path, index=False)

print(f"✅ Datos guardados en {excel_path}")