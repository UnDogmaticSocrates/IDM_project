from PIL import Image
import os

# Ruta al folder donde están las imágenes
folder_path = '/workspaces/IDM_project/Imagenes_omar'

# Lista de imágenes en el folder
images = [f for f in os.listdir(folder_path) if f.endswith(('.png', '.jpg', '.jpeg'))]

# Ordena las imágenes si es necesario (opcional)
images.sort()

# Abre las imágenes y conviértelas en formato RGB (si es necesario)
image_list = []
for image_name in images:
    image_path = os.path.join(folder_path, image_name)
    img = Image.open(image_path).convert('RGB')  # Convertir a RGB si es necesario
    image_list.append(img)

# Guarda las imágenes como un archivo PDF
output_pdf_path = '/workspaces/IDM_project/Imagenes_omar/graficas_omar.pdf'
image_list[0].save(output_pdf_path, save_all=True, append_images=image_list[1:], resolution=100.0, quality=95)

print(f"PDF generado con éxito: {output_pdf_path}")