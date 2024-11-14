import pytesseract
import re
import fitz  # PyMuPDF (solo esta librería)
from PIL import Image, ImageEnhance, ImageFilter

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\practicante.ti\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def preprocesar_imagen(imagen):

    imagen_gris = imagen.convert("L")

    enhancer = ImageEnhance.Contrast(imagen_gris)
    imagen_contraste = enhancer.enhance(2.0)  

    imagen_suavizada = imagen_contraste.filter(ImageFilter.MedianFilter(3))

    return imagen_suavizada

def extraer_rut_de_imagen(imagen):
    imagen = preprocesar_imagen(imagen)

    texto = pytesseract.image_to_string(imagen)
    texto = texto.replace("\n", " ").replace("\r", " ").strip()

    print("Texto extraído de la imagen:")
    print(texto)


    match = re.search(r'RUN:\s*([0-9]{1,2}\.[0-9]{3}\.[0-9]{3}[0-9Kk-]{1})', texto)
    
    if match:
        rut = match.group(1)
 
        if rut[-2] == '4' and rut[-1] in '0123456789Kk':
            rut = rut[:-2] + '-' + rut[-1]  
        
        return rut 
    return None 

def extraer_texto_de_pagina(pdf_path, pagina_num):
    try:

        doc = fitz.open(pdf_path)

       
        if pagina_num < 0 or pagina_num >= doc.page_count:
            print(f"Error: El número de página {pagina_num} no es válido. El documento tiene {doc.page_count} páginas.")
            return
        
        pagina = doc.load_page(pagina_num)
        
        pix = pagina.get_pixmap()

        imagen = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        rut = extraer_rut_de_imagen(imagen)
        
        if rut:
            print(f"RUT encontrado: {rut}")
        else:
            print("No se encontró el RUT.")
    
    except Exception as e:
        print(f"Error al procesar el PDF: {e}")

if __name__ == '__main__':
    pdf_path = '75 CERTIFICADO MANIPULADOR PAE.pdf'
    
    try:
        pagina_num = int(input("Ingrese el número de página que desea procesar: ")) - 1  # Restamos 1 para hacerlo 0-indexado
        extraer_texto_de_pagina(pdf_path, pagina_num)
    except ValueError:
        print("Por favor ingrese un número válido de página.")
