import pytesseract
import fitz  # PyMuPDF
import pandas as pd
import re
import gc
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\practicante.ti\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def extraer_rut_de_imagen(imagen):

    custom_config = r'--oem 3 --psm 6'

    texto = pytesseract.image_to_string(imagen, config=custom_config)
    texto = texto.replace("\n", " ").replace("\r", " ").strip()


    match = re.search(r'RUN:\s*([0-9]{1,2}\.[0-9]{3}\.[0-9]{3}-[0-9Kk]{1})', texto)
    
    if match:
        rut_con_puntos = match.group(1) 
        rut_sin_puntos = rut_con_puntos.replace('.', '') 
        
        if len(rut_sin_puntos) > 2 and rut_sin_puntos[-2] == '4':
            if rut_sin_puntos[-1].isdigit() or rut_sin_puntos[-1].upper() == 'K':

                rut_sin_puntos = rut_sin_puntos[:-2] + '-' + rut_sin_puntos[-1]

        print(f"RUT extraído modificado: {rut_sin_puntos}")
        return rut_sin_puntos, texto 
    return None, texto 

def procesar_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    ruts = []
    paginas_sin_rut = []  

    for i in range(doc.page_count): 
        page = doc.load_page(i)
        pix = page.get_pixmap() 
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        rut, texto_extraido = extraer_rut_de_imagen(img)
        if rut:
            ruts.append(rut) 
        else:
            paginas_sin_rut.append(i + 1) 
            print(f"Página {i + 1}: No se pudo extraer RUT. Texto extraído: {texto_extraido}")
        
        del img 
        gc.collect()

    return ruts, paginas_sin_rut

def guardar_en_excel(ruts, output_path):
    df = pd.DataFrame(ruts, columns=['RUT'])
    df.to_excel(output_path, index=False)

if __name__ == '__main__':
    pdf_path = '75 CERTIFICADO MANIPULADOR PAE.pdf' 
    
    ruts, paginas_sin_rut = procesar_pdf(pdf_path)
    if ruts:
        print("\nRUTs extraídos:")
        for rut in ruts:
            print(f"RUT extraído: {rut}")
    else:
        print("No se encontraron RUTs en el PDF.")
    
    if paginas_sin_rut:
        print(f"\nNo se pudo extraer el RUT en las siguientes páginas: {', '.join(map(str, paginas_sin_rut))}")
    else:
        print("Todos los RUTs fueron extraídos con éxito.")
