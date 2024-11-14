import pytesseract
import fitz  # PyMuPDF
import pandas as pd
import re
import gc
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\practicante.ti\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

def corregir_rut(rut_extraido):
    rut_extraido = rut_extraido.replace(',', '.')

    if rut_extraido.count('.') != 2:
        rut_extraido = rut_extraido[:2] + '.' + rut_extraido[2:5] + '.' + rut_extraido[5:8] + '-' + rut_extraido[8:]

    if '-' not in rut_extraido:
        rut_extraido = rut_extraido[:8] + '-' + rut_extraido[8:]

    if len(rut_extraido) > 9 and rut_extraido[-2] == '4':
        rut_extraido = rut_extraido[:-2] + '-' + rut_extraido[-1]
        
    return rut_extraido

def extraer_rut_de_imagen(imagen):
    custom_config = r'--oem 3 --psm 6'

    texto = pytesseract.image_to_string(imagen, config=custom_config)
    texto = texto.replace("\n", " ").replace("\r", " ").strip()

    match = re.search(r'RUN:\s*([0-9]{1,2}[.,]?\d{3}[.,]?\d{3}[-]?[0-9Kk]{1})', texto)

    if match:
        rut_extraido = match.group(1) 
        rut_corregido = corregir_rut(rut_extraido)
        print(f"RUT extraído (corregido): {rut_corregido}")
        return rut_corregido
    return None  

def procesar_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    ruts = []  
    paginas_sin_rut = []  

    for i in range(doc.page_count):  
        page = doc.load_page(i) 
        pix = page.get_pixmap() 
        
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        rut = extraer_rut_de_imagen(img)
        if rut:
            ruts.append({'Pagina': i + 1, 'RUT': rut})
        else:
            paginas_sin_rut.append(i + 1)  
        
        del img  
        gc.collect()

    return ruts, paginas_sin_rut

def guardar_en_excel(ruts, output_path):

    df = pd.DataFrame(ruts)
    

    df.to_excel(output_path, index=False)
    print(f"RUTs guardados en: {output_path}")

if __name__ == '__main__':
    pdf_path = '75 CERTIFICADO MANIPULADOR PAE.pdf' 
    ruts, paginas_sin_rut = procesar_pdf(pdf_path)

    if ruts:
        output_path = 'ruts_extraidos.xlsx'
        guardar_en_excel(ruts, output_path)
    else:
        print("No se encontraron RUTs en el PDF.")
    if paginas_sin_rut:
        print(f"\nNo se pudo extraer el RUT en las siguientes páginas: {', '.join(map(str, paginas_sin_rut))}")
    else:
        print("Todos los RUTs fueron extraídos con éxito.")
