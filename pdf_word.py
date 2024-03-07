import os
from pdf2docx import Converter

# Ruta de la carpeta con archivos .pdf
carpeta_pdf = r"C:\Users\usuario\Desktop\Documentos Almendro\DEMRE\PAES 2023\PRUEBAS"

# Ruta de la carpeta de salida para archivos .docx
carpeta_docx = r"C:\Users\usuario\Desktop\Documentos Almendro\DEMRE\PAES 2023\PRUEBAS\DOCX"

# Asegúrate de que la carpeta de salida exista, si no, créala
if not os.path.exists(carpeta_docx):
    os.makedirs(carpeta_docx)

# Recorre la carpeta con archivos .pdf
for archivo_pdf in os.listdir(carpeta_pdf):
    if archivo_pdf.endswith(".pdf"):
        ruta_pdf = os.path.join(carpeta_pdf, archivo_pdf)
        try:
            cv = Converter(ruta_pdf)
            cv.convert(carpeta_docx+'/' + archivo_pdf + '.docx', start=0, end=None)
            cv.close()
        except Exception as e:
            print(f"Error al convertir {archivo_pdf}: {e}")

print("Conversión completada.")
