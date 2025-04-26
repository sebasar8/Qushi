# Paso 1: Instalar dependencias si es necesario
try:
    import docx
except ImportError:
    import sys
    if "google.colab" in sys.modules:
        !pip install python-docx
    else:
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    import docx

# Paso 2: Código principal
import os
import sys

def extraer_texto_docx(ruta_archivo):
    doc = docx.Document(ruta_archivo)
    texto_completo = []
    for parrafo in doc.paragraphs:
        texto = parrafo.text.strip()
        if texto:
            texto_completo.append(texto)
    return texto_completo

def generar_foda(texto_lista):
    foda = {'Fortalezas': [], 'Debilidades': [], 'Oportunidades': [], 'Amenazas': []}
    seccion_actual = None
    for linea in texto_lista:
        linea = linea.strip()
        if linea.startswith("Fortalezas"):
            seccion_actual = 'Fortalezas'
        elif linea.startswith("Debilidades"):
            seccion_actual = 'Debilidades'
        elif linea.startswith("Oportunidades"):
            seccion_actual = 'Oportunidades'
        elif linea.startswith("Amenazas"):
            seccion_actual = 'Amenazas'
        elif seccion_actual and linea:
            foda[seccion_actual].append(f"- {linea}")
    return foda

def guardar_foda_txt(foda, ruta_salida="analisis_FODA.txt"):
    with open(ruta_salida, "w", encoding="utf-8") as archivo:
        for clave, valores in foda.items():
            archivo.write(f"{clave}:\n")
            archivo.write("\n".join(valores))
            archivo.write("\n\n")
    print(f"✅ Archivo generado: {os.path.abspath(ruta_salida)}")

def seleccionar_archivo():
    if "google.colab" in sys.modules:
        from google.colab import files
        print("🔼 Sube tu archivo .docx para continuar:")
        archivo = files.upload()
        ruta = next(iter(archivo))  # nombre del archivo
    else:
        import tkinter as tk
        from tkinter import filedialog
        tk.Tk().withdraw()
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo .docx",
            filetypes=[("Documentos Word", "*.docx")]
        )
    return ruta

# === EJECUCIÓN ===
ruta_docx = seleccionar_archivo()

if ruta_docx:
    texto = extraer_texto_docx(ruta_docx)
    foda = generar_foda(texto)
    guardar_foda_txt(foda)
else:
    print("⚠️ No se seleccionó ningún archivo.")
