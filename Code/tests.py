import fitz  # PyMuPDF
import pandas as pd
import re
import os

def extraer_info_pdf(ruta_pdf, numeros_objetivo):
    doc = fitz.open(ruta_pdf)
    datos_encontrados = []

    for pagina in doc:
        texto = pagina.get_text()
        lineas = texto.split('\n')
        for i, linea in enumerate(lineas):
            matches = re.findall(r'(\d{8})([A-Z])', linea)
            for num_8dig, letra in matches:
                ultimos_dos = num_8dig[-2:]

                if ultimos_dos in numeros_objetivo:
                    # Validar que existan líneas alrededor
                    if i - 3 < 0 or i + 3 >= len(lineas):
                        continue

                    fecha_operacion = lineas[i - 3].strip()
                    descripcion = linea.strip()       # <-- ahora conservamos TODO
                    deposito = lineas[i + 3].strip()

                    if deposito != "-" and deposito.startswith("$"):
                        datos_encontrados.append({
                            "Fecha de Operación": fecha_operacion,
                            "Descripción": descripcion,
                            "Depósitos": deposito
                        })
    #guardar_en_excel(datos_encontrados, "datos_extraidos.xlsx")
    doc.close()
    return datos_encontrados

