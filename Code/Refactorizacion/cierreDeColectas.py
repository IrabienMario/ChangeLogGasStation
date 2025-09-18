from editorDeExcel import leerDatosDelExcel,pegarDatosEnExcel

def obtenerColectasDeZaragozaYCoatza(path_excel):
    def logica_por_hoja(idx, nombre_hoja, ws, seen_keys):
        clave = nombre_hoja.split()[0].zfill(2)
        if clave in seen_keys:
            return None, None

        valor_av = ws["AV25"].value or 0
        valor_aw = ws["AW25"].value or 0
        valor_ax = ws["AX25"].value or 0

        return f"{idx+1:02d}", [valor_av, valor_aw, valor_ax]

    return leerDatosDelExcel(path_excel, logica_por_hoja)


def obtenerColectasDeJaltipan(path_excel):
    def logica_por_hoja(idx, nombre_hoja, ws, seen_keys):
        if 'manual' in nombre_hoja.lower():
            return None, None

        clave = ''.join(filter(str.isdigit, nombre_hoja))[:2].zfill(2)
        if clave in seen_keys:
            return None, None

        valor_av = ws["AV29"].value or 0
        valor_aw = ws["AW29"].value or 0
        valor_ax = ws["AX29"].value or 0

        return f"{len(seen_keys)+1:02d}", [valor_av, valor_aw, valor_ax]

    return leerDatosDelExcel(path_excel, logica_por_hoja)


def RegistrarColectas(ruta_target, valores_por_hoja, nombre_hoja_destino):
    """
    Pega valores en una hoja destino de un Excel, avanzando de tres en tres filas.
    
    - ruta_target: ruta del archivo Excel destino.
    - valores_por_hoja: diccionario con clave como hoja ("01", "02", etc.) y lista de 3 valores.
    - nombre_hoja_destino: hoja donde se pegarán todos los datos.
    """
    def logica_por_hoja(idx, nombre_hoja, hoja_xlwings):
        if nombre_hoja != nombre_hoja_destino:
            return  # Solo pegamos en la hoja destino

        fila_inicial = 5
        columna_destino = 'E'
        fila_actual = fila_inicial

        for hoja, valores in valores_por_hoja.items():
            if len(valores) != 3:
                print(f"Advertencia: La hoja '{hoja}' no tiene exactamente tres valores. Se omitirá.")
                continue

            for i in range(3):
                celda = f"{columna_destino}{fila_actual + i}"
                hoja_xlwings.range(celda).value = valores[i]

            fila_actual += 3

    return pegarDatosEnExcel(ruta_target, logica_por_hoja)