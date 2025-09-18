from openpyxl import load_workbook

def leer_excel_por_orden(path_excel):
    """
    Lee las celdas AV25, AW25 y AX25 de todas las hojas visibles del archivo Excel,
    y las guarda usando "01", "02", etc., como claves según el orden en el libro.
    Si hay hojas duplicadas con el mismo número, solo toma la primera que tenga valores.
    No se leen hojas ocultas.
    """
    try:
        wb = load_workbook(filename=path_excel, data_only=True)
    except Exception as e:
        return {}

    resultado = {}
    seen_keys = set()  # Conjunto para almacenar los números ya procesados

    # Recorremos todas las hojas, solo las visibles
    for i, nombre_hoja in enumerate(wb.sheetnames, start=1):
        ws = wb[nombre_hoja]

        # Comprobamos si la hoja está oculta
        if ws.sheet_state in ['hidden', 'veryHidden']:
            continue  # Omitir esta hoja y pasar a la siguiente

        # Extraemos el número de la hoja para compararlo (ej: "02 AB" → "02")
        clave = nombre_hoja.split()[0]  # Extrae solo el número (primer fragmento del nombre de la hoja)

        # Asegurarnos de que la clave es un número de 2 dígitos (si no tiene 2 dígitos, lo completamos)
        clave = clave.zfill(2)

        # Comprobamos si ya hemos procesado una hoja con este número de día
        if clave in seen_keys:
            continue  # Omitir esta hoja porque ya hemos procesado una con ese número

        # Si no hemos procesado una hoja con ese número, la procesamos
        valor_av = ws["AV25"].value or 0
        valor_aw = ws["AW25"].value or 0
        valor_ax = ws["AX25"].value or 0

        # Añadimos la hoja al resultado solo si tiene valores (incluso si son 0)
        resultado[f"{i:02d}"] = [valor_av, valor_aw, valor_ax]

        # Registramos que hemos procesado esta clave
        seen_keys.add(clave)

    return resultado


def leer_excel_jaltipan(path_excel):
    """
    Lee las celdas AV29, AW29 y AX29 de las hojas visibles del archivo Excel,
    ignorando aquellas que contienen la palabra 'manual' en el nombre.
    Usa como clave un número incremental ("01", "02", etc.) y evita duplicados por día.
    """
    try:
        wb = load_workbook(filename=path_excel, data_only=True)
    except Exception as e:
        print("Error al cargar el archivo:", e)
        return {}

    resultado = {}
    seen_keys = set()

    for i, nombre_hoja in enumerate(wb.sheetnames, start=1):
        ws = wb[nombre_hoja]

        if ws.sheet_state in ['hidden', 'veryHidden']:
            continue

        # Ignorar hojas que contienen 'manual'
        if 'manual' in nombre_hoja.lower():
            continue

        # Extraer el número (día) de la hoja
        clave = ''.join(filter(str.isdigit, nombre_hoja))[:2].zfill(2)

        if clave in seen_keys:
            continue

        valor_av = ws["AV29"].value or 0
        valor_aw = ws["AW29"].value or 0
        valor_ax = ws["AX29"].value or 0

        resultado[f"{len(resultado)+1:02d}"] = [valor_av, valor_aw, valor_ax]
        seen_keys.add(clave)

    return resultado