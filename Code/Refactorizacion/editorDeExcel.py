from openpyxl import load_workbook
import xlwings as xw


def leerDatosDelExcel(path_excel, logica_por_hoja):

    try:
        wb = load_workbook(filename=path_excel, data_only=True)
    except Exception as e:
        print("Error al cargar el archivo:", e)
        return {}

    resultado = {}
    seen_keys = set()

    for idx, nombre_hoja in enumerate(wb.sheetnames):
        ws = wb[nombre_hoja]

        if ws.sheet_state in ['hidden', 'veryHidden']:
            continue

        # Función personalizada define qué hacer con la hoja
        key, valores = logica_por_hoja(idx, nombre_hoja, ws, seen_keys)

        if key and key not in seen_keys:
            resultado[key] = valores
            seen_keys.add(key)

    return resultado

def pegarDatosEnExcel(path_excel, logica_por_hoja, save=True, close=True):
    """
    Aplica una lógica personalizada a cada hoja visible del archivo Excel usando xlwings.
    
    Parámetros:
    - path_excel: ruta al archivo Excel.
    - logica_por_hoja: función que recibe (idx, nombre_hoja, hoja_xlwings) y modifica la hoja.
    - save: si True, guarda los cambios.
    - close: si True, cierra el archivo al terminar.
    """
    try:
        app = xw.App(visible=False)
        wb = app.books.open(path_excel)
    except Exception as e:
        print("Error al abrir el archivo:", e)
        return False

    for idx, hoja in enumerate(wb.sheets):
        if hoja.api.Visible != -1:  # -1 significa visible (según la API de Excel)
            logica_por_hoja(idx, hoja.name, hoja)

    if save:
        wb.save()

    if close:
        wb.close()
        app.quit()

    return True

