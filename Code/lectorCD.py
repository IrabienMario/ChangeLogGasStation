from openpyxl import load_workbook

def leer_creditos_debitos(path_excel):
    wb = load_workbook(filename=path_excel, data_only=True)
    creditos = {}
    debitos = {}

    fechas_vistas = set()

    nombres_creditos = {
        "CRYSA": 28,
        "MINSA": 29,
        "FORA": 32,
        "TOLEDO": 31
    }

    nombres_debitos = {
        "ARTURO JUAREZ": 37,
        "ELMET ": 38,
        "MOLINA": 39,
        "CLEAN": 40,
        "CARRIPOLLO": 41,
        "MUNICIPIO": 42,
        "OSCAR LUNA": 44
    }

    columnas = ["M", "AB", "AQ"]

    for idx, hoja in enumerate(wb.sheetnames):
        # Solo procesamos hojas impares (0, 2, 4, ...)
        if idx % 2 != 0:
            continue

        ws = wb[hoja]

        fecha = ws["G2"].value
        if fecha in fechas_vistas:
            continue  # Ignorar fechas repetidas
        fechas_vistas.add(fecha)

        # Procesar Créditos
        for nombre, fila in nombres_creditos.items():
            valores = []
            for col in columnas:
                celda = f"{col}{fila}"
                valor = ws[celda].value or 0
                valores.append(valor)
            creditos[(fecha, nombre)] = valores

        # Procesar Débitos
        for nombre, fila in nombres_debitos.items():
            valores = []
            for col in columnas:
                celda = f"{col}{fila}"
                valor = ws[celda].value or 0
                valores.append(valor)
            debitos[(fecha, nombre)] = valores

    return creditos, debitos
