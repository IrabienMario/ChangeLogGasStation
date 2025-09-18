import xlwings as xw

def pegar_datos_en_hojas(path_excel_destino, datos):
    app = xw.App(visible=False)  # No muestra Excel
    wb = app.books.open(path_excel_destino)

    for (fecha, nombre), valores in datos.items():
        if nombre not in [s.name for s in wb.sheets]:
            continue  # Saltamos si no hay hoja con ese nombre

        sht = wb.sheets[nombre]
        celdas_fecha = sht.range("B9:B500").value  # Rango razonable

        filas_a_pegar = []

        # Buscar filas donde la fecha coincide
        for i, valor in enumerate(celdas_fecha):
            if hasattr(valor, "date") and valor.date() == fecha.date():
                filas_a_pegar.append(9 + i)

        if len(filas_a_pegar) >= 3:
            for idx, fila in enumerate(filas_a_pegar[:3]):
                sht.range(f"C{fila}").value = valores[idx]

    wb.save()
    wb.close()
    app.quit()
