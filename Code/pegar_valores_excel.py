import xlwings as xw

def pegar_valores_excel(ruta_target, valores_por_hoja, nombre_hoja):
    app = xw.App(visible=False)
    wb = app.books.open(ruta_target)
    
    hoja_destino = wb.sheets[nombre_hoja]  # Hoja de destino
    fila_inicial = 5
    columna_destino = 'E'

    fila_actual = fila_inicial

    for hoja, valores in valores_por_hoja.items():
        if len(valores) != 3:
            print(f"Advertencia: La hoja '{hoja}' no tiene exactamente tres valores. Se omitirá.")
            continue

        for i in range(3):
            celda = f"{columna_destino}{fila_actual + i}"
            hoja_destino.range(celda).value = valores[i]
        
        fila_actual += 3  # Avanzar a la siguiente tanda de 3

    wb.save()
    wb.close()
    app.quit()
