import openpyxl
import xlwings as xw
import re
from datetime import datetime
from tests import extraer_info_pdf  # Ajusta si tu módulo está en otro paquete

columna_numeros = "I"

def process_target_excel(ruta_target, ruta_pdf):
    app = xw.App(visible=False)
    wb = app.books.open(ruta_target)
    
    hojas = ['COATZA 1', 'ZARAGOZA', 'JALTIPAN']
    for hoja_nombre in hojas:
        sht = wb.sheets[hoja_nombre]
        texto_b1 = sht.range('B1').value

        match = re.search(r'(\d{2})/(?=\w)', texto_b1)
        if match:
            numeros_objetivo = match.group(1)
            print(f"Hoja: {hoja_nombre}, Números encontrados: {numeros_objetivo}")

            datos_encontrados = extraer_info_pdf(ruta_pdf, numeros_objetivo)

            # AGRUPAR fechas únicas en un set
            fechas_depositos = {}
            for deposito in datos_encontrados:
                fecha_operacion = deposito.get("Fecha de Operación")
                monto_str = deposito.get("Depósitos")  # Se obtiene el monto como cadena

                # Eliminar el símbolo de dólar y las comas, y convertir el monto a float
                if monto_str is not None:
                    monto_str = monto_str.replace('$', '').replace(',', '')  # Eliminar el '$' y las comas
                    try:
                        monto = float(monto_str)  # Convertir la cadena a número
                    except ValueError:
                        print(f"Error al convertir monto: {monto_str}")
                        continue  # Si no se puede convertir, se salta este depósito
                else:
                    print(f"Error: Monto no válido o ausente en el depósito: {deposito}")
                    continue

                if isinstance(fecha_operacion, str):
                    fecha_operacion = datetime.strptime(fecha_operacion, "%d/%m/%Y").date()
                elif isinstance(fecha_operacion, datetime):
                    fecha_operacion = fecha_operacion.date()

                # Agrupar depósitos por fecha
                if fecha_operacion not in fechas_depositos:
                    fechas_depositos[fecha_operacion] = []
                fechas_depositos[fecha_operacion].append({'Fecha de Operación': fecha_operacion, 'Monto': monto})

            row = 5
            while sht.range(f'K{row}').value is not None:
                fecha_excel = sht.range(f'K{row}').value

                if isinstance(fecha_excel, str):
                    fecha_excel = datetime.strptime(fecha_excel, "%d/%m/%Y").date()
                elif isinstance(fecha_excel, datetime):
                    fecha_excel = fecha_excel.date()

                ##
                # Verificar si hay depósitos para esa fecha
                if fecha_excel in fechas_depositos:
                    depositos = fechas_depositos[fecha_excel]
                    
                    i = 0  # Variable para contar las filas donde se van a colocar los depósitos

                    # Si hay más de tres depósitos
                    if len(depositos) > 3:
                        # Colocar el primer depósito en la fila actual (row)
                        sht.range(f'I{row + i}').value = str(depositos[0]['Monto'])  # Primer depósito
                        i += 1

                        # Colocar el segundo depósito en la fila siguiente (row + 1)
                        sht.range(f'I{row + i}').value = str(depositos[1]['Monto'])  # Segundo depósito
                        i += 1

                        # Crear la fórmula de la suma para los depósitos restantes en la fila row + 2
                        if len(depositos) > 2:  # Si hay más de dos depósitos
                            formula = "=" + "+".join([str(deposito['Monto']) for deposito in depositos[2:]])
                            sht.range(f'I{row + i}').formula = formula  # Sumar los demás depósitos
                    else:
                        # Si hay tres o menos depósitos, colocamos todos en las filas correspondientes
                        for j in range(len(depositos)):
                            sht.range(f'I{row + i}').value = str(depositos[j]['Monto'])
                            i += 1  # Avanzar a la siguiente fila

                    ##

                else:
                    sht.range(f'K{row + 1}').value = "NO HAY DEPOSITOS EN LA FECHA"

                row += 3  # Avanzamos tres filas para la siguiente fecha
        else:
            print(f"No se encontraron los dos números antes del '/' en la celda B1 de {hoja_nombre}.")
    
    wb.save()
    wb.close()
