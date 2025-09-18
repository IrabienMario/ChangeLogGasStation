import tkinter as tk
from tkinter import filedialog, messagebox
from cierreDeColectas import obtenerColectasDeZaragozaYCoatza, obtenerColectasDeJaltipan, RegistrarColectas


def procesar_varios_archivos():
    # Selección del archivo destino
    ruta_target = filedialog.askopenfilename(
        title="Selecciona el archivo DESTINO",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not ruta_target:
        return

    rutas_origen = []
    nombres = ["COATZA", "ZARAGOZA", "JALTIPAN"]
    hojas = ["COATZA 1", "ZARAGOZA", "JALTIPAN"]

    # Selección de archivos origen
    for nombre in nombres:
        ruta = filedialog.askopenfilename(
            title=f"Selecciona el archivo ORIGEN para {nombre}",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not ruta:
            messagebox.showerror("Cancelado", f"Selección cancelada para {nombre}.")
            return
        rutas_origen.append(ruta)

    try:
        for i in range(3):
            # Cargar los valores por hoja para cada archivo origen
            if nombres[i] == "JALTIPAN":
                valores_por_hoja = obtenerColectasDeJaltipan(rutas_origen[i])
            else:
                valores_por_hoja = obtenerColectasDeZaragozaYCoatza(rutas_origen[i])

            # Registrar los valores en el archivo destino
            RegistrarColectas(ruta_target, valores_por_hoja, hojas[i])

        messagebox.showinfo("Éxito", "Todos los valores fueron pegados correctamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

# Crear la ventana principal de tkinter para iniciar el proceso
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Llamar a la función principal para procesar los archivos
    procesar_varios_archivos()
