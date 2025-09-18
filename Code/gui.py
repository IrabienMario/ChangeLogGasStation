import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from excel_editor import process_target_excel
from lector_excel import leer_excel_por_orden, leer_excel_jaltipan
from pegar_valores_excel import pegar_valores_excel
from lectorCD import leer_creditos_debitos
from pegar_datos_en_hojas import pegar_datos_en_hojas

import threading

class QuickSearchApp:
    def __init__(self, master):
        self.master = master
        master.title("Procesar Estado de Cuenta")
        master.geometry("400x300")

        # Contenedor central
        frame = tk.Frame(master)
        frame.pack(expand=True)

        self.create_button(frame, "Procesar Estado de Cuenta", self.run_process).pack(pady=10)
        self.create_button(frame, "Procesar CIERRE DE COLECTAS", self.procesar_varios_archivos).pack(pady=10)
        self.create_button(frame, "Procesar Créditos y Débitos", self.procesar_creditos_debitos).pack(pady=10)

    def create_button(self, parent, text, command):
        return tk.Button(parent, text=text, width=30, command=lambda: self.run_with_loader(command))

    def show_loading(self):
        self.loading = tk.Toplevel(self.master)
        self.loading.title("Procesando...")
        self.loading.geometry("200x100")
        self.loading.resizable(False, False)
        self.loading.grab_set()
        ttk.Label(self.loading, text="Procesando, espera...").pack(pady=20)
        self.loading.update()

    def hide_loading(self):
        if hasattr(self, 'loading') and self.loading.winfo_exists():
            self.loading.destroy()

    def run_with_loader(self, func):
        def task():
            self.show_loading()
            try:
                func()
            finally:
                self.hide_loading()
        threading.Thread(target=task).start()

    def run_process(self):
        ruta_pdf = filedialog.askopenfilename(
            title="Selecciona el archivo PDF de estado de cuenta",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not ruta_pdf:
            return

        ruta_target = filedialog.askopenfilename(
            title="Selecciona el Excel de estado de cuenta a actualizar",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not ruta_target:
            return

        try:
            process_target_excel(ruta_target, ruta_pdf)
            messagebox.showinfo("Listo", f"Archivo actualizado:\n{ruta_target}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo procesar:\n{e}")

    def procesar_varios_archivos(self):
        ruta_target = filedialog.askopenfilename(
            title="Selecciona el archivo DESTINO",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not ruta_target:
            return

        rutas_origen = []
        nombres = ["COATZA", "ZARAGOZA", "JALTIPAN"]
        hojas = ["COATZA 1", "ZARAGOZA", "JALTIPAN"]

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
                if nombres[i] == "JALTIPAN":
                    valores_por_hoja = leer_excel_jaltipan(rutas_origen[i])
                else:
                    valores_por_hoja = leer_excel_por_orden(rutas_origen[i])

                pegar_valores_excel(ruta_target, valores_por_hoja, hojas[i])

            messagebox.showinfo("Éxito", "Todos los valores fueron pegados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

    def procesar_creditos_debitos(self):
        archivo_datos = filedialog.askopenfilename(
            title="Selecciona el archivo de DATOS",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not archivo_datos:
            return

        archivo_creditos = filedialog.askopenfilename(
            title="Selecciona el archivo DESTINO de CRÉDITOS",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not archivo_creditos:
            return

        archivo_debitos = filedialog.askopenfilename(
            title="Selecciona el archivo DESTINO de DÉBITOS",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not archivo_debitos:
            return

        try:
            creditos, debitos = leer_creditos_debitos(archivo_datos)
            pegar_datos_en_hojas(archivo_creditos, creditos)
            pegar_datos_en_hojas(archivo_debitos, debitos)
            messagebox.showinfo("Éxito", "Créditos y débitos pegados correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = QuickSearchApp(root)
    root.mainloop()
