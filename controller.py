import os
from tkinter import filedialog
import model

class Controller:
    def __init__(self, app):
        self.app = app

    def seleccionar_archivo(self):
        ruta_archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Word o Excel",
            filetypes=[("Documentos", "*.docx *.xlsx")]
        )
        if not ruta_archivo:
            return
        
        extension = os.path.splitext(ruta_archivo)[1].lower()
        carpeta_salida = self.app.carpeta_salida if self.app.carpeta_salida else None

        if extension == '.docx':
            exito = model.convertir_word(ruta_archivo, carpeta_salida)
        elif extension == '.xlsx':
            exito = model.convertir_excel(ruta_archivo, carpeta_salida)
        else:
            self.app.mostrar_error("Error", "Formato no compatible.")
            return

        if exito:
            self.app.mostrar_mensaje("Éxito", f"Archivo convertido exitosamente.\n\nGuardado en:\n{carpeta_salida if carpeta_salida else os.path.dirname(ruta_archivo)}")
        else:
            self.app.mostrar_error("Error", "Hubo un problema en la conversión.")

    def seleccionar_carpeta_salida(self):
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if carpeta:
            self.app.actualizar_carpeta(carpeta)
