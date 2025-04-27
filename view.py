import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk


def obtener_ruta_recurso(ruta_relativa):
    """Obtiene la ruta absoluta del recurso, compatible en .py normal y .exe."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

class App:
    def __init__(self):
        self.controller = None
        self.carpeta_salida = ""

        self.root = tk.Tk()
        self.root.title("Conversor PDF")

        # Tamaño de ventana
        ancho = 600
        alto = 400

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (ancho // 2)
        y = (screen_height // 2) - (alto // 2)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.root.resizable(False, False)
        self.root.config(bg="#121212")

        # Cambiar ícono de ventana usando logo.png
        ruta_icono = obtener_ruta_recurso("logo.png")
        if os.path.exists(ruta_icono):
            img_icono = Image.open(ruta_icono)
            img_icono = img_icono.resize((32, 32))
            self.icono = ImageTk.PhotoImage(img_icono)
            self.root.iconphoto(False, self.icono)

        self.imagen = None

    def set_controller(self, controller):
        self.controller = controller
        self.crear_widgets()

    def crear_widgets(self):
        bg_color = "#121212"
        fg_color = "#f0f0f0"
        button_color = "#1E88E5"
        button_hover = "#1565C0"
        button_text_color = "#ffffff"
        font_general = ("Segoe UI", 11)
        font_titulo = ("Segoe UI", 18, "bold")

        # Mostrar logo grande usando obtener_ruta_recurso
        ruta_logo = obtener_ruta_recurso("logo.png")
        if os.path.exists(ruta_logo):
            img = Image.open(ruta_logo)
            img = img.resize((80, 80))
            self.imagen = ImageTk.PhotoImage(img)
            label_img = tk.Label(self.root, image=self.imagen, bg=bg_color)
            label_img.pack(pady=(10, 0))

        titulo = tk.Label(self.root, text="Conversor Word/Excel → PDF", font=font_titulo, bg=bg_color, fg=fg_color)
        titulo.pack(pady=10)

        boton_archivo = tk.Button(
            self.root,
            text="Seleccionar archivo",
            command=self.controller.seleccionar_archivo,
            font=font_general,
            bg="#43A047",
            fg=button_text_color,
            activebackground="#388E3C",
            activeforeground=button_text_color,
            bd=0,
            relief="ridge",
            padx=15, pady=10
        )
        boton_archivo.pack(pady=15)

        boton_carpeta = tk.Button(
            self.root,
            text="Seleccionar carpeta de salida (opcional)",
            command=self.controller.seleccionar_carpeta_salida,
            font=font_general,
            bg=button_color,
            fg=button_text_color,
            activebackground=button_hover,
            activeforeground=button_text_color,
            bd=0,
            relief="ridge",
            padx=15, pady=10
        )
        boton_carpeta.pack(pady=10)

        self.label_carpeta = tk.Label(
            self.root,
            text="No hay carpeta seleccionada.",
            font=font_general,
            bg=bg_color,
            fg=fg_color,
            wraplength=500,
            justify="center"
        )
        self.label_carpeta.pack(pady=15)

    def mostrar_mensaje(self, titulo, mensaje):
        messagebox.showinfo(titulo, mensaje)

    def mostrar_error(self, titulo, mensaje):
        messagebox.showerror(titulo, mensaje)

    def actualizar_carpeta(self, carpeta):
        self.carpeta_salida = carpeta
        self.label_carpeta.config(text=f"Carpeta de salida:\n{self.carpeta_salida}")

    def ejecutar(self):
        self.root.mainloop()
