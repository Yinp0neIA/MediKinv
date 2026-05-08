# utils/helpers.py (o donde tengas este archivo)
import os
import sys
from PIL import Image
import customtkinter as ctk


class Helpers:
    """Funciones auxiliares"""

    @staticmethod
    def obtener_ruta_recurso(ruta_relativa):
        """Obtener ruta absoluta de un recurso, compatible con PyInstaller"""
        try:
            # Verificar si estamos en un ejecutable de PyInstaller
            if getattr(sys, 'frozen', False):
                # Estamos en el ejecutable
                base_path = sys._MEIPASS
            else:
                # Estamos en desarrollo
                # Obtener el directorio base del proyecto
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

            # Unir la ruta base con la ruta relativa
            return os.path.join(base_path, ruta_relativa)

        except Exception as e:
            print(f"Error obteniendo ruta: {e}")
            # Fallback: usar ruta relativa
            return ruta_relativa

    @staticmethod
    def cargar_imagen(ruta_relativa, tamaño=None):
        """Cargar una imagen desde assets"""
        try:
            ruta_completa = Helpers.obtener_ruta_recurso(ruta_relativa)

            if not os.path.exists(ruta_completa):
                print(f"Imagen no encontrada: {ruta_completa}")
                return None

            imagen = Image.open(ruta_completa)

            if tamaño:
                imagen = imagen.resize(tamaño, Image.Resampling.LANCZOS)

            return ctk.CTkImage(light_image=imagen, dark_image=imagen, size=tamaño)
        except Exception as e:
            print(f"Error cargando imagen {ruta_relativa}: {e}")
            return None

    @staticmethod
    def centrar_ventana(ventana, ancho, alto):
        """Centrar una ventana en la pantalla"""
        ventana.update_idletasks()
        pantalla_ancho = ventana.winfo_screenwidth()
        pantalla_alto = ventana.winfo_screenheight()

        x = (pantalla_ancho - ancho) // 2
        y = (pantalla_alto - alto) // 2

        ventana.geometry(f"{ancho}x{alto}+{x}+{y}")