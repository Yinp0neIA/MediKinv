import os


class IconManager:

    @staticmethod
    def aplicar_icono(ventana):
        """Aplica el icono a cualquier ventana (solo Windows)"""
        try:
            icono_path = os.path.join("assets", "images", "logo.ico")
            if os.path.exists(icono_path):
                ventana.iconbitmap(icono_path)
                return True
        except Exception as e:
            print(f"No se pudo cargar el icono: {e}")
        return False