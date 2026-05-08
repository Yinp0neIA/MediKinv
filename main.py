import sys
import os
import customtkinter as ctk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(BASE_DIR)

from views.login_view import LoginView


def debug_icono():
    """Función para debuggear el icono"""
    print("=" * 50)
    print("DEBUG DEL ICONO")
    print("=" * 50)

    # Ruta del icono
    icono_path = os.path.join(BASE_DIR, "assets", "images", "logo.ico")
    print(f"1. Ruta completa del icono: {icono_path}")

    # Verificar si existe el archivo
    existe = os.path.exists(icono_path)
    print(f"2. El archivo existe: {existe}")

    if existe:
        # Obtener tamaño del archivo
        tamaño = os.path.getsize(icono_path)
        print(f"3. Tamaño del archivo: {tamaño} bytes")

        # Verificar extensión
        extension = os.path.splitext(icono_path)[1].lower()
        print(f"4. Extensión del archivo: {extension}")

        # Leer los primeros bytes para identificar tipo de archivo
        try:
            with open(icono_path, 'rb') as f:
                cabecera = f.read(6)
                print(f"5. Cabecera del archivo (hex): {cabecera.hex()}")

                # Identificar tipo de archivo por cabecera
                if cabecera.startswith(b'\x00\x00\x01\x00'):
                    print("   >> Parece ser un archivo ICO válido (formato correcto)")
                elif extension == '.ico':
                    print("   >> Tiene extensión .ico pero la cabecera no es la esperada")
                    print("   >> Puede ser un PNG renombrado o archivo corrupto")
                else:
                    print(f"   >> Tipo de archivo desconocido")
        except Exception as e:
            print(f"   Error al leer cabecera: {e}")

        # Verificar si es realmente un icono usando imghdr
        try:
            import imghdr
            tipo = imghdr.what(icono_path)
            print(f"6. Tipo detectado por imghdr: {tipo}")
            if tipo == 'ico':
                print("   >> Es un icono ICO válido")
            elif tipo == 'png':
                print("   >> Es un archivo PNG, NO es un ICO válido")
            elif tipo is None:
                print("   >> No se pudo identificar el tipo")
        except:
            print("   imghdr no disponible")

        # Intentar cargar con Pillow si está disponible
        try:
            from PIL import Image
            img = Image.open(icono_path)
            print(f"7. Dimensiones de la imagen: {img.size}")
            print(f"8. Modo de la imagen: {img.mode}")
            print(f"9. Formato detectado por Pillow: {img.format}")
            if img.format == 'ICO':
                print("   >> Pillow lo reconoce como ICO válido")
            else:
                print(f"   >> Pillow lo detecta como {img.format}, NO es un ICO válido")
        except ImportError:
            print("   Pillow no instalado, no se puede verificar con Pillow")
        except Exception as e:
            print(f"   Error al abrir con Pillow: {e}")
    else:
        # Verificar si existe la carpeta
        dir_assets = os.path.join(BASE_DIR, "assets")
        dir_images = os.path.join(BASE_DIR, "assets", "images")

        print(f"   La carpeta assets existe: {os.path.exists(dir_assets)}")
        print(f"   La carpeta images existe: {os.path.exists(dir_images)}")

        if os.path.exists(dir_images):
            print("   Archivos en la carpeta images:")
            for archivo in os.listdir(dir_images):
                print(f"      - {archivo}")

    print("=" * 50)


# Monkey patch para el icono
_original_ctk_init = ctk.CTk.__init__


def _patched_ctk_init(self, *args, **kwargs):
    _original_ctk_init(self, *args, **kwargs)
    icono_path = os.path.join(BASE_DIR, "assets", "images", "logo.ico")
    try:
        if os.path.exists(icono_path):
            self.iconbitmap(icono_path)
            print(f"[OK] Icono aplicado a {self.__class__.__name__}")
        else:
            print(f"[ERROR] No se encontró el icono para {self.__class__.__name__}")
    except Exception as e:
        print(f"[ERROR] Falló al aplicar icono a {self.__class__.__name__}: {e}")


ctk.CTk.__init__ = _patched_ctk_init

# Parche para CTkToplevel
_original_toplevel_init = ctk.CTkToplevel.__init__


def _patched_toplevel_init(self, *args, **kwargs):
    _original_toplevel_init(self, *args, **kwargs)
    icono_path = os.path.join(BASE_DIR, "assets", "images", "logo.ico")
    try:
        if os.path.exists(icono_path):
            self.after(100, lambda: self.iconbitmap(icono_path))
    except Exception as e:
        pass


ctk.CTkToplevel.__init__ = _patched_toplevel_init


def main():
    try:
        # Ejecutar debug primero
        debug_icono()

        print("\n" + "=" * 50)
        print("INICIANDO APLICACIÓN")
        print("=" * 50)

        app = LoginView()
        app.mainloop()
    except Exception as e:
        print(f"Error al iniciar la aplicación: {e}")
        import traceback
        traceback.print_exc()
        input("Presione Enter para salir...")


if __name__ == "__main__":
    main()