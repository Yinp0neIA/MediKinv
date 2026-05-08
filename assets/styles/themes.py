import customtkinter as ctk


class AppTheme:
    """Paleta de colores principal - Rojo fuerte #7C281E"""

    # Colores primarios
    PRIMARY = "#7C281E"  # Rojo fuerte principal
    PRIMARY_DARK = "#5A1D15"  # Rojo más oscuro
    PRIMARY_LIGHT = "#C34424"  # Naranja/rojo claro para acentos

    # Colores secundarios
    SECONDARY = "#E0CFC7"  # Beige claro
    BACKGROUND = "#E0CFC7"  # Beige claro fondo
    SURFACE = "#FFFFFF"  # Blanco

    # Colores de texto
    TEXT_PRIMARY = "#2C3E50"
    TEXT_SECONDARY = "#7F8C8D"
    TEXT_DARK = "#2C3E50"
    TEXT_LIGHT = "#FFFFFF"
    TEXT_ON_PRIMARY = "#FFFFFF"

    # Colores de estado
    SUCCESS = "#27AE60"
    WARNING = "#F39C12"
    ERROR = "#E74C3C"
    INFO = "#3498DB"
    ACCENT = "#C34424"

    # Colores para tabla
    TABLE_HEADER_BG = "#7C281E"
    TABLE_HEADER_FG = "#FFFFFF"
    TABLE_ROW_EVEN = "#F9F9F9"
    TABLE_ROW_ODD = "#FFFFFF"
    TABLE_SELECTED = "#E0CFC7"

    # Bordes
    BORDER_RADIUS_SMALL = 8
    BORDER_RADIUS_MEDIUM = 12
    BORDER_RADIUS_LARGE = 15

    @classmethod
    def aplicar_tema(cls):
        """Aplicar tema global a la aplicación"""
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # Configuración de widgets
        ctk.set_widget_scaling(1.0)
        ctk.set_window_scaling(1.0)