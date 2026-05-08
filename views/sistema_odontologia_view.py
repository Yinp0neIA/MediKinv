from utils.icon_manager import IconManager
from views.sistema_view import SistemaView

class SistemaOdontologiaView(SistemaView):
    """Ventana del sistema para Odontología - Hereda de SistemaView"""
    def __init__(self, usuario, area_id, area_nombre):
        super().__init__(usuario, area_id, area_nombre)
        IconManager.aplicar_icono(self)
        self.debug_print("Sistema de Odontología inicializado")