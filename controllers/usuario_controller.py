# controllers/usuario_controller.py
from models.usuario import Usuario


class UsuarioController:
    """Controlador para operaciones de usuario"""

    def __init__(self, view):
        self.view = view

    def registrar(self, usuario, contrasena, confirmar):
        """Procesar el registro de un nuevo usuario"""

        try:
            # Registrar usuario
            success, mensaje = Usuario.registrar_usuario(usuario, contrasena)

            if success:
                self.view.mostrar_exito(mensaje)
                # El éxito maneja el volver al login
            else:
                self.view.mostrar_error(mensaje)
                self.view.txt_usuario.delete(0, "end")
                self.view.txt_usuario.focus()
                self.view.btn_registrar.configure(state="normal", text="CREAR CUENTA")

        except Exception as e:
            self.view.mostrar_error(f"Error al registrar: {str(e)}")
            self.view.btn_registrar.configure(state="normal", text="CREAR CUENTA")