# controllers/login_controller.py
from models.usuario import Usuario
from views.registro_view import RegistroView
from views.area_select_view import AreaSelectView


class LoginController:
    def __init__(self, view):
        self.view = view

    def login(self):
        usuario = self.view.txt_usuario.get().strip()
        contrasena = self.view.txt_contrasena.get()

        if not usuario or not contrasena:
            self.view.mostrar_error("Complete todos los campos")
            return

        # Deshabilitar botón mientras se procesa
        self.view.btn_login.configure(state="disabled", text="Validando...")
        self.view.update()

        try:
            usuario_obj = Usuario.validar_credenciales(usuario, contrasena)

            if usuario_obj:
                # Cancelar after pendientes antes de destruir
                try:
                    for after_id in self.view.tk.eval('after info').split():
                        try:
                            self.view.after_cancel(int(after_id))
                        except:
                            pass
                except:
                    pass

                # Guardar referencia a la ventana actual
                current_view = self.view
                # Crear la nueva ventana
                area_view = AreaSelectView(usuario_obj)
                # Destruir la actual después de que la nueva esté lista
                current_view.destroy()
                area_view.mainloop()
            else:
                self.view.mostrar_error("Usuario o contraseña incorrectos")
                self.view.txt_contrasena.delete(0, "end")
                self.view.txt_contrasena.focus()
                self.view.btn_login.configure(state="normal", text="INICIAR SESIÓN")
        except Exception as e:
            self.view.mostrar_error(f"Error: {str(e)}")
            self.view.btn_login.configure(state="normal", text="INICIAR SESIÓN")

    def abrir_registro(self, event=None):
        """Abrir ventana de registro - CERRAR el login"""
        try:
            # Cancelar after pendientes
            try:
                for after_id in self.view.tk.eval('after info').split():
                    try:
                        self.view.after_cancel(int(after_id))
                    except:
                        pass
            except:
                pass

            # Guardar referencia a la ventana actual
            current_view = self.view

            # Crear nueva ventana de registro
            registro_view = RegistroView()

            # Destruir el login
            current_view.destroy()

            # Iniciar la ventana de registro
            registro_view.mainloop()

        except Exception as e:
            print(f"Error al abrir registro: {e}")
            if hasattr(self.view, 'mostrar_error'):
                self.view.mostrar_error(f"No se pudo abrir el registro: {e}")