# views/registro_view.py
import customtkinter as ctk
from tkinter import messagebox
from assets.styles.themes import AppTheme
from controllers.usuario_controller import UsuarioController
from utils.helpers import Helpers
from utils.icon_manager import IconManager


class RegistroView(ctk.CTk):
    """Ventana de registro de usuarios - Diseño profesional"""

    def __init__(self):
        super().__init__()
        IconManager.aplicar_icono(self)

        # Configurar tema
        AppTheme.aplicar_tema()

        # Configurar ventana
        self.title("MediKinv - Crear Cuenta")
        self.geometry("1200x700")
        self.resizable(False, False)
        self.configure(fg_color=AppTheme.SECONDARY)

        # Centrar ventana
        Helpers.centrar_ventana(self, 1200, 700)

        # Inicializar controlador
        self.controller = UsuarioController(self)

        # Variable para controlar si se registró exitosamente
        self.registro_exitoso = False

        # Crear interfaz
        self.crear_interfaz()

        # Configurar cierre - Ahora sin confirmación
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def crear_interfaz(self):
        """Crear la interfaz de usuario"""

        # Frame principal con dos columnas
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Panel izquierdo - Información
        self.panel_izquierdo = ctk.CTkFrame(
            self,
            fg_color=AppTheme.PRIMARY,
            corner_radius=0
        )
        self.panel_izquierdo.grid(row=0, column=0, sticky="nsew")
        self.crear_panel_izquierdo()

        # Panel derecho - Formulario de registro
        self.panel_derecho = ctk.CTkFrame(
            self,
            fg_color=AppTheme.SECONDARY,
            corner_radius=0
        )
        self.panel_derecho.grid(row=0, column=1, sticky="nsew")
        self.crear_panel_derecho()

    def crear_panel_izquierdo(self):
        """Panel izquierdo con información"""
        contenedor = ctk.CTkFrame(
            self.panel_izquierdo,
            fg_color="transparent"
        )
        contenedor.pack(expand=True, fill="both", padx=45, pady=40)

        # Logo o ícono (opcional)
        try:
            logo = Helpers.cargar_imagen("assets/images/logo.png", (80, 80))
            if logo:
                lbl_logo = ctk.CTkLabel(contenedor, image=logo, text="")
                lbl_logo.pack(pady=(0, 10))
        except:
            pass

        # Título
        titulo = ctk.CTkLabel(
            contenedor,
            text="MediKinv",
            font=("Segoe UI", 48, "bold"),
            text_color="white"
        )
        titulo.pack(pady=(0, 8))

        # Subtítulo
        subtitulo = ctk.CTkLabel(
            contenedor,
            text="Sistema de Gestión de Medicamentos",
            font=("Segoe UI", 16),
            text_color="white",
            justify="center"
        )
        subtitulo.pack(pady=(0, 35))

        # Beneficios reales
        beneficios = [
            "✓ Gestión completa de inventario",
            "✓ Control de fechas de vencimiento",
            "✓ Registro de pacientes",
            "✓ Reportes personalizados",
            "✓ Datos almacenados localmente"
        ]

        for texto in beneficios:
            lbl = ctk.CTkLabel(
                contenedor,
                text=texto,
                font=("Segoe UI", 12),
                text_color="white",
                justify="left"
            )
            lbl.pack(pady=8, anchor="w")

    def crear_panel_derecho(self):
        """Panel derecho con formulario de registro - Todo visible"""

        # Contenedor principal
        contenedor = ctk.CTkFrame(
            self.panel_derecho,
            fg_color="transparent"
        )
        contenedor.pack(expand=True, fill="both", padx=70, pady=35)

        # Título
        titulo_form = ctk.CTkLabel(
            contenedor,
            text="Crear Cuenta",
            font=("Segoe UI", 34, "bold"),
            text_color=AppTheme.TEXT_PRIMARY
        )
        titulo_form.pack(pady=(0, 5))

        # Subtítulo
        subtitulo_form = ctk.CTkLabel(
            contenedor,
            text="Completa los siguientes datos para registrarte",
            font=("Segoe UI", 12),
            text_color=AppTheme.TEXT_SECONDARY
        )
        subtitulo_form.pack(pady=(0, 25))

        # Campo nombre de usuario
        lbl_usuario = ctk.CTkLabel(
            contenedor,
            text="NOMBRE DE USUARIO",
            font=("Segoe UI", 11, "bold"),
            text_color=AppTheme.TEXT_SECONDARY,
            anchor="w"
        )
        lbl_usuario.pack(fill="x", pady=(0, 3))

        self.txt_usuario = ctk.CTkEntry(
            contenedor,
            placeholder_text="Mínimo 3 caracteres",
            width=430,
            height=45,
            font=("Segoe UI", 13),
            corner_radius=10,
            border_width=2,
            border_color=AppTheme.ACCENT,
            fg_color="white",
            text_color=AppTheme.TEXT_PRIMARY
        )
        self.txt_usuario.pack(pady=(0, 15))

        # Campo contraseña
        lbl_contrasena = ctk.CTkLabel(
            contenedor,
            text="CONTRASEÑA",
            font=("Segoe UI", 11, "bold"),
            text_color=AppTheme.TEXT_SECONDARY,
            anchor="w"
        )
        lbl_contrasena.pack(fill="x", pady=(0, 3))

        self.txt_contrasena = ctk.CTkEntry(
            contenedor,
            placeholder_text="Mínimo 4 caracteres",
            width=430,
            height=45,
            font=("Segoe UI", 13),
            show="●",
            corner_radius=10,
            border_width=2,
            border_color=AppTheme.ACCENT,
            fg_color="white",
            text_color=AppTheme.TEXT_PRIMARY
        )
        self.txt_contrasena.pack(pady=(0, 5))

        # Indicador visual de fuerza de contraseña
        frame_indicador = ctk.CTkFrame(contenedor, fg_color="transparent")
        frame_indicador.pack(fill="x", pady=(0, 2))

        # Barra de progreso de fuerza
        self.barra_fuerza = ctk.CTkProgressBar(
            frame_indicador,
            width=430,
            height=6,
            corner_radius=3,
            progress_color="#FF6B6B",
            fg_color="#E0E0E0"
        )
        self.barra_fuerza.pack()
        self.barra_fuerza.set(0)

        # Texto de fuerza
        self.lbl_fuerza = ctk.CTkLabel(
            contenedor,
            text="⚪ Ingresa una contraseña",
            font=("Segoe UI", 10),
            text_color="#6B6B6B"
        )
        self.lbl_fuerza.pack(pady=(4, 10))

        # Campo confirmar contraseña
        lbl_confirmar = ctk.CTkLabel(
            contenedor,
            text="CONFIRMAR CONTRASEÑA",
            font=("Segoe UI", 11, "bold"),
            text_color=AppTheme.TEXT_SECONDARY,
            anchor="w"
        )
        lbl_confirmar.pack(fill="x", pady=(0, 3))

        self.txt_confirmar = ctk.CTkEntry(
            contenedor,
            placeholder_text="Repite tu contraseña",
            width=430,
            height=45,
            font=("Segoe UI", 13),
            show="●",
            corner_radius=10,
            border_width=2,
            border_color=AppTheme.ACCENT,
            fg_color="white",
            text_color=AppTheme.TEXT_PRIMARY
        )
        self.txt_confirmar.pack(pady=(0, 6))

        # Mensaje de validación de contraseñas
        self.lbl_validacion = ctk.CTkLabel(
            contenedor,
            text="",
            font=("Segoe UI", 10),
            text_color="green"
        )
        self.lbl_validacion.pack(pady=(0, 12))

        # Botón registrar
        self.btn_registrar = ctk.CTkButton(
            contenedor,
            text="CREAR CUENTA",
            width=430,
            height=48,
            font=("Segoe UI", 14, "bold"),
            fg_color=AppTheme.PRIMARY,
            hover_color=AppTheme.PRIMARY_DARK,
            corner_radius=10,
            command=self.registrar
        )
        self.btn_registrar.pack(pady=(0, 12))

        # Separador
        separator = ctk.CTkFrame(
            contenedor,
            height=1,
            fg_color=AppTheme.ACCENT
        )
        separator.pack(fill="x", pady=12)

        # Frame para volver al login
        frame_login = ctk.CTkFrame(contenedor, fg_color="transparent")
        frame_login.pack(pady=5)

        lbl_ya_tienes = ctk.CTkLabel(
            frame_login,
            text="¿Ya tienes una cuenta?",
            font=("Segoe UI", 12),
            text_color=AppTheme.TEXT_SECONDARY
        )
        lbl_ya_tienes.pack(side="left", padx=5)

        self.link_login = ctk.CTkLabel(
            frame_login,
            text="Inicia sesión aquí",
            font=("Segoe UI", 12, "bold"),
            text_color=AppTheme.PRIMARY,
            cursor="hand2"
        )
        self.link_login.pack(side="left")
        self.link_login.bind("<Button-1>", lambda e: self.volver_login())

        # Eventos para validación en tiempo real
        self.txt_contrasena.bind("<KeyRelease>", self.validar_contrasena_tiempo_real)
        self.txt_confirmar.bind("<KeyRelease>", self.validar_confirmacion)
        self.txt_usuario.bind("<KeyRelease>", self.validar_usuario)

        # Enfocar primer campo
        self.txt_usuario.focus()

    def validar_usuario(self, event=None):
        """Validar nombre de usuario en tiempo real"""
        usuario = self.txt_usuario.get()
        if usuario and len(usuario) < 3:
            self.txt_usuario.configure(border_color="#FF6B6B")
        elif len(usuario) >= 3:
            self.txt_usuario.configure(border_color="#4CAF50")
        else:
            self.txt_usuario.configure(border_color=AppTheme.ACCENT)

    def validar_contrasena_tiempo_real(self, event=None):
        """Validar fuerza de contraseña con barra de progreso e íconos universales"""
        contrasena = self.txt_contrasena.get()
        longitud = len(contrasena)

        # Determinar fuerza y color con íconos universales
        if longitud == 0:
            self.barra_fuerza.set(0)
            self.barra_fuerza.configure(progress_color="#FF6B6B")
            self.lbl_fuerza.configure(text="⚪ Ingresa una contraseña", text_color="#6B6B6B")
            self.txt_contrasena.configure(border_color=AppTheme.ACCENT)
        elif longitud < 4:
            self.barra_fuerza.set(0.33)
            self.barra_fuerza.configure(progress_color="#FF6B6B")
            self.lbl_fuerza.configure(text="❌ Débil - Mínimo 4 caracteres", text_color="#FF6B6B")
            self.txt_contrasena.configure(border_color="#FF6B6B")
        elif 4 <= longitud < 6:
            self.barra_fuerza.set(0.66)
            self.barra_fuerza.configure(progress_color="#FFA500")
            self.lbl_fuerza.configure(text="⚠️ Media - Buena", text_color="#FFA500")
            self.txt_contrasena.configure(border_color="#FFA500")
        else:
            self.barra_fuerza.set(1.0)
            self.barra_fuerza.configure(progress_color="#4CAF50")
            self.lbl_fuerza.configure(text="✅ Fuerte - Excelente", text_color="#4CAF50")
            self.txt_contrasena.configure(border_color="#4CAF50")

        # Validar confirmación si ya hay texto
        if self.txt_confirmar.get():
            self.validar_confirmacion()

    def validar_confirmacion(self, event=None):
        """Validar que las contraseñas coincidan"""
        contrasena = self.txt_contrasena.get()
        confirmar = self.txt_confirmar.get()

        if confirmar:
            if contrasena == confirmar:
                self.lbl_validacion.configure(text="✓ Las contraseñas coinciden", text_color="#4CAF50")
                self.txt_confirmar.configure(border_color="#4CAF50")
            else:
                self.lbl_validacion.configure(text="✗ Las contraseñas no coinciden", text_color="#FF6B6B")
                self.txt_confirmar.configure(border_color="#FF6B6B")
        else:
            self.lbl_validacion.configure(text="")
            self.txt_confirmar.configure(border_color=AppTheme.ACCENT)

    def registrar(self):
        """Registrar nuevo usuario con validaciones"""
        usuario = self.txt_usuario.get().strip()
        contrasena = self.txt_contrasena.get()
        confirmar = self.txt_confirmar.get()

        # Validaciones
        if not usuario:
            self.mostrar_error("Por favor, ingresa un nombre de usuario")
            self.txt_usuario.focus()
            return

        if len(usuario) < 3:
            self.mostrar_error("El nombre de usuario debe tener al menos 3 caracteres")
            self.txt_usuario.focus()
            return

        if not contrasena:
            self.mostrar_error("Por favor, ingresa una contraseña")
            self.txt_contrasena.focus()
            return

        if len(contrasena) < 4:
            self.mostrar_error("La contraseña debe tener al menos 4 caracteres")
            self.txt_contrasena.focus()
            return

        if contrasena != confirmar:
            self.mostrar_error("Las contraseñas no coinciden")
            self.txt_confirmar.focus()
            return

        # Deshabilitar botón mientras procesa
        self.btn_registrar.configure(state="disabled", text="Registrando...")
        self.update()

        # Llamar al controlador
        self.controller.registrar(usuario, contrasena, confirmar)

    def mostrar_error(self, mensaje):
        messagebox.showerror("Error", mensaje)
        self.btn_registrar.configure(state="normal", text="CREAR CUENTA")

    def mostrar_exito(self, mensaje):
        messagebox.showinfo("Éxito", mensaje)
        self.registro_exitoso = True
        # Después del éxito, volver al login después de 1.5 segundos
        self.after(1500, self.volver_login)

    def volver_login(self):
        """Volver a la ventana de login"""
        try:
            from views.login_view import LoginView
            login = LoginView()
            self.destroy()
            login.mainloop()
        except Exception as e:
            print(f"Error al volver al login: {e}")

    def on_closing(self):
        """Cerrar completamente la aplicación sin confirmación"""
        self.quit()
        self.destroy()