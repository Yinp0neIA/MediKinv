# views/login_view.py
import customtkinter as ctk
from tkinter import messagebox
from assets.styles.themes import AppTheme
from controllers.login_controller import LoginController
from utils.helpers import Helpers
from utils.icon_manager import IconManager


class LoginView(ctk.CTk):
    """Ventana de inicio de sesión - Diseño profesional"""

    def __init__(self):
        super().__init__()
        IconManager.aplicar_icono(self)

        # Configurar tema
        AppTheme.aplicar_tema()

        # Configurar ventana
        self.title("MediKinv - Sistema de Gestión de Medicamentos")
        self.geometry("1200x700")
        self.resizable(False, False)
        self.configure(fg_color=AppTheme.SECONDARY)

        # Centrar ventana
        Helpers.centrar_ventana(self, 1200, 700)

        # Inicializar controlador
        self.controller = LoginController(self)

        # Crear interfaz
        self.crear_interfaz()

        # Configurar cierre
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def crear_interfaz(self):
        """Crear la interfaz de usuario"""

        # Frame principal con dos columnas
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Panel izquierdo - Información/Logo
        self.panel_izquierdo = ctk.CTkFrame(
            self,
            fg_color=AppTheme.PRIMARY,
            corner_radius=0
        )
        self.panel_izquierdo.grid(row=0, column=0, sticky="nsew")
        self.crear_panel_izquierdo()

        # Panel derecho - Formulario de login
        self.panel_derecho = ctk.CTkFrame(
            self,
            fg_color=AppTheme.SECONDARY,
            corner_radius=0
        )
        self.panel_derecho.grid(row=0, column=1, sticky="nsew")
        self.crear_panel_derecho()

    def crear_panel_izquierdo(self):
        """Crear panel izquierdo con información"""
        contenedor = ctk.CTkFrame(
            self.panel_izquierdo,
            fg_color="transparent"
        )
        contenedor.pack(expand=True, fill="both", padx=50, pady=50)

        # Logo o ícono (opcional)
        try:
            logo = Helpers.cargar_imagen("assets/images/logo.png", (100, 100))
            if logo:
                lbl_logo = ctk.CTkLabel(contenedor, image=logo, text="")
                lbl_logo.pack(pady=(0, 20))
        except:
            pass

        # Título principal
        titulo = ctk.CTkLabel(
            contenedor,
            text="MediKinv",
            font=("Segoe UI", 56, "bold"),
            text_color="white"
        )
        titulo.pack(pady=(0, 15))

        # Subtítulo
        subtitulo = ctk.CTkLabel(
            contenedor,
            text="Sistema de Gestión de Medicamentos",
            font=("Segoe UI", 20),
            text_color="white",
            justify="center"
        )
        subtitulo.pack(pady=(0, 60))

        # Características reales
        caracteristicas = [
            "📦 Control de inventario de medicamentos",
            "⚠️ Alertas de medicamentos próximos a vencer",
            "📊 Reportes de stock y movimientos",
            "👥 Gestión de pacientes y tratamientos",
            "🔒 Acceso seguro con autenticación local"
        ]

        for texto in caracteristicas:
            lbl = ctk.CTkLabel(
                contenedor,
                text=texto,
                font=("Segoe UI", 14),
                text_color="white",
                justify="left"
            )
            lbl.pack(pady=12, anchor="w")

    def crear_panel_derecho(self):
        """Crear panel derecho con formulario de login"""
        contenedor = ctk.CTkFrame(
            self.panel_derecho,
            fg_color="transparent"
        )
        contenedor.pack(expand=True, fill="both", padx=80, pady=60)

        # Título
        titulo_form = ctk.CTkLabel(
            contenedor,
            text="Bienvenido",
            font=("Segoe UI", 40, "bold"),
            text_color=AppTheme.TEXT_PRIMARY
        )
        titulo_form.pack(pady=(0, 15))

        # Subtítulo
        subtitulo_form = ctk.CTkLabel(
            contenedor,
            text="Ingresa tus credenciales para acceder al sistema",
            font=("Segoe UI", 14),
            text_color=AppTheme.TEXT_SECONDARY
        )
        subtitulo_form.pack(pady=(0, 50))

        # Campo usuario
        lbl_usuario = ctk.CTkLabel(
            contenedor,
            text="USUARIO",
            font=("Segoe UI", 12, "bold"),
            text_color=AppTheme.TEXT_SECONDARY,
            anchor="w"
        )
        lbl_usuario.pack(fill="x", pady=(0, 5))

        self.txt_usuario = ctk.CTkEntry(
            contenedor,
            placeholder_text="Nombre de usuario",
            width=450,
            height=55,
            font=("Segoe UI", 14),
            corner_radius=12,
            border_width=2,
            border_color=AppTheme.ACCENT,
            fg_color="white",
            text_color=AppTheme.TEXT_PRIMARY
        )
        self.txt_usuario.pack(pady=(0, 25))

        # Campo contraseña
        lbl_contrasena = ctk.CTkLabel(
            contenedor,
            text="CONTRASEÑA",
            font=("Segoe UI", 12, "bold"),
            text_color=AppTheme.TEXT_SECONDARY,
            anchor="w"
        )
        lbl_contrasena.pack(fill="x", pady=(0, 5))

        self.txt_contrasena = ctk.CTkEntry(
            contenedor,
            placeholder_text="Contraseña",
            width=450,
            height=55,
            font=("Segoe UI", 14),
            show="●",
            corner_radius=12,
            border_width=2,
            border_color=AppTheme.ACCENT,
            fg_color="white",
            text_color=AppTheme.TEXT_PRIMARY
        )
        self.txt_contrasena.pack(pady=(0, 35))

        # Botón login
        self.btn_login = ctk.CTkButton(
            contenedor,
            text="INICIAR SESIÓN",
            width=450,
            height=60,
            font=("Segoe UI", 16, "bold"),
            fg_color=AppTheme.PRIMARY,
            hover_color=AppTheme.PRIMARY_DARK,
            corner_radius=12,
            command=self.controller.login
        )
        self.btn_login.pack(pady=(0, 25))

        # Separador
        separator_frame = ctk.CTkFrame(contenedor, fg_color="transparent")
        separator_frame.pack(fill="x", pady=20)

        separator = ctk.CTkFrame(
            separator_frame,
            height=2,
            fg_color=AppTheme.ACCENT
        )
        separator.pack(fill="x")

        # Frame para registro
        frame_registro = ctk.CTkFrame(contenedor, fg_color="transparent")
        frame_registro.pack(pady=20)

        lbl_no_cuenta = ctk.CTkLabel(
            frame_registro,
            text="¿No tienes una cuenta?",
            font=("Segoe UI", 13),
            text_color=AppTheme.TEXT_SECONDARY
        )
        lbl_no_cuenta.pack(side="left", padx=5)

        self.link_registro = ctk.CTkLabel(
            frame_registro,
            text="Regístrate aquí",
            font=("Segoe UI", 13, "bold"),
            text_color=AppTheme.PRIMARY,
            cursor="hand2"
        )
        self.link_registro.pack(side="left")
        self.link_registro.bind("<Button-1>", self.controller.abrir_registro)

        # Botón salir
        self.btn_salir = ctk.CTkButton(
            contenedor,
            text="SALIR",
            width=200,
            height=45,
            font=("Segoe UI", 13),
            fg_color="gray",
            hover_color="darkgray",
            corner_radius=10,
            command=self.on_closing
        )
        self.btn_salir.pack(pady=(15, 0))

        # Evento Enter
        self.txt_contrasena.bind("<Return>", lambda e: self.controller.login())
        self.txt_usuario.focus()

    def mostrar_error(self, mensaje):
        messagebox.showerror("Error", mensaje)

    def mostrar_exito(self, mensaje):
        messagebox.showinfo("Éxito", mensaje)

    def on_closing(self):
            self.quit()
            self.destroy()