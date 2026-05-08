import customtkinter as ctk
from tkinter import messagebox
from assets.styles.themes import AppTheme
from utils.helpers import Helpers
from utils.icon_manager import IconManager


class AreaSelectView(ctk.CTk):

    def __init__(self, usuario):
        super().__init__()
        IconManager.aplicar_icono(self)

        AppTheme.aplicar_tema()

        self.usuario = usuario

        self.title("MediKinv - Seleccionar Área")
        self.geometry("1200x700")
        self.resizable(False, False)
        self.configure(fg_color=AppTheme.SECONDARY)

        Helpers.centrar_ventana(self, 1200, 700)

        self.crear_interfaz()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def crear_interfaz(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.panel_izquierdo = ctk.CTkFrame(
            self,
            fg_color=AppTheme.PRIMARY,
            corner_radius=0
        )
        self.panel_izquierdo.grid(row=0, column=0, sticky="nsew")
        self.crear_panel_izquierdo()

        self.panel_derecho = ctk.CTkFrame(
            self,
            fg_color=AppTheme.SECONDARY,
            corner_radius=0
        )
        self.panel_derecho.grid(row=0, column=1, sticky="nsew")
        self.crear_panel_derecho()

    def crear_panel_izquierdo(self):
        contenedor = ctk.CTkFrame(
            self.panel_izquierdo,
            fg_color="transparent"
        )
        contenedor.pack(expand=True, fill="both", padx=45, pady=40)

        try:
            self.iconbitmap("assets/images/logo.ico")
            logo = Helpers.cargar_imagen("assets/images/logo.png", (80, 80))
            if logo:
                lbl_logo = ctk.CTkLabel(contenedor, image=logo, text="")
                lbl_logo.pack(pady=(0, 10))
        except:
            pass

        titulo = ctk.CTkLabel(
            contenedor,
            text="MediKinv",
            font=("Segoe UI", 48, "bold"),
            text_color="white"
        )
        titulo.pack(pady=(0, 8))

        subtitulo = ctk.CTkLabel(
            contenedor,
            text="Sistema de Gestión de Medicamentos",
            font=("Segoe UI", 16),
            text_color="white",
            justify="center"
        )
        subtitulo.pack(pady=(0, 35))

        frame_usuario = ctk.CTkFrame(
            contenedor,
            fg_color="white",
            corner_radius=15,
            border_width=2,
            border_color=AppTheme.ACCENT
        )
        frame_usuario.pack(pady=20, fill="x")

        lbl_usuario_titulo = ctk.CTkLabel(
            frame_usuario,
            text="BIENVENIDO",
            font=("Segoe UI", 12, "bold"),
            text_color=AppTheme.PRIMARY,
            justify="center"
        )
        lbl_usuario_titulo.pack(pady=(20, 5))

        lbl_usuario_nombre = ctk.CTkLabel(
            frame_usuario,
            text=f"👤 {self.usuario.nombre_usuario}",
            font=("Segoe UI", 18, "bold"),
            text_color=AppTheme.TEXT_PRIMARY,
            justify="center"
        )
        lbl_usuario_nombre.pack(pady=(0, 20))

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
        contenedor = ctk.CTkFrame(
            self.panel_derecho,
            fg_color="transparent"
        )
        contenedor.pack(expand=True, fill="both", padx=70, pady=35)

        titulo_form = ctk.CTkLabel(
            contenedor,
            text="Seleccionar Área",
            font=("Segoe UI", 34, "bold"),
            text_color=AppTheme.TEXT_PRIMARY
        )
        titulo_form.pack(pady=(0, 5))

        subtitulo_form = ctk.CTkLabel(
            contenedor,
            text="Elige el módulo al que deseas acceder",
            font=("Segoe UI", 12),
            text_color=AppTheme.TEXT_SECONDARY
        )
        subtitulo_form.pack(pady=(0, 35))

        frame_areas = ctk.CTkFrame(contenedor, fg_color="transparent")
        frame_areas.pack(expand=True, fill="both", pady=20)

        frame_areas.grid_columnconfigure(0, weight=1)
        frame_areas.grid_columnconfigure(1, weight=1)
        frame_areas.grid_columnconfigure(2, weight=1)

        frame_areas.grid_rowconfigure(0, weight=1)

        self.crear_tarjeta_area(
            frame_areas,
            "Farmacia",
            "Gestión de inventario de medicamentos\nControl de fechas de vencimiento\nAlertas de stock bajo",
            "💊",
            1,
            0,
            0
        )

        self.crear_tarjeta_area(
            frame_areas,
            "Odontología",
            "Gestión de insumos odontológicos\nControl de materiales dentales\nRegistro de tratamientos",
            "🦷",
            2,
            0,
            1
        )

        self.crear_tarjeta_area(
            frame_areas,
            "Catálogo",
            "Administración de catálogo\nGestor de medicamentos\nControl de productos",
            "📋",
            3,
            0,
            2
        )

        frame_botones = ctk.CTkFrame(contenedor, fg_color="transparent")
        frame_botones.pack(pady=30)

        btn_cerrar_sesion = ctk.CTkButton(
            frame_botones,
            text="CERRAR SESIÓN",
            width=250,
            height=45,
            font=("Segoe UI", 13, "bold"),
            fg_color=AppTheme.PRIMARY,
            hover_color=AppTheme.PRIMARY_DARK,
            corner_radius=10,
            command=self.cerrar_sesion
        )
        btn_cerrar_sesion.pack()

        lbl_texto = ctk.CTkLabel(
            frame_botones,
            text="Al cerrar sesión volverás a la pantalla de inicio",
            font=("Segoe UI", 10),
            text_color=AppTheme.TEXT_SECONDARY
        )
        lbl_texto.pack(pady=(10, 0))

    def crear_tarjeta_area(self, parent, titulo, descripcion, icono, area_id, row, column):
        card = ctk.CTkFrame(
            parent,
            fg_color=AppTheme.SURFACE,
            corner_radius=15,
            border_width=2,
            border_color=AppTheme.ACCENT
        )
        card.grid(row=row, column=column, padx=15, pady=20, sticky="nsew")

        card.grid_propagate(False)
        card.configure(height=380, width=300)

        contenido_card = ctk.CTkFrame(card, fg_color="transparent")
        contenido_card.pack(expand=True, fill="both", padx=20, pady=20)

        icono_label = ctk.CTkLabel(
            contenido_card,
            text=icono,
            font=("Segoe UI", 60),
            text_color=AppTheme.PRIMARY
        )
        icono_label.pack(pady=(20, 10))

        label_titulo = ctk.CTkLabel(
            contenido_card,
            text=titulo,
            font=("Segoe UI", 22, "bold"),
            text_color=AppTheme.PRIMARY
        )
        label_titulo.pack(pady=(10, 15))

        label_desc = ctk.CTkLabel(
            contenido_card,
            text=descripcion,
            font=("Segoe UI", 11),
            text_color=AppTheme.TEXT_SECONDARY,
            justify="center",
            wraplength=260
        )
        label_desc.pack(pady=(0, 25))

        btn_ingresar = ctk.CTkButton(
            contenido_card,
            text="INGRESAR",
            width=180,
            height=45,
            font=("Segoe UI", 13, "bold"),
            fg_color=AppTheme.PRIMARY,
            hover_color=AppTheme.PRIMARY_DARK,
            corner_radius=10,
            command=lambda: self.seleccionar_area(area_id, titulo)
        )
        btn_ingresar.pack(pady=(10, 10))

    def seleccionar_area(self, area_id, area_nombre):
        if area_id == 3:
            self.abrir_catalogo()
        else:
            self.destroy()
            try:
                from views.sistema_view import SistemaView
                sistema = SistemaView(self.usuario, area_id, area_nombre)
                sistema.mainloop()
            except Exception as e:
                print(f"Error al abrir sistema: {e}")
                messagebox.showerror("Error", f"No se pudo abrir el sistema: {e}")

    def abrir_catalogo(self):
        try:
            from views.catalogo_medicamentos_view import CatalogoMedicamentosView
            catalogo = CatalogoMedicamentosView(self, self.usuario)
            self.destroy()
            catalogo.mainloop()
        except Exception as e:
            print(f"Error al abrir catálogo: {e}")
            messagebox.showerror("Error", f"No se pudo abrir el catálogo: {e}")

    def cerrar_sesion(self):
        self.destroy()
        try:
            from views.login_view import LoginView
            login = LoginView()
            login.mainloop()
        except Exception as e:
            print(f"Error al cerrar sesión: {e}")

    def on_closing(self):
        self.quit()
        self.destroy()