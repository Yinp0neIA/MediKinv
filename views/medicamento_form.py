import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from assets.styles.themes import AppTheme
from models.medicamento import Medicamento
from utils.icon_manager import IconManager


class MedicamentoForm(ctk.CTkToplevel):
    """Formulario para crear/editar medicamentos"""

    def __init__(self, area_id, usuario, medicamento_id=None):
        super().__init__()
        IconManager.aplicar_icono(self)

        self.area_id = area_id
        self.usuario = usuario
        self.medicamento_id = medicamento_id
        self.es_edicion = medicamento_id is not None

        # Configurar ventana
        titulo = "Editar Medicamento" if self.es_edicion else "Nuevo Medicamento"
        self.title(f"MediKinv - {titulo}")
        self.geometry("600x700")
        self.resizable(False, False)

        # Centrar ventana
        self.centrar_ventana()

        # Crear interfaz
        self.crear_interfaz()

        # Si es edición, cargar datos
        if self.es_edicion:
            self.cargar_datos()

    def centrar_ventana(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.winfo_screenheight() // 2) - (700 // 2)
        self.geometry(f'600x700+{x}+{y}')

    def crear_interfaz(self):
        # Frame principal
        frame = ctk.CTkFrame(self, fg_color=AppTheme.SECONDARY, corner_radius=0)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Título
        titulo = "EDITAR MEDICAMENTO" if self.es_edicion else "NUEVO MEDICAMENTO"
        label_titulo = ctk.CTkLabel(
            frame,
            text=titulo,
            font=("Segoe UI", 20, "bold"),
            text_color=AppTheme.PRIMARY_DARK
        )
        label_titulo.pack(pady=(20, 30))

        # Frame para campos
        campos = ctk.CTkFrame(frame, fg_color="transparent")
        campos.pack(fill="both", expand=True, padx=40)

        # Nombre
        label_nombre = ctk.CTkLabel(campos, text="Nombre del medicamento", font=("Segoe UI", 12, "bold"))
        label_nombre.pack(anchor="w", pady=(10, 5))
        self.txt_nombre = ctk.CTkEntry(campos, width=400, height=40, font=("Segoe UI", 12))
        self.txt_nombre.pack(fill="x", pady=(0, 10))

        # Presentación
        label_presentacion = ctk.CTkLabel(campos, text="Presentación", font=("Segoe UI", 12, "bold"))
        label_presentacion.pack(anchor="w", pady=(10, 5))
        self.txt_presentacion = ctk.CTkEntry(campos, width=400, height=40, font=("Segoe UI", 12))
        self.txt_presentacion.pack(fill="x", pady=(0, 10))

        # Ubicación
        label_ubicacion = ctk.CTkLabel(campos, text="Ubicación", font=("Segoe UI", 12, "bold"))
        label_ubicacion.pack(anchor="w", pady=(10, 5))
        self.txt_ubicacion = ctk.CTkEntry(campos, width=400, height=40, font=("Segoe UI", 12))
        self.txt_ubicacion.pack(fill="x", pady=(0, 10))

        # Stock
        label_stock = ctk.CTkLabel(campos, text="Stock inicial", font=("Segoe UI", 12, "bold"))
        label_stock.pack(anchor="w", pady=(10, 5))
        self.txt_stock = ctk.CTkEntry(campos, width=400, height=40, font=("Segoe UI", 12))
        self.txt_stock.pack(fill="x", pady=(0, 10))

        # Caducidad
        label_caducidad = ctk.CTkLabel(campos, text="Fecha de caducidad (DD/MM/AAAA)", font=("Segoe UI", 12, "bold"))
        label_caducidad.pack(anchor="w", pady=(10, 5))
        self.txt_caducidad = ctk.CTkEntry(campos, width=400, height=40, font=("Segoe UI", 12))
        self.txt_caducidad.pack(fill="x", pady=(0, 10))

        # Es donación
        self.es_donacion = ctk.BooleanVar(value=False)
        check_donacion = ctk.CTkCheckBox(
            campos,
            text="¿Es una donación?",
            variable=self.es_donacion,
            font=("Segoe UI", 12),
            onvalue=True,
            offvalue=False,
            command=self.toggle_donante
        )
        check_donacion.pack(anchor="w", pady=(10, 5))

        # Donante (inicialmente oculto)
        self.frame_donante = ctk.CTkFrame(campos, fg_color="transparent")
        self.txt_donante = ctk.CTkEntry(self.frame_donante, width=400, height=40, font=("Segoe UI", 12))
        self.txt_donante.pack(fill="x")

        # Observaciones
        label_obs = ctk.CTkLabel(campos, text="Observaciones", font=("Segoe UI", 12, "bold"))
        label_obs.pack(anchor="w", pady=(10, 5))
        self.txt_observaciones = ctk.CTkTextbox(campos, width=400, height=80, font=("Segoe UI", 11))
        self.txt_observaciones.pack(fill="x", pady=(0, 10))

        # Botones
        frame_botones = ctk.CTkFrame(campos, fg_color="transparent")
        frame_botones.pack(fill="x", pady=20)

        btn_guardar = ctk.CTkButton(
            frame_botones,
            text="GUARDAR",
            width=150,
            height=40,
            font=("Segoe UI", 12, "bold"),
            fg_color=AppTheme.PRIMARY,
            hover_color=AppTheme.PRIMARY_DARK,
            command=self.guardar
        )
        btn_guardar.pack(side="left", padx=10)

        btn_cancelar = ctk.CTkButton(
            frame_botones,
            text="CANCELAR",
            width=150,
            height=40,
            font=("Segoe UI", 12),
            fg_color="gray",
            hover_color="darkgray",
            command=self.destroy
        )
        btn_cancelar.pack(side="left", padx=10)

    def toggle_donante(self):
        """Mostrar u ocultar campo donante"""
        if self.es_donacion.get():
            self.frame_donante.pack(fill="x", pady=(0, 10))
        else:
            self.frame_donante.pack_forget()

    def cargar_datos(self):
        """Cargar datos del medicamento para edición"""
        medicamento = Medicamento.obtener_por_id(self.medicamento_id)
        if medicamento:
            self.txt_nombre.insert(0, medicamento['nombre'])
            self.txt_presentacion.insert(0, medicamento['presentacion'])
            self.txt_ubicacion.insert(0, medicamento['ubicacion'])
            self.txt_stock.insert(0, str(medicamento['stock']))
            if medicamento['caducidad']:
                self.txt_caducidad.insert(0, medicamento['caducidad'].strftime("%d/%m/%Y"))
            self.es_donacion.set(medicamento['es_donacion'])
            if medicamento['es_donacion'] and medicamento['donante']:
                self.frame_donante.pack(fill="x", pady=(0, 10))
                self.txt_donante.insert(0, medicamento['donante'])
            if medicamento['observaciones']:
                self.txt_observaciones.insert("1.0", medicamento['observaciones'])

    def guardar(self):
        """Guardar medicamento"""
        nombre = self.txt_nombre.get().strip()
        presentacion = self.txt_presentacion.get().strip()
        ubicacion = self.txt_ubicacion.get().strip()
        stock = self.txt_stock.get().strip()
        caducidad_str = self.txt_caducidad.get().strip()
        observaciones = self.txt_observaciones.get("1.0", "end-1c").strip()
        es_donacion = self.es_donacion.get()
        donante = self.txt_donante.get().strip() if es_donacion else ""

        # Validaciones
        if not nombre:
            messagebox.showerror("Error", "Ingrese el nombre del medicamento")
            return

        if not presentacion:
            messagebox.showerror("Error", "Ingrese la presentación")
            return

        if not stock:
            messagebox.showerror("Error", "Ingrese el stock")
            return

        try:
            stock_int = int(stock)
        except ValueError:
            messagebox.showerror("Error", "El stock debe ser un número")
            return

        # Procesar fecha
        caducidad = None
        if caducidad_str:
            try:
                caducidad = datetime.strptime(caducidad_str, "%d/%m/%Y").date()
            except ValueError:
                messagebox.showerror("Error", "Formato de fecha inválido. Use DD/MM/AAAA")
                return

        # Guardar
        if self.es_edicion:
            success = Medicamento.actualizar(
                self.medicamento_id, nombre, presentacion, ubicacion,
                caducidad, observaciones, stock_int, es_donacion, donante
            )
        else:
            success = Medicamento.crear(
                self.area_id, nombre, presentacion, ubicacion,
                caducidad, observaciones, stock_int, es_donacion, donante
            )

        if success:
            messagebox.showinfo("Éxito", "Medicamento guardado correctamente")
            self.destroy()
        else:
            messagebox.showerror("Error", "No se pudo guardar el medicamento")