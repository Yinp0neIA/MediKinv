import tkinter as tk
import customtkinter as ctk
from tkinter import ttk
from assets.styles.themes import AppTheme
from models.movimiento import Movimiento
from utils.icon_manager import IconManager


class MovimientosView(ctk.CTkToplevel):
    """Ventana para ver historial de movimientos"""

    def __init__(self, medicamento_id, medicamento_nombre):
        super().__init__()
        IconManager.aplicar_icono(self)

        self.medicamento_id = medicamento_id
        self.medicamento_nombre = medicamento_nombre

        # Configurar ventana
        self.title(f"MediKinv - Movimientos de {medicamento_nombre}")
        self.geometry("900x500")

        # Centrar ventana
        self.centrar_ventana()

        # Crear interfaz
        self.crear_interfaz()

        # Cargar datos
        self.cargar_movimientos()

    def centrar_ventana(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.winfo_screenheight() // 2) - (500 // 2)
        self.geometry(f'900x500+{x}+{y}')

    def crear_interfaz(self):
        # Frame principal
        frame = ctk.CTkFrame(self, fg_color=AppTheme.SECONDARY, corner_radius=0)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Título
        label_titulo = ctk.CTkLabel(
            frame,
            text=f"Historial de Movimientos - {self.medicamento_nombre}",
            font=("Segoe UI", 18, "bold"),
            text_color=AppTheme.PRIMARY_DARK
        )
        label_titulo.pack(pady=20)

        # Tabla
        frame_tabla = ctk.CTkFrame(frame, fg_color="white", corner_radius=10)
        frame_tabla.pack(fill="both", expand=True, padx=20, pady=10)

        # Scrollbars
        scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical")
        scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal")

        # Treeview
        self.tabla = ttk.Treeview(
            frame_tabla,
            columns=("Fecha", "Tipo", "Cantidad", "Quien Recibe", "Quien Retira", "Motivo"),
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        scroll_y.config(command=self.tabla.yview)
        scroll_x.config(command=self.tabla.xview)

        # Configurar columnas
        self.tabla.heading("Fecha", text="Fecha")
        self.tabla.heading("Tipo", text="Tipo")
        self.tabla.heading("Cantidad", text="Cantidad")
        self.tabla.heading("Quien Recibe", text="Quien Recibe")
        self.tabla.heading("Quien Retira", text="Quien Retira")
        self.tabla.heading("Motivo", text="Motivo")

        self.tabla.column("Fecha", width=100, anchor="center")
        self.tabla.column("Tipo", width=80, anchor="center")
        self.tabla.column("Cantidad", width=80, anchor="center")
        self.tabla.column("Quien Recibe", width=150, anchor="w")
        self.tabla.column("Quien Retira", width=150, anchor="w")
        self.tabla.column("Motivo", width=300, anchor="w")

        # Empaquetar
        self.tabla.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)

        # Botón cerrar
        btn_cerrar = ctk.CTkButton(
            frame,
            text="CERRAR",
            width=150,
            height=35,
            font=("Segoe UI", 11, "bold"),
            fg_color="gray",
            hover_color="darkgray",
            command=self.destroy
        )
        btn_cerrar.pack(pady=20)

    def cargar_movimientos(self):
        """Cargar movimientos del medicamento"""
        movimientos = Movimiento.obtener_por_medicamento(self.medicamento_id)

        for mov in movimientos:
            fecha = mov['fecha'].strftime("%d/%m/%Y") if mov['fecha'] else ""
            tipo = "📥 Entrada" if mov['tipo'] == "entrada" else "📤 Salida"

            self.tabla.insert("", "end", values=(
                fecha,
                tipo,
                mov['cantidad'],
                mov['quien_recibe'],
                mov['quien_retira'],
                mov['motivo']
            ))