import customtkinter as ctk
from tkinter import messagebox
from assets.styles.themes import AppTheme
from models.movimiento import Movimiento
from models.medicamento import Medicamento
from utils.icon_manager import IconManager


class MovimientoForm(ctk.CTkToplevel):
    """Formulario para registrar movimientos"""

    def __init__(self, medicamento_id, medicamento_nombre, usuario):
        super().__init__()
        IconManager.aplicar_icono(self)

        self.medicamento_id = medicamento_id
        self.medicamento_nombre = medicamento_nombre
        self.usuario = usuario

        # Configurar ventana
        self.title("MediKinv - Registrar Movimiento")
        self.geometry("500x550")
        self.resizable(False, False)

        # Centrar ventana
        self.centrar_ventana()

        # Crear interfaz
        self.crear_interfaz()

    def centrar_ventana(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.winfo_screenheight() // 2) - (550 // 2)
        self.geometry(f'500x550+{x}+{y}')

    def crear_interfaz(self):
        # Frame principal
        frame = ctk.CTkFrame(self, fg_color=AppTheme.SECONDARY, corner_radius=0)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Título
        label_titulo = ctk.CTkLabel(
            frame,
            text="REGISTRAR MOVIMIENTO",
            font=("Segoe UI", 20, "bold"),
            text_color=AppTheme.PRIMARY_DARK
        )
        label_titulo.pack(pady=(20, 10))

        # Medicamento
        label_med = ctk.CTkLabel(
            frame,
            text=f"Medicamento: {self.medicamento_nombre}",
            font=("Segoe UI", 12),
            text_color=AppTheme.TEXT_DARK
        )
        label_med.pack(pady=(10, 20))

        # Frame para campos
        campos = ctk.CTkFrame(frame, fg_color="transparent")
        campos.pack(fill="both", expand=True, padx=30)

        # Tipo de movimiento
        label_tipo = ctk.CTkLabel(campos, text="Tipo de movimiento", font=("Segoe UI", 12, "bold"))
        label_tipo.pack(anchor="w", pady=(10, 5))

        self.tipo_movimiento = ctk.StringVar(value="entrada")
        radio_entrada = ctk.CTkRadioButton(
            campos,
            text="Entrada (Ingreso)",
            variable=self.tipo_movimiento,
            value="entrada",
            font=("Segoe UI", 11)
        )
        radio_entrada.pack(anchor="w", pady=5)

        radio_salida = ctk.CTkRadioButton(
            campos,
            text="Salida (Retiro)",
            variable=self.tipo_movimiento,
            value="salida",
            font=("Segoe UI", 11)
        )
        radio_salida.pack(anchor="w", pady=5)

        # Cantidad
        label_cantidad = ctk.CTkLabel(campos, text="Cantidad", font=("Segoe UI", 12, "bold"))
        label_cantidad.pack(anchor="w", pady=(15, 5))
        self.txt_cantidad = ctk.CTkEntry(campos, width=350, height=40, font=("Segoe UI", 12))
        self.txt_cantidad.pack(fill="x", pady=(0, 10))

        # Quien recibe (para entradas)
        self.frame_recibe = ctk.CTkFrame(campos, fg_color="transparent")
        label_recibe = ctk.CTkLabel(self.frame_recibe, text="¿Quién recibe?", font=("Segoe UI", 12, "bold"))
        label_recibe.pack(anchor="w", pady=(10, 5))
        self.txt_recibe = ctk.CTkEntry(self.frame_recibe, width=350, height=40, font=("Segoe UI", 12))
        self.txt_recibe.pack(fill="x", pady=(0, 5))

        # Quien retira (para salidas)
        self.frame_retira = ctk.CTkFrame(campos, fg_color="transparent")
        label_retira = ctk.CTkLabel(self.frame_retira, text="¿Quién retira?", font=("Segoe UI", 12, "bold"))
        label_retira.pack(anchor="w", pady=(10, 5))
        self.txt_retira = ctk.CTkEntry(self.frame_retira, width=350, height=40, font=("Segoe UI", 12))
        self.txt_retira.pack(fill="x", pady=(0, 5))

        # Motivo
        label_motivo = ctk.CTkLabel(campos, text="Motivo", font=("Segoe UI", 12, "bold"))
        label_motivo.pack(anchor="w", pady=(10, 5))
        self.txt_motivo = ctk.CTkEntry(campos, width=350, height=40, font=("Segoe UI", 12))
        self.txt_motivo.pack(fill="x", pady=(0, 10))

        # Mostrar campos según tipo
        self.frame_recibe.pack(fill="x")
        self.frame_retira.pack_forget()

        # Evento para cambiar campos según tipo
        def on_tipo_change(*args):
            if self.tipo_movimiento.get() == "entrada":
                self.frame_recibe.pack(fill="x")
                self.frame_retira.pack_forget()
            else:
                self.frame_recibe.pack_forget()
                self.frame_retira.pack(fill="x")

        self.tipo_movimiento.trace_add("write", on_tipo_change)

        # Botones
        frame_botones = ctk.CTkFrame(campos, fg_color="transparent")
        frame_botones.pack(fill="x", pady=20)

        btn_guardar = ctk.CTkButton(
            frame_botones,
            text="REGISTRAR",
            width=150,
            height=40,
            font=("Segoe UI", 12, "bold"),
            fg_color=AppTheme.PRIMARY,
            hover_color=AppTheme.PRIMARY_DARK,
            command=self.registrar
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

    def registrar(self):
        """Registrar el movimiento"""
        tipo = self.tipo_movimiento.get()
        cantidad = self.txt_cantidad.get().strip()
        motivo = self.txt_motivo.get().strip()

        if tipo == "entrada":
            quien = self.txt_recibe.get().strip()
            quien_recibe = quien
            quien_retira = ""
        else:
            quien = self.txt_retira.get().strip()
            quien_recibe = ""
            quien_retira = quien

        # Validaciones
        if not cantidad:
            messagebox.showerror("Error", "Ingrese la cantidad")
            return

        try:
            cantidad_int = int(cantidad)
            if cantidad_int <= 0:
                messagebox.showerror("Error", "La cantidad debe ser mayor a 0")
                return
        except ValueError:
            messagebox.showerror("Error", "La cantidad debe ser un número")
            return

        if not quien:
            msg = "¿Quién recibe?" if tipo == "entrada" else "¿Quién retira?"
            messagebox.showerror("Error", msg)
            return

        if not motivo:
            messagebox.showerror("Error", "Ingrese el motivo")
            return

        # Registrar movimiento
        success = Movimiento.registrar(
            self.medicamento_id, tipo, cantidad_int,
            quien_recibe, quien_retira, motivo
        )

        if success:
            # Actualizar stock
            medicamento = Medicamento.obtener_por_id(self.medicamento_id)
            nuevo_stock = medicamento['stock']

            if tipo == "entrada":
                nuevo_stock += cantidad_int
            else:
                if nuevo_stock >= cantidad_int:
                    nuevo_stock -= cantidad_int
                else:
                    messagebox.showerror("Error", "Stock insuficiente")
                    return

            Medicamento.actualizar(
                self.medicamento_id,
                medicamento['nombre'],
                medicamento['presentacion'],
                medicamento['ubicacion'],
                medicamento['caducidad'],
                medicamento['observaciones'],
                nuevo_stock,
                medicamento['es_donacion'],
                medicamento['donante']
            )

            messagebox.showinfo("Éxito", "Movimiento registrado correctamente")
            self.destroy()
        else:
            messagebox.showerror("Error", "No se pudo registrar el movimiento")