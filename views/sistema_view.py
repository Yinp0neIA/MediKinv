import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
from datetime import datetime, date
import calendar
from assets.styles.themes import AppTheme
from models.medicamento import Medicamento
from models.movimiento import Movimiento
from config.database import Database
from utils.icon_manager import IconManager

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


class SistemaView(ctk.CTk):

    def __init__(self, usuario, area_id, area_nombre):
        super().__init__()
        IconManager.aplicar_icono(self)

        self.usuario = usuario
        self.area_id = area_id
        self.area_nombre = area_nombre
        self.id_medicamento = None
        self.id_movimiento_editando = None
        self.fecha_desde = date(date.today().year, 1, 1)
        self.fecha_hasta = date.today()
        self._debug = True
        self.fullscreen_flag = False

        self.medicamento_model = Medicamento
        self.movimiento_model = Movimiento

        self.title(f"MediKinv - {area_nombre}")
        self.geometry("1600x900")
        AppTheme.aplicar_tema()
        self.configurar_estilos()
        self.crear_interfaz()
        self.after(200, self.maximizar_y_cargar)

        self.bind('<Escape>', self.toggle_fullscreen)

    def debug_print(self, msg):
        if self._debug:
            print(f"[DEBUG {self.area_nombre}] {msg}")

    def maximizar_y_cargar(self):
        try:
            self.state('zoomed')
        except:
            pass

        self.id_medicamento = None
        self.id_movimiento_editando = None

        if hasattr(self, 'tabla_inventario'):
            for item in self.tabla_inventario.selection():
                self.tabla_inventario.selection_remove(item)

        if hasattr(self, 'tabla_mov'):
            for item in self.tabla_mov.selection():
                self.tabla_mov.selection_remove(item)

        self.cargar_inventario()
        self.cargar_combo_medicamentos_filtro()
        self.cargar_combo_ubicaciones()
        self.cargar_combo_medicamentos_mov()
        self.cargar_historial()

        self.limpiar_formulario_medicamento()
        self.limpiar_formulario_movimientos()

        self.modo_crud = "medicamento"
        self.mostrar_formulario_medicamento()

    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Inventory.Treeview", font=("Segoe UI", 10), rowheight=34,
                        background="white", fieldbackground="white", foreground=AppTheme.TEXT_PRIMARY)
        style.configure("Inventory.Treeview.Heading", font=("Segoe UI", 11, "bold"),
                        background=AppTheme.PRIMARY, foreground=AppTheme.TEXT_ON_PRIMARY, relief="flat")
        style.configure("Movements.Treeview", font=("Segoe UI", 10), rowheight=34,
                        background="white", fieldbackground="white", foreground=AppTheme.TEXT_PRIMARY)
        style.configure("Movements.Treeview.Heading", font=("Segoe UI", 11, "bold"),
                        background=AppTheme.PRIMARY, foreground=AppTheme.TEXT_ON_PRIMARY, relief="flat")

    def crear_interfaz(self):
        self.frame_principal = ctk.CTkFrame(self, fg_color=AppTheme.BACKGROUND, corner_radius=0)
        self.frame_principal.pack(fill="both", expand=True)
        self.crear_header()

        self.tab_view = ctk.CTkTabview(self.frame_principal, fg_color=AppTheme.SURFACE, corner_radius=15)
        self.tab_view.pack(fill="both", expand=True, padx=20, pady=20)

        self.tab_inventario = self.tab_view.add("INVENTARIO")
        self.tab_movimientos = self.tab_view.add("MOVIMIENTOS")

        self.crear_pestana_inventario()
        self.crear_pestana_movimientos()

    def crear_header(self):
        header = ctk.CTkFrame(self.frame_principal, fg_color=AppTheme.PRIMARY, height=80, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)
        icono = "F" if self.area_id == 1 else "O"
        titulo = ctk.CTkLabel(header, text=f"{icono} MediKinv - {self.area_nombre}",
                              font=("Segoe UI", 24, "bold"), text_color=AppTheme.TEXT_ON_PRIMARY)
        titulo.pack(side="left", padx=30, pady=20)
        frame_usuario = ctk.CTkFrame(header, fg_color="transparent")
        frame_usuario.pack(side="right", padx=30, pady=20)
        label_usuario = ctk.CTkLabel(frame_usuario, text=f"Usuario: {self.usuario.nombre_usuario}",
                                     font=("Segoe UI", 12), text_color=AppTheme.TEXT_ON_PRIMARY)
        label_usuario.pack(side="left", padx=10)

        btn_fullscreen = ctk.CTkButton(frame_usuario, text="Pantalla Completa", width=150, height=38,
                                       font=("Segoe UI", 11), fg_color="transparent",
                                       hover_color=AppTheme.PRIMARY_LIGHT, text_color=AppTheme.TEXT_ON_PRIMARY,
                                       corner_radius=12, command=self.toggle_fullscreen)
        btn_fullscreen.pack(side="left", padx=10)

        btn_volver = ctk.CTkButton(frame_usuario, text="Volver a Areas", width=130, height=38,
                                   font=("Segoe UI", 11), fg_color="transparent",
                                   hover_color=AppTheme.PRIMARY_LIGHT, text_color=AppTheme.TEXT_ON_PRIMARY,
                                   corner_radius=12, command=self.volver_areas)
        btn_volver.pack(side="left", padx=10)
        btn_cerrar = ctk.CTkButton(frame_usuario, text="Cerrar Sesion", width=120, height=38,
                                   font=("Segoe UI", 11), fg_color=AppTheme.PRIMARY_LIGHT,
                                   hover_color=AppTheme.PRIMARY_DARK, text_color=AppTheme.TEXT_ON_PRIMARY,
                                   corner_radius=12, command=self.cerrar_sesion)
        btn_cerrar.pack(side="left", padx=10)

    def toggle_fullscreen(self, event=None):
        self.fullscreen_flag = not self.fullscreen_flag
        self.attributes('-fullscreen', self.fullscreen_flag)
        if not self.fullscreen_flag:
            self.geometry("1600x900")
            self.update_idletasks()

    def volver_areas(self):
        try:
            for after_id in self.tk.eval('after info').split():
                try:
                    self.after_cancel(int(after_id))
                except:
                    pass
        except:
            pass
        from views.area_select_view import AreaSelectView
        area_view = AreaSelectView(self.usuario)
        self.destroy()
        area_view.mainloop()

    def cerrar_sesion(self):
        try:
            for after_id in self.tk.eval('after info').split():
                try:
                    self.after_cancel(int(after_id))
                except:
                    pass
        except:
            pass
        from views.login_view import LoginView
        login = LoginView()
        self.destroy()
        login.mainloop()

    def crear_pestana_inventario(self):
        panel_principal = ctk.CTkFrame(self.tab_inventario, fg_color="transparent")
        panel_principal.pack(fill="both", expand=True, padx=10, pady=10)

        panel_izquierdo = ctk.CTkFrame(panel_principal, fg_color=AppTheme.SURFACE, corner_radius=15)
        panel_izquierdo.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.panel_derecho_inventario = ctk.CTkFrame(panel_principal, fg_color=AppTheme.SURFACE, corner_radius=15,
                                                     width=500)
        self.panel_derecho_inventario.pack(side="right", fill="y", padx=(10, 0))
        self.panel_derecho_inventario.pack_propagate(False)

        self.crear_filtros_inventario(panel_izquierdo)

        self.crear_leyenda_colores(panel_izquierdo)

        self.crear_tabla_inventario(panel_izquierdo)

        self.crear_botones_crud_inventario()
        self.frame_medicamento_form = ctk.CTkFrame(self.panel_derecho_inventario, fg_color="transparent")
        self.frame_movimiento_form = ctk.CTkFrame(self.panel_derecho_inventario, fg_color="transparent")

        self.crear_formulario_medicamento(self.frame_medicamento_form)
        self.crear_formulario_movimiento_rapido(self.frame_movimiento_form)

        self.modo_crud = "medicamento"
        self.mostrar_formulario_medicamento()

    def crear_leyenda_colores(self, parent):
        frame_leyenda = ctk.CTkFrame(parent, fg_color="transparent")
        frame_leyenda.pack(fill="x", pady=(5, 5), padx=15)

        ctk.CTkLabel(frame_leyenda, text="🎯 1▶2▶3▶4▶5▶6",
                     font=("Segoe UI", 9, "bold"),
                     text_color=AppTheme.TEXT_PRIMARY).pack(side="left", padx=(0, 10))

        colores = [
            ("#D4A837", "Sin stock"),
            ("#C75C5C", "Caducado"),
            ("#D97A3A", "Caduca mes"),
            ("#5A9E6F", "Nuevo"),
            ("#8B6B9E", "Controlado"),
            ("#5B7BA5", "Normal"),
        ]

        for color, texto in colores:
            frame_item = ctk.CTkFrame(frame_leyenda, fg_color="transparent")
            frame_item.pack(side="left", padx=4)

            cuadro = ctk.CTkFrame(frame_item, width=14, height=14, fg_color=color, corner_radius=3)
            cuadro.pack(side="left")
            cuadro.pack_propagate(False)

            ctk.CTkLabel(frame_item, text=texto, font=("Segoe UI", 9),
                         text_color=AppTheme.TEXT_SECONDARY).pack(side="left", padx=(3, 0))

    def crear_botones_crud_inventario(self):
        frame_botones = ctk.CTkFrame(self.panel_derecho_inventario, fg_color="transparent", height=45)
        frame_botones.pack(fill="x", padx=15, pady=(10, 0))
        frame_botones.pack_propagate(False)

        self.btn_crud_med = ctk.CTkButton(frame_botones, text="INVENTARIO CRUD", width=120, height=35,
                                          font=("Segoe UI", 12, "bold"), fg_color=AppTheme.PRIMARY,
                                          hover_color=AppTheme.PRIMARY_DARK, corner_radius=8,
                                          command=self.mostrar_formulario_medicamento)
        self.btn_crud_med.pack(side="left", padx=5)

        self.btn_crud_mov = ctk.CTkButton(frame_botones, text="MOVIMIENTOS CRUD", width=140, height=35,
                                          font=("Segoe UI", 12, "bold"), fg_color="#4a5568",
                                          hover_color="#2d3748", corner_radius=8,
                                          command=self.mostrar_formulario_movimiento_rapido)
        self.btn_crud_mov.pack(side="left", padx=5)

    def mostrar_formulario_medicamento(self):
        self.modo_crud = "medicamento"
        self.frame_movimiento_form.pack_forget()
        self.frame_medicamento_form.pack(fill="both", expand=True, padx=15, pady=15)
        self.btn_crud_med.configure(fg_color=AppTheme.PRIMARY)
        self.btn_crud_mov.configure(fg_color="#4a5568")
        if self.id_medicamento:
            self.cargar_medicamento_formulario(self.id_medicamento)

    def mostrar_formulario_movimiento_rapido(self):
        self.modo_crud = "movimiento"
        self.frame_medicamento_form.pack_forget()
        self.frame_movimiento_form.pack(fill="both", expand=True, padx=15, pady=15)
        self.btn_crud_mov.configure(fg_color=AppTheme.PRIMARY)
        self.btn_crud_med.configure(fg_color="#4a5568")
        if self.id_medicamento:
            self.preseleccionar_medicamento_movimiento_rapido(self.id_medicamento)

    def crear_filtros_inventario(self, parent):
        frame_filtros = ctk.CTkFrame(parent, fg_color="transparent")
        frame_filtros.pack(fill="x", pady=10, padx=15)

        frame_buscar = ctk.CTkFrame(frame_filtros, fg_color="transparent")
        frame_buscar.pack(fill="x", pady=5)
        self.txt_buscar = ctk.CTkEntry(frame_buscar, placeholder_text="Buscar por nombre...",
                                       height=42, font=("Segoe UI", 12), corner_radius=12)
        self.txt_buscar.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.txt_buscar.bind("<KeyRelease>", self.buscar_medicamento)
        btn_actualizar = ctk.CTkButton(frame_buscar, text="Actualizar", width=110, height=42,
                                       font=("Segoe UI", 11), fg_color=AppTheme.PRIMARY,
                                       hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                       command=self.cargar_inventario)
        btn_actualizar.pack(side="right", padx=(0, 5))
        btn_exportar_inv = ctk.CTkButton(frame_buscar, text="Exportar Inventario", width=150, height=42,
                                         font=("Segoe UI", 11), fg_color="#2c7a4d",
                                         hover_color="#1e5a38", corner_radius=12,
                                         command=self.exportar_excel_inventario)
        btn_exportar_inv.pack(side="right", padx=5)

        frame_filtros_opciones = ctk.CTkFrame(frame_filtros, fg_color="transparent")
        frame_filtros_opciones.pack(fill="x", pady=5)
        ctk.CTkLabel(frame_filtros_opciones, text="Filtrar por:", font=("Segoe UI", 11, "bold"),
                     text_color=AppTheme.TEXT_PRIMARY).pack(side="left", padx=(0, 10))

        ctk.CTkLabel(frame_filtros_opciones, text="Presentacion:", font=("Segoe UI", 11)).pack(side="left")
        self.txt_presentacion_filtro = ctk.CTkEntry(frame_filtros_opciones, width=180, height=35,
                                                    corner_radius=12, placeholder_text="Ej: Tableta")
        self.txt_presentacion_filtro.pack(side="left", padx=5)
        self.txt_presentacion_filtro.bind("<KeyRelease>", self.buscar_medicamento)

        ctk.CTkLabel(frame_filtros_opciones, text="Ubicacion:", font=("Segoe UI", 11)).pack(side="left", padx=(15, 5))
        self.combo_ubicacion_filtro = ctk.CTkComboBox(frame_filtros_opciones, values=["Todas"], width=180,
                                                      height=35, corner_radius=12, state="readonly",
                                                      command=self.buscar_medicamento)
        self.combo_ubicacion_filtro.pack(side="left", padx=5)

        frame_filtros_avanzados = ctk.CTkFrame(frame_filtros, fg_color="transparent")
        frame_filtros_avanzados.pack(fill="x", pady=5)

        ctk.CTkLabel(frame_filtros_avanzados, text="Controlado:", font=("Segoe UI", 11, "bold")).pack(side="left",
                                                                                                      padx=(0, 5))
        self.combo_controlado = ctk.CTkComboBox(frame_filtros_avanzados, values=["Todos", "Si", "No"],
                                                width=100, height=35, corner_radius=12, state="readonly",
                                                command=self.buscar_medicamento)
        self.combo_controlado.pack(side="left", padx=5)
        self.combo_controlado.set("Todos")

        ctk.CTkLabel(frame_filtros_avanzados, text="Caducidad:", font=("Segoe UI", 11, "bold")).pack(side="left",
                                                                                                     padx=(15, 5))
        self.combo_caducidad = ctk.CTkComboBox(frame_filtros_avanzados,
                                               values=["Todos", "Vigentes", "Proximos a caducar (30 dias)",
                                                       "Caducados"],
                                               width=200, height=35, corner_radius=12, state="readonly",
                                               command=self.buscar_medicamento)
        self.combo_caducidad.pack(side="left", padx=5)
        self.combo_caducidad.set("Todos")

        ctk.CTkLabel(frame_filtros_avanzados, text="Stock:", font=("Segoe UI", 11, "bold")).pack(side="left",
                                                                                                 padx=(15, 5))
        self.combo_stock = ctk.CTkComboBox(frame_filtros_avanzados,
                                           values=["Todos", "Sin stock", "Bajo (1-29)", "Normal (30-100)",
                                                   "Alto (>100)"],
                                           width=140, height=35, corner_radius=12, state="readonly",
                                           command=self.buscar_medicamento)
        self.combo_stock.pack(side="left", padx=5)
        self.combo_stock.set("Todos")

        btn_limpiar_filtros = ctk.CTkButton(frame_filtros_avanzados, text="Limpiar filtros", width=120, height=35,
                                            font=("Segoe UI", 11), fg_color="#4a5568", hover_color="#2d3748",
                                            corner_radius=12, command=self.limpiar_filtros_inventario)
        btn_limpiar_filtros.pack(side="left", padx=20)

    def limpiar_filtros_inventario(self):
        self.txt_buscar.delete(0, "end")
        self.txt_presentacion_filtro.delete(0, "end")
        self.combo_ubicacion_filtro.set("Todas")
        self.combo_controlado.set("Todos")
        self.combo_caducidad.set("Todos")
        self.combo_stock.set("Todos")
        self.cargar_inventario()

    def cargar_combo_ubicaciones(self):
        try:
            tabla_inventario = "inventario_farmacia" if self.area_id == 1 else "inventario_odontologia"
            query = f"SELECT DISTINCT ubicacion FROM {tabla_inventario} WHERE ubicacion IS NOT NULL AND ubicacion != ''"
            result = Database.execute_query(query)
            ubicaciones = sorted([row[0] for row in result]) if result else []
            valores = ["Todas"] + ubicaciones
            self.combo_ubicacion_filtro.configure(values=valores)
            if self.combo_ubicacion_filtro.get() not in valores:
                self.combo_ubicacion_filtro.set("Todas")
        except Exception as e:
            self.debug_print(f"Error cargando ubicaciones: {e}")

    def cargar_inventario(self, busqueda=""):
        try:
            for item in self.tabla_inventario.get_children():
                self.tabla_inventario.delete(item)

            if busqueda == "":
                busqueda = self.txt_buscar.get().strip()
            filtro_presentacion = self.txt_presentacion_filtro.get().strip().lower()
            filtro_ubicacion = self.combo_ubicacion_filtro.get()
            filtro_controlado = self.combo_controlado.get()
            filtro_caducidad = self.combo_caducidad.get()
            filtro_stock = self.combo_stock.get()

            tabla_inventario = "inventario_farmacia" if self.area_id == 1 else "inventario_odontologia"
            tabla_movimientos = "movimientos_farmacia" if self.area_id == 1 else "movimientos_odontologia"

            query = f"""
                SELECT 
                    i.id_inventario,
                    c.nombre,
                    CONCAT(ISNULL(c.forma_farmaceutica, ''), ' ', ISNULL(c.concentracion, '')) as presentacion,
                    i.ubicacion,
                    i.stock,
                    i.caducidad,
                    i.es_donacion,
                    i.donante,
                    i.observaciones,
                    c.es_controlado,
                    (SELECT MIN(fecha) FROM {tabla_movimientos} WHERE id_inventario = i.id_inventario) as fecha_primer_movimiento
                FROM {tabla_inventario} i
                JOIN catalogo_medicamentos c ON i.id_catalogo = c.id_catalogo
                WHERE 1=1
                ORDER BY i.id_inventario
            """
            params = []

            if busqueda:
                query = query.replace("WHERE 1=1", "WHERE 1=1 AND (c.nombre LIKE ? OR i.observaciones LIKE ?)")
                params.append(f"%{busqueda}%")
                params.append(f"%{busqueda}%")

            result = Database.execute_query(query, tuple(params) if params else None)
            if not result:
                self.tabla_inventario.insert("", "end", values=("", "", "", "", "", "", "", "", "No hay medicamentos"))
                return

            fecha_hoy = date.today()
            mes_actual = fecha_hoy.month
            anio_actual = fecha_hoy.year
            registros_mostrados = 0

            for row in result:
                id_med = row[0]
                nombre = row[1]
                presentacion = row[2].strip() or "---"
                ubicacion = row[3] or ""
                stock = row[4]
                caducidad = row[5]
                es_donacion = row[6]
                donante = row[7] or ""
                observaciones = row[8] or ""
                es_controlado = row[9]
                fecha_primer_movimiento = row[10]

                if filtro_presentacion and filtro_presentacion not in presentacion.lower():
                    continue

                if filtro_ubicacion != "Todas" and ubicacion != filtro_ubicacion:
                    continue

                if filtro_controlado == "Si" and not es_controlado:
                    continue
                if filtro_controlado == "No" and es_controlado:
                    continue

                estado_cad = "Vigentes"
                if caducidad:
                    if caducidad < fecha_hoy:
                        estado_cad = "Caducados"
                    elif (caducidad - fecha_hoy).days <= 30:
                        estado_cad = "Proximos a caducar (30 dias)"
                    else:
                        estado_cad = "Vigentes"

                if filtro_caducidad != "Todos":
                    if filtro_caducidad == "Vigentes" and estado_cad != "Vigentes":
                        continue
                    elif filtro_caducidad == "Proximos a caducar (30 dias)" and estado_cad != "Proximos a caducar (30 dias)":
                        continue
                    elif filtro_caducidad == "Caducados" and estado_cad != "Caducados":
                        continue

                if stock == 0:
                    estado_stock = "Sin stock"
                elif 1 <= stock <= 29:
                    estado_stock = "Bajo (1-29)"
                elif 30 <= stock <= 100:
                    estado_stock = "Normal (30-100)"
                else:
                    estado_stock = "Alto (>100)"

                if filtro_stock != "Todos" and estado_stock != filtro_stock:
                    continue

                donacion_str = "Si" if es_donacion else "No"
                caducidad_str = caducidad.strftime("%d/%m/%Y") if caducidad else "---"
                if len(observaciones) > 50:
                    observaciones = observaciones[:47] + "..."

                tags = []

                if stock == 0:
                    tags.append('sin_stock')
                elif caducidad and caducidad < fecha_hoy:
                    tags.append('caducado')
                elif caducidad and caducidad.month == mes_actual and caducidad.year == anio_actual:
                    tags.append('caduca_mes')
                elif fecha_primer_movimiento and (fecha_hoy - fecha_primer_movimiento).days <= 7:
                    tags.append('nuevo')
                elif es_controlado:
                    tags.append('controlado')
                else:
                    tags.append('normal')

                self.tabla_inventario.insert("", "end", values=(
                    id_med, nombre, presentacion, ubicacion, stock,
                    caducidad_str, donacion_str, donante, observaciones
                ), tags=tags)
                registros_mostrados += 1

            if registros_mostrados == 0:
                self.tabla_inventario.insert("", "end", values=("", "", "", "", "", "", "", "",
                                                                "No hay medicamentos con esos filtros"))

            # Configurar colores
            self.tabla_inventario.tag_configure('sin_stock', background='#D4A837', foreground='#FFFFFF')
            self.tabla_inventario.tag_configure('caducado', background='#C75C5C', foreground='#FFFFFF')
            self.tabla_inventario.tag_configure('caduca_mes', background='#D97A3A', foreground='#FFFFFF')
            self.tabla_inventario.tag_configure('nuevo', background='#5A9E6F', foreground='#FFFFFF')
            self.tabla_inventario.tag_configure('controlado', background='#8B6B9E', foreground='#FFFFFF')
            self.tabla_inventario.tag_configure('normal', background='#5B7BA5', foreground='#FFFFFF')

        except Exception as e:
            self.debug_print(f"Error al cargar inventario: {e}")
            messagebox.showerror("Error", f"No se pudo cargar el inventario:\n{e}")

    def buscar_medicamento(self, event):
        self.cargar_inventario(self.txt_buscar.get())

    def crear_tabla_inventario(self, parent):
        frame_tabla = ctk.CTkFrame(parent, fg_color="transparent")
        frame_tabla.pack(fill="both", expand=True, pady=10, padx=15)
        scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical")
        scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal")
        self.tabla_inventario = ttk.Treeview(
            frame_tabla,
            columns=("ID", "Nombre", "Presentacion", "Ubicacion", "Stock", "Caducidad", "Donacion", "Donante",
                     "Observaciones"),
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
            style="Inventory.Treeview",
            height=18
        )
        scroll_y.config(command=self.tabla_inventario.yview)
        scroll_x.config(command=self.tabla_inventario.xview)
        columnas = [("ID", 50, "center"), ("Nombre", 280, "w"), ("Presentacion", 200, "w"),
                    ("Ubicacion", 180, "w"), ("Stock", 80, "center"), ("Caducidad", 100, "center"),
                    ("Donacion", 80, "center"), ("Donante", 180, "w"), ("Observaciones", 250, "w")]
        for col, width, anchor in columnas:
            self.tabla_inventario.heading(col, text=col)
            self.tabla_inventario.column(col, width=width, anchor=anchor)
        self.tabla_inventario.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)
        self.tabla_inventario.bind("<<TreeviewSelect>>", self.on_medicamento_seleccionado)

    def on_medicamento_seleccionado(self, event):
        seleccion = self.tabla_inventario.selection()
        if seleccion:
            valores = self.tabla_inventario.item(seleccion[0])['values']
            if valores and valores[0]:
                self.id_medicamento = valores[0]
                if self.modo_crud == "medicamento":
                    self.cargar_medicamento_formulario(self.id_medicamento)
                else:
                    self.preseleccionar_medicamento_movimiento_rapido(self.id_medicamento)

    # ---------- Formulario de medicamento (CRUD inventario) ----------
    def crear_formulario_medicamento(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent", height=750)
        scroll.pack(fill="both", expand=True, padx=15, pady=15)

        self.form_titulo_med = ctk.CTkLabel(scroll, text="AGREGAR MEDICAMENTO",
                                            font=("Segoe UI", 18, "bold"), text_color=AppTheme.PRIMARY)
        self.form_titulo_med.pack(pady=(0, 15))

        ctk.CTkLabel(scroll, text="Medicamento", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.combo_catalogo = ctk.CTkComboBox(scroll, values=[], width=400, height=38,
                                              font=("Segoe UI", 11), corner_radius=12,
                                              state="readonly", command=self.on_catalogo_seleccionado)
        self.combo_catalogo.pack(fill="x", pady=(0, 5))
        self.search_entry = ctk.CTkEntry(scroll, placeholder_text="Escriba para buscar medicamento...",
                                         height=38, font=("Segoe UI", 11), corner_radius=12)
        self.search_entry.pack(fill="x", pady=(0, 5))
        self.search_entry.bind("<KeyRelease>", self.actualizar_combo_catalogo)
        self.catalogo_items = []

        ctk.CTkLabel(scroll, text="Presentacion", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_presentacion = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11),
                                             corner_radius=12, state="readonly")
        self.txt_presentacion.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Ubicacion", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_ubicacion = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_ubicacion.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Stock", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_stock = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_stock.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Fecha de caducidad", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        frame_fecha = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_fecha.pack(fill="x", pady=(0, 5))
        self.txt_caducidad = ctk.CTkEntry(frame_fecha, placeholder_text="DD/MM/AAAA", height=38,
                                          font=("Segoe UI", 11), corner_radius=12)
        self.txt_caducidad.pack(side="left", fill="x", expand=True, padx=(0, 5))
        btn_cal = ctk.CTkButton(frame_fecha, text="Calendario", width=65, height=38, font=("Segoe UI", 11),
                                fg_color=AppTheme.PRIMARY, hover_color=AppTheme.PRIMARY_DARK,
                                corner_radius=12, command=self.abrir_calendario_medicamento)
        btn_cal.pack(side="right")

        ctk.CTkLabel(scroll, text="Quien recibe (Registro inicial)", font=("Segoe UI", 11, "bold")).pack(anchor="w",
                                                                                                         pady=(10, 2))
        self.txt_quien_recibe_inicial = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_quien_recibe_inicial.pack(fill="x", pady=(0, 5))
        self.txt_quien_recibe_inicial.insert(0, "Registro inicial")

        self.es_donacion = ctk.BooleanVar(value=False)
        frame_donacion = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_donacion.pack(fill="x", pady=(10, 2))
        check_donacion = ctk.CTkCheckBox(frame_donacion, text="Este medicamento es una donacion",
                                         variable=self.es_donacion, font=("Segoe UI", 11),
                                         command=self.toggle_donante)
        check_donacion.pack(anchor="w")
        self.frame_donante = ctk.CTkFrame(scroll, fg_color="transparent")
        ctk.CTkLabel(self.frame_donante, text="Donante", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.txt_donante = ctk.CTkEntry(self.frame_donante, height=38, font=("Segoe UI", 11),
                                        corner_radius=12, state="disabled")
        self.txt_donante.pack(fill="x", pady=(2, 0))
        self.frame_donante.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Observaciones", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        frame_obs = ctk.CTkFrame(scroll, fg_color="transparent", border_width=1,
                                 border_color=AppTheme.PRIMARY_LIGHT, corner_radius=12)
        frame_obs.pack(fill="x", pady=(0, 5))
        self.txt_observaciones = ctk.CTkTextbox(frame_obs, height=80, font=("Segoe UI", 11),
                                                corner_radius=12, fg_color="white")
        self.txt_observaciones.pack(fill="both", expand=True, padx=2, pady=2)

        frame_botones = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_botones.pack(fill="x", pady=15)
        self.btn_guardar_med = ctk.CTkButton(frame_botones, text="GUARDAR", height=45,
                                             font=("Segoe UI", 13, "bold"), fg_color=AppTheme.PRIMARY,
                                             hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                             command=self.guardar_medicamento)
        self.btn_guardar_med.pack(fill="x", pady=2)
        self.btn_eliminar_med = ctk.CTkButton(frame_botones, text="ELIMINAR", height=45,
                                              font=("Segoe UI", 13), fg_color="#e53e3e",
                                              hover_color="#c53030", text_color=AppTheme.TEXT_ON_PRIMARY,
                                              corner_radius=12, command=self.eliminar_medicamento, state="disabled")
        self.btn_eliminar_med.pack(fill="x", pady=2)
        self.btn_limpiar_med = ctk.CTkButton(frame_botones, text="LIMPIAR FORMULARIO", height=40,
                                             font=("Segoe UI", 12), fg_color="#4a5568",
                                             hover_color="#2d3748", text_color=AppTheme.TEXT_ON_PRIMARY,
                                             corner_radius=12, command=self.limpiar_formulario_medicamento)
        self.btn_limpiar_med.pack(fill="x", pady=2)

        self.cargar_catalogo_inicial()

    def cargar_catalogo_inicial(self):
        catalogo = self.medicamento_model.obtener_catalogo()
        self.catalogo_items = []
        valores = []
        for item in catalogo:
            self.catalogo_items.append({
                'id_catalogo': item['id_catalogo'],
                'nombre': item['nombre'],
                'presentacion': item['presentacion']
            })
            valores.append(f"{item['nombre']} - {item['presentacion']}")
        self.combo_catalogo.configure(values=valores)
        if valores:
            self.combo_catalogo.set(valores[0])
            self.on_catalogo_seleccionado(None)

    def actualizar_combo_catalogo(self, event=None):
        busqueda = self.search_entry.get().strip()
        if not busqueda:
            self.cargar_catalogo_inicial()
            return
        catalogo = self.medicamento_model.obtener_catalogo(busqueda)
        self.catalogo_items = []
        valores = []
        for item in catalogo:
            self.catalogo_items.append({
                'id_catalogo': item['id_catalogo'],
                'nombre': item['nombre'],
                'presentacion': item['presentacion']
            })
            valores.append(f"{item['nombre']} - {item['presentacion']}")
        self.combo_catalogo.configure(values=valores)
        if valores:
            self.combo_catalogo.set(valores[0])
            self.on_catalogo_seleccionado(None)
        else:
            self.combo_catalogo.set("")
            self.txt_presentacion.configure(state="normal")
            self.txt_presentacion.delete(0, "end")
            self.txt_presentacion.insert(0, "No hay coincidencias")
            self.txt_presentacion.configure(state="readonly")

    def on_catalogo_seleccionado(self, choice):
        seleccion = self.combo_catalogo.get()
        if not seleccion:
            return
        for item in self.catalogo_items:
            if f"{item['nombre']} - {item['presentacion']}" == seleccion:
                self.txt_presentacion.configure(state="normal")
                self.txt_presentacion.delete(0, "end")
                self.txt_presentacion.insert(0, item['presentacion'])
                self.txt_presentacion.configure(state="readonly")
                self.id_catalogo_seleccionado = item['id_catalogo']
                break

    def cargar_medicamento_formulario(self, medicamento_id):
        try:
            medicamento = self.medicamento_model.obtener_por_id(medicamento_id)
            if medicamento:
                self.form_titulo_med.configure(text="EDITAR MEDICAMENTO")
                self.id_medicamento = medicamento['id_medicamento']
                catalogo = self.medicamento_model.obtener_catalogo()
                for item in catalogo:
                    if item['nombre'] == medicamento['nombre']:
                        self.id_catalogo_seleccionado = item['id_catalogo']
                        self.search_entry.delete(0, "end")
                        self.search_entry.insert(0, medicamento['nombre'])
                        self.actualizar_combo_catalogo()
                        valor_combo = f"{medicamento['nombre']} - {medicamento['presentacion']}"
                        if valor_combo in self.combo_catalogo.cget('values'):
                            self.combo_catalogo.set(valor_combo)
                        break
                self.txt_ubicacion.delete(0, "end")
                self.txt_ubicacion.insert(0, medicamento['ubicacion'])
                self.txt_stock.delete(0, "end")
                self.txt_stock.insert(0, str(medicamento['stock']))
                if medicamento['caducidad']:
                    self.txt_caducidad.delete(0, "end")
                    self.txt_caducidad.insert(0, medicamento['caducidad'].strftime("%d/%m/%Y"))
                self.es_donacion.set(medicamento['es_donacion'])
                if medicamento['es_donacion'] and medicamento['donante']:
                    self.txt_donante.configure(state="normal")
                    self.txt_donante.delete(0, "end")
                    self.txt_donante.insert(0, medicamento['donante'])
                else:
                    self.txt_donante.configure(state="disabled")
                    self.txt_donante.delete(0, "end")
                if medicamento.get('observaciones'):
                    self.txt_observaciones.delete("1.0", "end")
                    self.txt_observaciones.insert("1.0", medicamento['observaciones'])
                self.btn_eliminar_med.configure(state="normal")
        except Exception as e:
            self.debug_print(f"Error al cargar medicamento: {e}")

    def guardar_medicamento(self):
        try:
            if not hasattr(self, 'id_catalogo_seleccionado') or not self.id_catalogo_seleccionado:
                messagebox.showerror("Error", "Debe seleccionar un medicamento del catalogo")
                return

            catalogo = self.medicamento_model.obtener_catalogo()
            item_seleccionado = None
            for item in catalogo:
                if item['id_catalogo'] == self.id_catalogo_seleccionado:
                    item_seleccionado = item
                    break
            if not item_seleccionado:
                messagebox.showerror("Error", "Medicamento no encontrado en catalogo")
                return

            nombre = item_seleccionado['nombre']
            presentacion = item_seleccionado['presentacion']
            ubicacion = self.txt_ubicacion.get().strip()
            stock = self.txt_stock.get().strip()
            caducidad_str = self.txt_caducidad.get().strip()
            observaciones = self.txt_observaciones.get("1.0", "end-1c").strip()
            es_donacion = self.es_donacion.get()
            donante = self.txt_donante.get().strip() if es_donacion else ""
            quien_recibe_inicial = self.txt_quien_recibe_inicial.get().strip()

            if not nombre or not presentacion:
                messagebox.showerror("Error", "Datos de medicamento invalidos")
                return

            try:
                stock_int = int(stock) if stock else 0
                if not self.id_medicamento and stock_int < 1:
                    messagebox.showerror("Error", "El stock inicial debe ser al menos 1 unidad")
                    return
                if stock_int < 0:
                    messagebox.showerror("Error", "El stock no puede ser negativo")
                    return
            except ValueError:
                messagebox.showerror("Error", "El stock debe ser un numero")
                return

            if not self.id_medicamento and not quien_recibe_inicial:
                messagebox.showerror("Error", "Debe especificar quien recibe el medicamento")
                return

            caducidad = None
            if caducidad_str:
                try:
                    caducidad = datetime.strptime(caducidad_str, "%d/%m/%Y").date()
                except ValueError:
                    messagebox.showerror("Error", "Formato de fecha invalido. Use DD/MM/AAAA")
                    return

            if self.id_medicamento:
                medicamento = self.medicamento_model.obtener_por_id(self.id_medicamento)
                if not medicamento:
                    messagebox.showerror("Error", "Medicamento no encontrado")
                    return
                success = self.medicamento_model.actualizar(
                    self.id_medicamento, nombre, presentacion, ubicacion,
                    caducidad, observaciones, stock_int, es_donacion, donante
                )
                msg = "actualizado"
            else:
                success = self.medicamento_model.crear_desde_catalogo(
                    self.area_id, self.id_catalogo_seleccionado, ubicacion,
                    caducidad, stock_int, es_donacion, donante, observaciones
                )
                msg = "agregado"
                if success and stock_int > 0:
                    medicamentos = self.medicamento_model.obtener_por_area(self.area_id)
                    if medicamentos:
                        nuevo_id = medicamentos[-1]['id_medicamento']
                        self.movimiento_model.registrar(
                            nuevo_id, "entrada", stock_int,
                            quien_recibe_inicial, "",
                            f"Registro inicial de medicamento - {quien_recibe_inicial}"
                        )

            if success:
                messagebox.showinfo("Exito", f"Medicamento {msg} correctamente")
                self.limpiar_formulario_medicamento()
                self.cargar_inventario()
                self.cargar_combo_medicamentos_filtro()
                self.cargar_combo_ubicaciones()
                self.cargar_historial()
                self.cargar_combo_medicamentos_mov()
            else:
                messagebox.showerror("Error", f"No se pudo {msg} el medicamento")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar: {str(e)}")

    def eliminar_medicamento(self):
        if self.id_medicamento:
            self.medicamento_model.eliminar(self.id_medicamento)
            messagebox.showinfo("Exito", "Medicamento eliminado correctamente")
            self.limpiar_formulario_medicamento()
            self.cargar_inventario()
            self.cargar_combo_medicamentos_filtro()
            self.cargar_combo_ubicaciones()
            self.cargar_historial()
            self.cargar_combo_medicamentos_mov()

    def limpiar_formulario_medicamento(self):
        self.id_medicamento = None
        self.form_titulo_med.configure(text="AGREGAR MEDICAMENTO")
        self.search_entry.delete(0, "end")
        self.cargar_catalogo_inicial()
        self.txt_ubicacion.delete(0, "end")
        self.txt_stock.delete(0, "end")
        self.txt_caducidad.delete(0, "end")
        self.txt_observaciones.delete("1.0", "end")
        self.txt_quien_recibe_inicial.delete(0, "end")
        self.txt_quien_recibe_inicial.insert(0, "Registro inicial")
        self.es_donacion.set(False)
        self.txt_donante.configure(state="disabled")
        self.txt_donante.delete(0, "end")
        self.btn_eliminar_med.configure(state="disabled")
        for item in self.tabla_inventario.selection():
            self.tabla_inventario.selection_remove(item)

    def toggle_donante(self):
        if self.es_donacion.get():
            self.txt_donante.configure(state="normal")
            self.txt_donante.configure(placeholder_text="Nombre del donante...")
        else:
            self.txt_donante.configure(state="disabled")
            self.txt_donante.delete(0, "end")
            self.txt_donante.configure(placeholder_text="")

    # ---------- Formulario de movimiento rapido (dentro de inventario) ----------
    def crear_formulario_movimiento_rapido(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent", height=750)
        scroll.pack(fill="both", expand=True, padx=15, pady=15)

        self.form_titulo_mov_rapido = ctk.CTkLabel(scroll, text="REGISTRAR MOVIMIENTO",
                                                   font=("Segoe UI", 18, "bold"), text_color=AppTheme.PRIMARY)
        self.form_titulo_mov_rapido.pack(pady=(0, 15))

        ctk.CTkLabel(scroll, text="Medicamento", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.combo_medicamento_rapido = ctk.CTkComboBox(scroll, values=[], width=400, height=38,
                                                        font=("Segoe UI", 11), corner_radius=12, state="readonly",
                                                        command=self.actualizar_stock_rapido)
        self.combo_medicamento_rapido.pack(fill="x", pady=(0, 5))

        self.lbl_stock_rapido = ctk.CTkLabel(scroll, text="Stock disponible: -- unidades",
                                             font=("Segoe UI", 11), text_color=AppTheme.PRIMARY)
        self.lbl_stock_rapido.pack(anchor="w", pady=(0, 10))

        ctk.CTkLabel(scroll, text="Tipo de Movimiento", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.tipo_mov_rapido = ctk.StringVar(value="entrada")
        frame_tipo = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_tipo.pack(fill="x", pady=5)
        radio_entrada = ctk.CTkRadioButton(frame_tipo, text="ENTRADA", variable=self.tipo_mov_rapido,
                                           value="entrada", font=("Segoe UI", 11, "bold"),
                                           fg_color="#2c7a4d", hover_color="#1e5a38",
                                           text_color=AppTheme.TEXT_PRIMARY,
                                           command=self.actualizar_label_quien_rapido)
        radio_entrada.pack(side="left", padx=10)
        radio_salida = ctk.CTkRadioButton(frame_tipo, text="SALIDA", variable=self.tipo_mov_rapido,
                                          value="salida", font=("Segoe UI", 11, "bold"),
                                          fg_color="#e53e3e", hover_color="#c53030",
                                          text_color=AppTheme.TEXT_PRIMARY,
                                          command=self.actualizar_label_quien_rapido)
        radio_salida.pack(side="left", padx=10)

        ctk.CTkLabel(scroll, text="Cantidad", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_cantidad_rapido = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_cantidad_rapido.pack(fill="x", pady=(0, 5))

        self.label_quien_rapido = ctk.CTkLabel(scroll, text="Quien Recibe", font=("Segoe UI", 11, "bold"))
        self.label_quien_rapido.pack(anchor="w", pady=(10, 2))
        self.txt_quien_rapido = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_quien_rapido.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Motivo", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_motivo_rapido = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_motivo_rapido.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Fecha", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        frame_fecha = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_fecha.pack(fill="x", pady=(0, 5))
        self.txt_fecha_rapido = ctk.CTkEntry(frame_fecha, placeholder_text="DD/MM/AAAA", height=38,
                                             font=("Segoe UI", 11), corner_radius=12)
        self.txt_fecha_rapido.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.txt_fecha_rapido.insert(0, date.today().strftime("%d/%m/%Y"))
        btn_cal = ctk.CTkButton(frame_fecha, text="Calendario", width=65, height=38, font=("Segoe UI", 11),
                                fg_color=AppTheme.PRIMARY, hover_color=AppTheme.PRIMARY_DARK,
                                corner_radius=12, command=self.abrir_calendario_movimiento_rapido)
        btn_cal.pack(side="right")

        frame_botones = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_botones.pack(fill="x", pady=15)
        self.btn_registrar_rapido = ctk.CTkButton(frame_botones, text="REGISTRAR MOVIMIENTO", height=45,
                                                  font=("Segoe UI", 12, "bold"), fg_color=AppTheme.PRIMARY,
                                                  hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                                  command=self.registrar_movimiento_rapido)
        self.btn_registrar_rapido.pack(fill="x", pady=2)

        self.actualizar_label_quien_rapido()
        self.cargar_combo_medicamentos_mov()

    def cargar_combo_medicamentos_mov(self):
        medicamentos = self.medicamento_model.obtener_por_area(self.area_id)
        valores = [f"{m['id_medicamento']} - {m['nombre']}" for m in medicamentos]
        self.combo_medicamento_rapido.configure(values=valores)
        if valores:
            self.combo_medicamento_rapido.set(valores[0])
            self.actualizar_stock_rapido()

    def preseleccionar_medicamento_movimiento_rapido(self, medicamento_id):
        medicamento = self.medicamento_model.obtener_por_id(medicamento_id)
        if medicamento:
            valor = f"{medicamento['id_medicamento']} - {medicamento['nombre']}"
            if valor in self.combo_medicamento_rapido.cget('values'):
                self.combo_medicamento_rapido.set(valor)
                self.actualizar_stock_rapido()
            else:
                self.cargar_combo_medicamentos_mov()
                self.combo_medicamento_rapido.set(valor)
                self.actualizar_stock_rapido()

    def actualizar_stock_rapido(self):
        try:
            seleccion = self.combo_medicamento_rapido.get()
            if seleccion:
                medicamento_id = int(seleccion.split(" - ")[0])
                medicamento = self.medicamento_model.obtener_por_id(medicamento_id)
                if medicamento:
                    stock_actual = medicamento['stock']
                    self.lbl_stock_rapido.configure(text=f"Stock disponible: {stock_actual} unidades")
                    if stock_actual == 0:
                        self.lbl_stock_rapido.configure(text_color="#e53e3e")
                    elif stock_actual < 10:
                        self.lbl_stock_rapido.configure(text_color="#ed8936")
                    else:
                        self.lbl_stock_rapido.configure(text_color=AppTheme.PRIMARY)
        except Exception as e:
            self.debug_print(f"Error actualizando stock: {e}")

    def actualizar_label_quien_rapido(self):
        if self.tipo_mov_rapido.get() == "entrada":
            self.label_quien_rapido.configure(text="Quien Recibe")
        else:
            self.label_quien_rapido.configure(text="Quien Retira")

    def registrar_movimiento_rapido(self):
        try:
            seleccion_med = self.combo_medicamento_rapido.get()
            if not seleccion_med:
                messagebox.showerror("Error", "Seleccione un medicamento")
                return

            medicamento_id = int(seleccion_med.split(" - ")[0])
            cantidad = self.txt_cantidad_rapido.get().strip()
            quien = self.txt_quien_rapido.get().strip()
            motivo = self.txt_motivo_rapido.get().strip()
            tipo = self.tipo_mov_rapido.get()
            fecha_str = self.txt_fecha_rapido.get().strip()

            if not cantidad or not quien or not motivo:
                messagebox.showerror("Error", "Complete todos los campos del movimiento")
                return

            try:
                cantidad_int = int(cantidad)
                if cantidad_int <= 0:
                    messagebox.showerror("Error", "La cantidad debe ser mayor a 0")
                    return
            except ValueError:
                messagebox.showerror("Error", "La cantidad debe ser un numero")
                return

            fecha_mov = date.today()
            if fecha_str:
                try:
                    fecha_mov = datetime.strptime(fecha_str, "%d/%m/%Y").date()
                except ValueError:
                    messagebox.showerror("Error", "Formato de fecha invalido. Use DD/MM/AAAA")
                    return

            medicamento = self.medicamento_model.obtener_por_id(medicamento_id)
            if not medicamento:
                messagebox.showerror("Error", "Medicamento no encontrado")
                return

            if tipo == "salida" and medicamento['stock'] < cantidad_int:
                messagebox.showerror("Error", f"Stock insuficiente.\nStock actual: {medicamento['stock']}")
                return

            quien_recibe = quien if tipo == "entrada" else ""
            quien_retira = quien if tipo == "salida" else ""

            success = self.movimiento_model.registrar(
                medicamento_id, tipo, cantidad_int,
                quien_recibe, quien_retira, motivo, fecha_mov
            )

            if success:
                nuevo_stock = medicamento['stock']
                if tipo == "entrada":
                    nuevo_stock += cantidad_int
                else:
                    nuevo_stock -= cantidad_int
                self.medicamento_model.actualizar_stock(medicamento_id, nuevo_stock)
                messagebox.showinfo("Exito", "Movimiento registrado correctamente")
                self.txt_cantidad_rapido.delete(0, "end")
                self.txt_quien_rapido.delete(0, "end")
                self.txt_motivo_rapido.delete(0, "end")
                self.txt_fecha_rapido.delete(0, "end")
                self.txt_fecha_rapido.insert(0, date.today().strftime("%d/%m/%Y"))
                self.cargar_inventario()
                self.cargar_historial()
                self.cargar_combo_medicamentos_mov()
                self.cargar_combo_ubicaciones()
                if self.id_medicamento == medicamento_id:
                    self.actualizar_stock_rapido()
            else:
                messagebox.showerror("Error", "No se pudo registrar el movimiento")
        except Exception as e:
            messagebox.showerror("Error", f"Error al registrar movimiento: {str(e)}")

    # ------------------- PESTAÑA MOVIMIENTOS -------------------
    def crear_pestana_movimientos(self):
        panel_principal = ctk.CTkFrame(self.tab_movimientos, fg_color="transparent")
        panel_principal.pack(fill="both", expand=True, padx=10, pady=10)

        panel_izquierdo = ctk.CTkFrame(panel_principal, fg_color=AppTheme.SURFACE, corner_radius=15)
        panel_izquierdo.pack(side="left", fill="both", expand=True, padx=(0, 10))

        panel_derecho_mov = ctk.CTkFrame(panel_principal, fg_color=AppTheme.SURFACE, corner_radius=15, width=500)
        panel_derecho_mov.pack(side="right", fill="y", padx=(10, 0))
        panel_derecho_mov.pack_propagate(False)

        self.crear_filtros_movimientos(panel_izquierdo)
        self.crear_tabla_movimientos(panel_izquierdo)
        self.crear_formulario_movimientos(panel_derecho_mov)

    def crear_filtros_movimientos(self, parent):
        frame_filtros = ctk.CTkFrame(parent, fg_color="transparent")
        frame_filtros.pack(fill="x", pady=10, padx=15)

        frame_buscar = ctk.CTkFrame(frame_filtros, fg_color="transparent")
        frame_buscar.pack(fill="x", pady=5)
        self.txt_buscar_mov = ctk.CTkEntry(frame_buscar,
                                           placeholder_text="Buscar por medicamento, motivo, quien recibe o quien retira...",
                                           height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_buscar_mov.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.txt_buscar_mov.bind("<KeyRelease>", self.buscar_historial)

        frame_med = ctk.CTkFrame(frame_filtros, fg_color="transparent")
        frame_med.pack(fill="x", pady=5)
        ctk.CTkLabel(frame_med, text="Medicamento:", font=("Segoe UI", 11, "bold"),
                     text_color=AppTheme.TEXT_PRIMARY).pack(side="left", padx=5)
        self.combo_med_filtro = ctk.CTkComboBox(frame_med, width=400, height=38,
                                                font=("Segoe UI", 11), corner_radius=12,
                                                state="readonly", command=self.cargar_historial)
        self.combo_med_filtro.pack(side="left", padx=10, fill="x", expand=True)

        frame_fechas = ctk.CTkFrame(frame_filtros, fg_color="transparent")
        frame_fechas.pack(fill="x", pady=5)
        ctk.CTkLabel(frame_fechas, text="Desde:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=5)
        self.btn_desde = ctk.CTkButton(frame_fechas, text=self.fecha_desde.strftime("%d/%m/%Y"),
                                       width=110, height=35, font=("Segoe UI", 11),
                                       fg_color=AppTheme.PRIMARY, hover_color=AppTheme.PRIMARY_DARK,
                                       corner_radius=12, command=lambda: self.abrir_calendario("desde"))
        self.btn_desde.pack(side="left", padx=5)
        ctk.CTkLabel(frame_fechas, text="Hasta:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=15)
        self.btn_hasta = ctk.CTkButton(frame_fechas, text=self.fecha_hasta.strftime("%d/%m/%Y"),
                                       width=110, height=35, font=("Segoe UI", 11),
                                       fg_color=AppTheme.PRIMARY, hover_color=AppTheme.PRIMARY_DARK,
                                       corner_radius=12, command=lambda: self.abrir_calendario("hasta"))
        self.btn_hasta.pack(side="left", padx=5)
        ctk.CTkLabel(frame_fechas, text="Tipo:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=15)
        self.combo_tipo_filtro = ctk.CTkComboBox(frame_fechas, values=["Todos", "Entrada", "Salida"],
                                                 width=100, height=35, corner_radius=12, state="readonly",
                                                 command=self.cargar_historial)
        self.combo_tipo_filtro.pack(side="left", padx=5)
        self.combo_tipo_filtro.set("Todos")
        btn_filtrar = ctk.CTkButton(frame_fechas, text="Filtrar", width=100, height=35,
                                    font=("Segoe UI", 11, "bold"), fg_color=AppTheme.PRIMARY,
                                    hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                    command=self.cargar_historial)
        btn_filtrar.pack(side="left", padx=20)
        btn_limpiar = ctk.CTkButton(frame_fechas, text="Limpiar filtros", width=110, height=35,
                                    font=("Segoe UI", 11), fg_color="#4a5568", hover_color="#2d3748",
                                    corner_radius=12, command=self.limpiar_filtros_historial)
        btn_limpiar.pack(side="left", padx=5)

        btn_exportar = ctk.CTkButton(frame_fechas, text="Exportar Movimientos", width=150, height=35,
                                     font=("Segoe UI", 11), fg_color="#2c7a4d",
                                     hover_color="#1e5a38", corner_radius=12,
                                     command=self.exportar_excel_movimientos)
        btn_exportar.pack(side="left", padx=20)

    def crear_tabla_movimientos(self, parent):
        frame_tabla = ctk.CTkFrame(parent, fg_color="transparent")
        frame_tabla.pack(fill="both", expand=True, pady=10, padx=15)
        scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical")
        scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal")
        self.tabla_mov = ttk.Treeview(
            frame_tabla,
            columns=("ID", "Medicamento", "Tipo", "Cantidad", "Quien Recibe", "Quien Retira", "Motivo", "Fecha"),
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
            style="Movements.Treeview",
            height=18
        )
        scroll_y.config(command=self.tabla_mov.yview)
        scroll_x.config(command=self.tabla_mov.xview)
        columnas = [("ID", 60, "center"), ("Medicamento", 220, "w"), ("Tipo", 110, "center"),
                    ("Cantidad", 90, "center"), ("Quien Recibe", 200, "w"), ("Quien Retira", 200, "w"),
                    ("Motivo", 350, "w"), ("Fecha", 110, "center")]
        for col, width, anchor in columnas:
            self.tabla_mov.heading(col, text=col)
            self.tabla_mov.column(col, width=width, anchor=anchor)
        self.tabla_mov.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)
        self.tabla_mov.bind("<<TreeviewSelect>>", self.on_movimiento_seleccionado)

        self.tabla_mov.tag_configure('stock_cero', background='#FFF9C4')
        self.tabla_mov.tag_configure('reciente', background='#C8E6C9')
        self.tabla_mov.tag_configure('caducado', background='#FFCDD2')
        self.tabla_mov.tag_configure('proximo_caducar', background='#FFE0B2')

    def limpiar_formulario_movimientos(self):
        self.id_movimiento_editando = None
        self.form_titulo_mov.configure(text="REGISTRAR MOVIMIENTO")
        self.tipo_mov.set("entrada")
        self.txt_cantidad_mov.delete(0, "end")
        self.txt_quien_mov.delete(0, "end")
        self.txt_motivo_mov.delete(0, "end")
        self.txt_fecha_mov.delete(0, "end")
        self.txt_fecha_mov.insert(0, date.today().strftime("%d/%m/%Y"))
        self.btn_guardar_mov.configure(state="normal")
        self.btn_actualizar_mov.configure(state="disabled")
        self.btn_eliminar_mov.configure(state="disabled")
        self.cargar_combo_movimientos()
        for item in self.tabla_mov.selection():
            self.tabla_mov.selection_remove(item)

    def cargar_combo_movimientos(self):
        medicamentos = self.medicamento_model.obtener_por_area(self.area_id)
        valores = [f"{m['id_medicamento']} - {m['nombre']}" for m in medicamentos]
        self.combo_med_mov.configure(values=valores)
        if valores:
            self.combo_med_mov.set(valores[0])
            self.actualizar_stock_mov()

    def actualizar_stock_mov(self):
        try:
            seleccion = self.combo_med_mov.get()
            if seleccion:
                medicamento_id = int(seleccion.split(" - ")[0])
                medicamento = self.medicamento_model.obtener_por_id(medicamento_id)
                if medicamento:
                    stock_actual = medicamento['stock']
                    self.lbl_stock_mov.configure(text=f"Stock disponible: {stock_actual} unidades")
                    if stock_actual == 0:
                        self.lbl_stock_mov.configure(text_color="#e53e3e")
                    elif stock_actual < 10:
                        self.lbl_stock_mov.configure(text_color="#ed8936")
                    else:
                        self.lbl_stock_mov.configure(text_color=AppTheme.PRIMARY)
        except Exception as e:
            self.debug_print(f"Error actualizando stock: {e}")

    def actualizar_label_quien_mov(self):
        if self.tipo_mov.get() == "entrada":
            self.label_quien_mov.configure(text="Quien Recibe")
        else:
            self.label_quien_mov.configure(text="Quien Retira")

    def cargar_combo_medicamentos_filtro(self):
        try:
            medicamentos = self.medicamento_model.obtener_por_area(self.area_id)
            valores = ["-- TODOS LOS MEDICAMENTOS --"] + [f"{m['id_medicamento']} - {m['nombre']}" for m in
                                                          medicamentos]
            self.combo_med_filtro.configure(values=valores)
            if valores:
                self.combo_med_filtro.set(valores[0])
        except Exception as e:
            self.debug_print(f"Error al cargar combo filtro: {e}")

    def cargar_historial(self, event=None):
        try:
            for item in self.tabla_mov.get_children():
                self.tabla_mov.delete(item)

            busqueda = self.txt_buscar_mov.get().strip().lower()
            seleccion = self.combo_med_filtro.get()
            if not seleccion:
                return

            fecha_desde = self.fecha_desde if hasattr(self, 'fecha_desde') else date(date.today().year, 1, 1)
            fecha_hasta = self.fecha_hasta if hasattr(self, 'fecha_hasta') else date.today()
            tipo_filtro = self.combo_tipo_filtro.get() if self.combo_tipo_filtro.get() != "Todos" else None

            movimientos = self.movimiento_model.obtener_todos_por_area(
                self.area_id, fecha_desde, fecha_hasta, tipo_filtro
            )

            if not movimientos:
                self.tabla_mov.insert("", "end", values=("", "", "", "", "", "", "No hay movimientos para mostrar", ""))
                return

            for mov in reversed(movimientos):
                nombre_med = mov.get('medicamento_nombre', '')
                motivo = mov.get('motivo', '')
                quien_recibe = mov.get('quien_recibe', '---')
                quien_retira = mov.get('quien_retira', '---')

                if busqueda:
                    if (busqueda not in motivo.lower() and
                            busqueda not in quien_recibe.lower() and
                            busqueda not in quien_retira.lower() and
                            busqueda not in nombre_med.lower()):
                        continue

                fecha = mov['fecha'].strftime("%d/%m/%Y") if isinstance(mov['fecha'], date) else str(mov['fecha'])
                tipo = "ENTRADA" if mov['tipo'] == "entrada" else "SALIDA"

                medicamento_actual = self.medicamento_model.obtener_por_id(mov.get('id_medicamento'))
                tag = self.obtener_color_movimiento(mov, medicamento_actual)

                self.tabla_mov.insert("", "end", values=(
                    mov['id_movimiento'], nombre_med, tipo, mov['cantidad'],
                    quien_recibe, quien_retira, motivo, fecha
                ), tags=(tag,) if tag else ())
        except Exception as e:
            self.debug_print(f"Error cargando historial: {e}")
            messagebox.showerror("Error", f"Error cargando historial: {str(e)}")

    def obtener_color_movimiento(self, movimiento, medicamento_actual):
        fecha_hoy = date.today()
        if movimiento['tipo'] == 'salida' and medicamento_actual and medicamento_actual.get('stock', 0) == 0:
            return 'stock_cero'
        if movimiento['fecha'] == fecha_hoy:
            return 'reciente'
        if medicamento_actual and medicamento_actual.get('caducidad'):
            if medicamento_actual['caducidad'] < fecha_hoy:
                return 'caducado'
            elif (medicamento_actual['caducidad'] - fecha_hoy).days <= 30:
                return 'proximo_caducar'
        return ''

    def buscar_historial(self, event):
        self.cargar_historial()

    def limpiar_filtros_historial(self):
        self.fecha_desde = date(date.today().year, 1, 1)
        self.fecha_hasta = date.today()
        self.btn_desde.configure(text=self.fecha_desde.strftime("%d/%m/%Y"))
        self.btn_hasta.configure(text=self.fecha_hasta.strftime("%d/%m/%Y"))
        self.combo_tipo_filtro.set("Todos")
        self.txt_buscar_mov.delete(0, "end")
        self.cargar_historial()

    def on_movimiento_seleccionado(self, event):
        seleccion = self.tabla_mov.selection()
        if seleccion:
            valores = self.tabla_mov.item(seleccion[0])['values']
            if valores and valores[0] != "":
                self.id_movimiento_editando = valores[0]
                self.cargar_movimiento_formulario(valores[0])
                self.btn_actualizar_mov.configure(state="normal")
                self.btn_eliminar_mov.configure(state="normal")
                self.btn_guardar_mov.configure(state="disabled")

    def cargar_movimiento_formulario(self, movimiento_id):
        try:
            movimiento = self.movimiento_model.obtener_por_id(movimiento_id, self.area_id)
            if movimiento:
                self.form_titulo_mov.configure(text="EDITAR MOVIMIENTO")
                if movimiento['tipo'] == "entrada":
                    self.tipo_mov.set("entrada")
                else:
                    self.tipo_mov.set("salida")
                self.txt_cantidad_mov.delete(0, "end")
                self.txt_cantidad_mov.insert(0, str(movimiento['cantidad']))
                self.txt_quien_mov.delete(0, "end")
                if movimiento['tipo'] == "entrada":
                    self.txt_quien_mov.insert(0, movimiento['quien_recibe'] or "")
                else:
                    self.txt_quien_mov.insert(0, movimiento['quien_retira'] or "")
                self.txt_motivo_mov.delete(0, "end")
                self.txt_motivo_mov.insert(0, movimiento['motivo'] or "")
                self.txt_fecha_mov.delete(0, "end")
                self.txt_fecha_mov.insert(0, movimiento['fecha'].strftime("%d/%m/%Y") if movimiento['fecha'] else "")

                medicamentos = self.medicamento_model.obtener_por_area(self.area_id)
                for med in medicamentos:
                    if med['id_medicamento'] == movimiento['id_medicamento']:
                        self.combo_med_mov.set(f"{med['id_medicamento']} - {med['nombre']}")
                        self.actualizar_stock_mov()
                        break
        except Exception as e:
            self.debug_print(f"Error al cargar movimiento: {e}")

    def crear_formulario_movimientos(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent", height=750)
        scroll.pack(fill="both", expand=True, padx=15, pady=15)

        self.form_titulo_mov = ctk.CTkLabel(scroll, text="REGISTRAR MOVIMIENTO",
                                            font=("Segoe UI", 18, "bold"), text_color=AppTheme.PRIMARY)
        self.form_titulo_mov.pack(pady=(0, 15))

        ctk.CTkLabel(scroll, text="Medicamento", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.combo_med_mov = ctk.CTkComboBox(scroll, values=[], width=400, height=38,
                                             font=("Segoe UI", 11), corner_radius=12, state="readonly",
                                             command=self.actualizar_stock_mov)
        self.combo_med_mov.pack(fill="x", pady=(0, 5))

        self.lbl_stock_mov = ctk.CTkLabel(scroll, text="Stock disponible: -- unidades",
                                          font=("Segoe UI", 11), text_color=AppTheme.PRIMARY)
        self.lbl_stock_mov.pack(anchor="w", pady=(0, 10))

        ctk.CTkLabel(scroll, text="Tipo de Movimiento", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.tipo_mov = ctk.StringVar(value="entrada")
        frame_tipo = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_tipo.pack(fill="x", pady=5)
        radio_entrada = ctk.CTkRadioButton(frame_tipo, text="ENTRADA", variable=self.tipo_mov,
                                           value="entrada", font=("Segoe UI", 11, "bold"),
                                           fg_color="#2c7a4d", hover_color="#1e5a38",
                                           text_color=AppTheme.TEXT_PRIMARY,
                                           command=self.actualizar_label_quien_mov)
        radio_entrada.pack(side="left", padx=10)
        radio_salida = ctk.CTkRadioButton(frame_tipo, text="SALIDA", variable=self.tipo_mov,
                                          value="salida", font=("Segoe UI", 11, "bold"),
                                          fg_color="#e53e3e", hover_color="#c53030",
                                          text_color=AppTheme.TEXT_PRIMARY,
                                          command=self.actualizar_label_quien_mov)
        radio_salida.pack(side="left", padx=10)

        ctk.CTkLabel(scroll, text="Cantidad", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_cantidad_mov = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_cantidad_mov.pack(fill="x", pady=(0, 5))

        self.label_quien_mov = ctk.CTkLabel(scroll, text="Quien Recibe", font=("Segoe UI", 11, "bold"))
        self.label_quien_mov.pack(anchor="w", pady=(10, 2))
        self.txt_quien_mov = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_quien_mov.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Motivo", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_motivo_mov = ctk.CTkEntry(scroll, height=38, font=("Segoe UI", 11), corner_radius=12)
        self.txt_motivo_mov.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Fecha", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 2))
        frame_fecha = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_fecha.pack(fill="x", pady=(0, 5))
        self.txt_fecha_mov = ctk.CTkEntry(frame_fecha, placeholder_text="DD/MM/AAAA", height=38,
                                          font=("Segoe UI", 11), corner_radius=12)
        self.txt_fecha_mov.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.txt_fecha_mov.insert(0, date.today().strftime("%d/%m/%Y"))
        btn_cal = ctk.CTkButton(frame_fecha, text="Calendario", width=65, height=38, font=("Segoe UI", 11),
                                fg_color=AppTheme.PRIMARY, hover_color=AppTheme.PRIMARY_DARK,
                                corner_radius=12, command=self.abrir_calendario_movimiento)
        btn_cal.pack(side="right")

        frame_botones = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_botones.pack(fill="x", pady=15)
        self.btn_guardar_mov = ctk.CTkButton(frame_botones, text="REGISTRAR MOVIMIENTO", height=45,
                                             font=("Segoe UI", 12, "bold"), fg_color=AppTheme.PRIMARY,
                                             hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                             command=self.registrar_movimiento)
        self.btn_guardar_mov.pack(fill="x", pady=2)
        self.btn_actualizar_mov = ctk.CTkButton(frame_botones, text="ACTUALIZAR MOVIMIENTO", height=45,
                                                font=("Segoe UI", 12, "bold"), fg_color=AppTheme.PRIMARY,
                                                hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                                command=self.actualizar_movimiento, state="disabled")
        self.btn_actualizar_mov.pack(fill="x", pady=2)
        self.btn_eliminar_mov = ctk.CTkButton(frame_botones, text="ELIMINAR MOVIMIENTO", height=45,
                                              font=("Segoe UI", 12), fg_color="#e53e3e",
                                              hover_color="#c53030", text_color=AppTheme.TEXT_ON_PRIMARY,
                                              corner_radius=12, command=self.eliminar_movimiento, state="disabled")
        self.btn_eliminar_mov.pack(fill="x", pady=2)
        self.btn_limpiar_mov = ctk.CTkButton(frame_botones, text="LIMPIAR FORMULARIO", height=40,
                                             font=("Segoe UI", 12), fg_color="#4a5568",
                                             hover_color="#2d3748", text_color=AppTheme.TEXT_ON_PRIMARY,
                                             corner_radius=12, command=self.limpiar_formulario_movimientos)
        self.btn_limpiar_mov.pack(fill="x", pady=2)

        self.actualizar_label_quien_mov()
        self.cargar_combo_movimientos()

    def registrar_movimiento(self):
        try:
            seleccion_med = self.combo_med_mov.get()
            if not seleccion_med:
                messagebox.showerror("Error", "Seleccione un medicamento")
                return

            medicamento_id = int(seleccion_med.split(" - ")[0])
            cantidad = self.txt_cantidad_mov.get().strip()
            quien = self.txt_quien_mov.get().strip()
            motivo = self.txt_motivo_mov.get().strip()
            tipo = self.tipo_mov.get()
            fecha_str = self.txt_fecha_mov.get().strip()

            if not cantidad or not quien or not motivo:
                messagebox.showerror("Error", "Complete todos los campos del movimiento")
                return

            try:
                cantidad_int = int(cantidad)
                if cantidad_int <= 0:
                    messagebox.showerror("Error", "La cantidad debe ser mayor a 0")
                    return
            except ValueError:
                messagebox.showerror("Error", "La cantidad debe ser un numero")
                return

            fecha_mov = date.today()
            if fecha_str:
                try:
                    fecha_mov = datetime.strptime(fecha_str, "%d/%m/%Y").date()
                except ValueError:
                    messagebox.showerror("Error", "Formato de fecha invalido. Use DD/MM/AAAA")
                    return

            medicamento = self.medicamento_model.obtener_por_id(medicamento_id)
            if not medicamento:
                messagebox.showerror("Error", "Medicamento no encontrado")
                return

            if tipo == "salida" and medicamento['stock'] < cantidad_int:
                messagebox.showerror("Error", f"Stock insuficiente.\nStock actual: {medicamento['stock']}")
                return

            quien_recibe = quien if tipo == "entrada" else ""
            quien_retira = quien if tipo == "salida" else ""

            success = self.movimiento_model.registrar(
                medicamento_id, tipo, cantidad_int,
                quien_recibe, quien_retira, motivo, fecha_mov
            )

            if success:
                nuevo_stock = medicamento['stock']
                if tipo == "entrada":
                    nuevo_stock += cantidad_int
                else:
                    nuevo_stock -= cantidad_int
                self.medicamento_model.actualizar_stock(medicamento_id, nuevo_stock)
                messagebox.showinfo("Exito", "Movimiento registrado correctamente")
                self.limpiar_formulario_movimientos()
                self.cargar_historial()
                self.cargar_inventario()
                self.cargar_combo_movimientos()
                self.cargar_combo_medicamentos_filtro()
                self.cargar_combo_ubicaciones()
                self.cargar_combo_medicamentos_mov()
            else:
                messagebox.showerror("Error", "No se pudo registrar el movimiento")
        except Exception as e:
            messagebox.showerror("Error", f"Error al registrar movimiento: {str(e)}")

    def actualizar_movimiento(self):
        try:
            if not self.id_movimiento_editando:
                messagebox.showerror("Error", "No hay movimiento seleccionado")
                return

            seleccion_med = self.combo_med_mov.get()
            if not seleccion_med:
                messagebox.showerror("Error", "Seleccione un medicamento")
                return

            medicamento_id = int(seleccion_med.split(" - ")[0])
            nueva_cantidad = int(self.txt_cantidad_mov.get().strip())
            nuevo_tipo = self.tipo_mov.get()
            nuevo_quien = self.txt_quien_mov.get().strip()
            nuevo_motivo = self.txt_motivo_mov.get().strip()
            nueva_fecha_str = self.txt_fecha_mov.get().strip()

            if nueva_cantidad <= 0:
                messagebox.showerror("Error", "La cantidad debe ser mayor a 0")
                return

            if not nuevo_motivo:
                messagebox.showerror("Error", "El motivo es obligatorio")
                return

            nueva_fecha = date.today()
            if nueva_fecha_str:
                try:
                    nueva_fecha = datetime.strptime(nueva_fecha_str, "%d/%m/%Y").date()
                except ValueError:
                    messagebox.showerror("Error", "Formato de fecha invalido. Use DD/MM/AAAA")
                    return

            movimiento_original = self.movimiento_model.obtener_por_id(self.id_movimiento_editando, self.area_id)
            if not movimiento_original:
                messagebox.showerror("Error", "Movimiento no encontrado")
                return

            medicamento = self.medicamento_model.obtener_por_id(medicamento_id)
            stock_actual = medicamento['stock']

            if movimiento_original['tipo'] == "entrada":
                stock_actual -= movimiento_original['cantidad']
            else:
                stock_actual += movimiento_original['cantidad']

            if nuevo_tipo == "entrada":
                nuevo_stock = stock_actual + nueva_cantidad
            else:
                if stock_actual < nueva_cantidad:
                    messagebox.showerror("Error", f"Stock insuficiente. Stock actual: {stock_actual}")
                    return
                nuevo_stock = stock_actual - nueva_cantidad

            quien_recibe = nuevo_quien if nuevo_tipo == "entrada" else ""
            quien_retira = nuevo_quien if nuevo_tipo == "salida" else ""

            tabla_movimientos = "movimientos_farmacia" if self.area_id == 1 else "movimientos_odontologia"
            query = f"""
                UPDATE {tabla_movimientos}
                SET tipo = ?, cantidad = ?, quien_recibe = ?, quien_retira = ?, motivo = ?, fecha = ?, id_inventario = ?
                WHERE id_movimiento = ?
            """
            params = (nuevo_tipo, nueva_cantidad, quien_recibe, quien_retira,
                      nuevo_motivo, nueva_fecha, medicamento_id, self.id_movimiento_editando)
            success = Database.execute_non_query(query, params)

            if success:
                self.medicamento_model.actualizar_stock(medicamento_id, nuevo_stock)
                messagebox.showinfo("Exito", "Movimiento actualizado correctamente")
                self.limpiar_formulario_movimientos()
                self.cargar_historial()
                self.cargar_inventario()
                self.cargar_combo_movimientos()
                self.cargar_combo_medicamentos_filtro()
                self.cargar_combo_ubicaciones()
                self.cargar_combo_medicamentos_mov()
            else:
                messagebox.showerror("Error", "No se pudo actualizar el movimiento")
        except ValueError:
            messagebox.showerror("Error", "La cantidad debe ser un numero")
        except Exception as e:
            messagebox.showerror("Error", f"Error al actualizar: {str(e)}")

    def eliminar_movimiento(self):
        if not self.id_movimiento_editando:
            messagebox.showerror("Error", "No hay movimiento seleccionado")
            return

        try:
            movimiento = self.movimiento_model.obtener_por_id(self.id_movimiento_editando, self.area_id)
            if not movimiento:
                messagebox.showerror("Error", "Movimiento no encontrado")
                return

            medicamento = self.medicamento_model.obtener_por_id(movimiento['id_medicamento'])

            if movimiento['tipo'] == "entrada":
                nuevo_stock = medicamento['stock'] - movimiento['cantidad']
                if nuevo_stock < 0:
                    messagebox.showerror("Error", "No se puede eliminar este movimiento porque dejaria stock negativo")
                    return
            else:
                nuevo_stock = medicamento['stock'] + movimiento['cantidad']

            tabla_movimientos = "movimientos_farmacia" if self.area_id == 1 else "movimientos_odontologia"
            query = f"DELETE FROM {tabla_movimientos} WHERE id_movimiento = ?"
            success = Database.execute_non_query(query, (self.id_movimiento_editando,))

            if success:
                self.medicamento_model.actualizar_stock(movimiento['id_medicamento'], nuevo_stock)
                messagebox.showinfo("Exito", "Movimiento eliminado correctamente")
                self.limpiar_formulario_movimientos()
                self.cargar_historial()
                self.cargar_inventario()
                self.cargar_combo_movimientos()
                self.cargar_combo_medicamentos_filtro()
                self.cargar_combo_ubicaciones()
                self.cargar_combo_medicamentos_mov()
            else:
                messagebox.showerror("Error", "No se pudo eliminar el movimiento")
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar movimiento: {str(e)}")

    # ---------- Exportar Excel ----------
    def exportar_excel_inventario(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Error", "Instale openpyxl: pip install openpyxl")
            return
        try:
            medicamentos = self.medicamento_model.obtener_por_area(self.area_id, busqueda="")
            if not medicamentos:
                messagebox.showwarning("Sin datos", "No hay medicamentos para exportar.")
                return
            fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_por_defecto = f"MediKinv_{self.area_nombre}_Inventario_{fecha_hora}.xlsx"
            archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos Excel", "*.xlsx")],
                initialfile=nombre_por_defecto,
                title="Guardar exportacion de inventario"
            )
            if not archivo:
                return
            wb = Workbook()
            ws = wb.active
            ws.title = "Inventario"
            titulo_font = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
            titulo_fill = PatternFill(start_color='2C7A4D', end_color='2C7A4D', fill_type='solid')
            header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='4A5568', end_color='4A5568', fill_type='solid')
            center_alignment = Alignment(horizontal='center', vertical='center')
            left_alignment = Alignment(horizontal='left', vertical='center')
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            ws.merge_cells('A1:I1')
            ws['A1'].value = "Universitarios en Accion"
            ws['A1'].font = titulo_font
            ws['A1'].fill = titulo_fill
            ws['A1'].alignment = center_alignment
            ws.merge_cells('A2:I2')
            ws.row_dimensions[2].height = 10
            headers = ["Medicamento", "Presentacion", "Proveedor", "Entrada", "Salidas", "Existencias",
                       "Ubicacion", "Caducidad mas reciente", "Anotaciones"]
            for col_idx, header in enumerate(headers, start=1):
                celda = ws.cell(row=3, column=col_idx, value=header)
                celda.font = header_font
                celda.fill = header_fill
                celda.alignment = center_alignment
                celda.border = thin_border
            ws.column_dimensions['A'].width = 30
            ws.column_dimensions['B'].width = 25
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 12
            ws.column_dimensions['E'].width = 12
            ws.column_dimensions['F'].width = 12
            ws.column_dimensions['G'].width = 20
            ws.column_dimensions['H'].width = 18
            ws.column_dimensions['I'].width = 35
            fila_actual = 4
            for med in medicamentos:
                med_id = med['id_medicamento']
                movimientos = self.movimiento_model.obtener_por_medicamento(med_id)
                total_entradas = sum(m['cantidad'] for m in movimientos if m['tipo'] == 'entrada')
                total_salidas = sum(m['cantidad'] for m in movimientos if m['tipo'] == 'salida')
                existencias = med['stock']
                caducidad = med['caducidad'].strftime("%d/%m/%Y") if med['caducidad'] else "---"
                anotaciones = med.get('observaciones', '') or ''
                ws.cell(row=fila_actual, column=1, value=med['nombre']).alignment = left_alignment
                ws.cell(row=fila_actual, column=2, value=med['presentacion']).alignment = left_alignment
                ws.cell(row=fila_actual, column=3, value="")
                ws.cell(row=fila_actual, column=4, value=total_entradas).alignment = center_alignment
                ws.cell(row=fila_actual, column=5, value=total_salidas).alignment = center_alignment
                ws.cell(row=fila_actual, column=6, value=existencias).alignment = center_alignment
                ws.cell(row=fila_actual, column=7, value=med['ubicacion'] or '').alignment = left_alignment
                ws.cell(row=fila_actual, column=8, value=caducidad).alignment = center_alignment
                ws.cell(row=fila_actual, column=9, value=anotaciones).alignment = left_alignment
                for col in range(1, 10):
                    ws.cell(row=fila_actual, column=col).border = thin_border
                fila_actual += 1
            wb.save(archivo)
            messagebox.showinfo("Exportacion exitosa", f"Inventario exportado a:\n{archivo}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar inventario:\n{str(e)}")

    def exportar_excel_movimientos(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Error", "Instale openpyxl: pip install openpyxl")
            return
        try:
            movimientos_data = []
            for child in self.tabla_mov.get_children():
                valores = self.tabla_mov.item(child)['values']
                if valores and valores[0] != "":
                    movimientos_data.append({
                        'id': valores[0],
                        'medicamento': valores[1],
                        'tipo': valores[2],
                        'cantidad': valores[3],
                        'quien_recibe': valores[4],
                        'quien_retira': valores[5],
                        'motivo': valores[6],
                        'fecha': valores[7]
                    })
            if not movimientos_data:
                messagebox.showwarning("Sin datos", "No hay movimientos para exportar con los filtros actuales.")
                return
            fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_por_defecto = f"MediKinv_{self.area_nombre}_Movimientos_{fecha_hora}.xlsx"
            archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos Excel", "*.xlsx")],
                initialfile=nombre_por_defecto,
                title="Guardar exportacion de movimientos"
            )
            if not archivo:
                return
            wb = Workbook()
            ws = wb.active
            ws.title = "Movimientos"
            titulo_font = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
            titulo_fill = PatternFill(start_color='2C7A4D', end_color='2C7A4D', fill_type='solid')
            header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
            header_fill = PatternFill(start_color='4A5568', end_color='4A5568', fill_type='solid')
            center_alignment = Alignment(horizontal='center', vertical='center')
            left_alignment = Alignment(horizontal='left', vertical='center')
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            ws.merge_cells('A1:H1')
            ws['A1'].value = "Universitarios en Accion - Movimientos"
            ws['A1'].font = titulo_font
            ws['A1'].fill = titulo_fill
            ws['A1'].alignment = center_alignment
            ws.merge_cells('A2:H2')
            ws.row_dimensions[2].height = 10
            headers = ["ID", "Medicamento", "Tipo", "Cantidad", "Quien Recibe", "Quien Retira", "Motivo", "Fecha"]
            for col_idx, header in enumerate(headers, start=1):
                celda = ws.cell(row=3, column=col_idx, value=header)
                celda.font = header_font
                celda.fill = header_fill
                celda.alignment = center_alignment
                celda.border = thin_border
            ws.column_dimensions['A'].width = 10
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 12
            ws.column_dimensions['E'].width = 25
            ws.column_dimensions['F'].width = 25
            ws.column_dimensions['G'].width = 40
            ws.column_dimensions['H'].width = 15
            fila_actual = 4
            for mov in movimientos_data:
                ws.cell(row=fila_actual, column=1, value=mov['id']).alignment = center_alignment
                ws.cell(row=fila_actual, column=2, value=mov['medicamento']).alignment = left_alignment
                ws.cell(row=fila_actual, column=3, value=mov['tipo']).alignment = center_alignment
                ws.cell(row=fila_actual, column=4, value=mov['cantidad']).alignment = center_alignment
                ws.cell(row=fila_actual, column=5, value=mov['quien_recibe']).alignment = left_alignment
                ws.cell(row=fila_actual, column=6, value=mov['quien_retira']).alignment = left_alignment
                ws.cell(row=fila_actual, column=7, value=mov['motivo']).alignment = left_alignment
                ws.cell(row=fila_actual, column=8, value=mov['fecha']).alignment = center_alignment
                for col in range(1, 9):
                    ws.cell(row=fila_actual, column=col).border = thin_border
                fila_actual += 1
            wb.save(archivo)
            messagebox.showinfo("Exportacion exitosa", f"Movimientos exportados a:\n{archivo}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar movimientos:\n{str(e)}")

    # ---------- Calendarios ----------
    def abrir_calendario_medicamento(self):
        self._abrir_calendario_generico("medicamento")

    def abrir_calendario_movimiento(self):
        self._abrir_calendario_generico("movimiento")

    def abrir_calendario_movimiento_rapido(self):
        self._abrir_calendario_generico("movimiento_rapido")

    def abrir_calendario(self, tipo):
        self._abrir_calendario_generico(tipo)

    def _abrir_calendario_generico(self, destino):
        cal_window = ctk.CTkToplevel(self)
        cal_window.title("Seleccionar fecha")
        cal_window.geometry("350x450")
        cal_window.resizable(False, False)
        cal_window.grab_set()
        cal_window.configure(fg_color=AppTheme.PRIMARY_DARK)
        cal_window.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 175
        y = self.winfo_y() + (self.winfo_height() // 2) - 225
        cal_window.geometry(f"+{x}+{y}")

        frame_cal = ctk.CTkFrame(cal_window, fg_color=AppTheme.SURFACE, corner_radius=20)
        frame_cal.pack(fill="both", expand=True, padx=15, pady=15)

        self.cal_year = datetime.now().year
        self.cal_month = datetime.now().month
        self.cal_destino = destino
        self.cal_window = cal_window

        nav_frame = ctk.CTkFrame(frame_cal, fg_color="transparent")
        nav_frame.pack(fill="x", pady=10)

        btn_prev = ctk.CTkButton(nav_frame, text="<", width=35, height=35, fg_color=AppTheme.PRIMARY,
                                 hover_color=AppTheme.PRIMARY_DARK,
                                 command=lambda: self.cambiar_mes_calendario(-1, frame_cal))
        btn_prev.pack(side="left", padx=5)

        self.lbl_mes_anio = ctk.CTkLabel(nav_frame,
                                         text=self.obtener_nombre_mes(self.cal_month) + " " + str(self.cal_year),
                                         font=("Segoe UI", 16, "bold"), text_color=AppTheme.PRIMARY)
        self.lbl_mes_anio.pack(side="left", expand=True)

        btn_next = ctk.CTkButton(nav_frame, text=">", width=35, height=35, fg_color=AppTheme.PRIMARY,
                                 hover_color=AppTheme.PRIMARY_DARK,
                                 command=lambda: self.cambiar_mes_calendario(1, frame_cal))
        btn_next.pack(side="right", padx=5)

        dias_semana = ctk.CTkFrame(frame_cal, fg_color="transparent")
        dias_semana.pack(fill="x", pady=10)
        for dia in ["Lun", "Mar", "Mie", "Jue", "Vie", "Sab", "Dom"]:
            label = ctk.CTkLabel(dias_semana, text=dia, width=45, font=("Segoe UI", 11, "bold"),
                                 text_color=AppTheme.PRIMARY)
            label.pack(side="left", expand=True, fill="x")

        self.dias_frame_cal = ctk.CTkFrame(frame_cal, fg_color="transparent")
        self.dias_frame_cal.pack(fill="both", expand=True, pady=5)

        self.actualizar_calendario()

        btn_aceptar = ctk.CTkButton(frame_cal, text="Aceptar", width=100, height=35,
                                    font=("Segoe UI", 11, "bold"), fg_color=AppTheme.PRIMARY,
                                    hover_color=AppTheme.PRIMARY_DARK, command=cal_window.destroy)
        btn_aceptar.pack(pady=10)

    def obtener_nombre_mes(self, mes):
        meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                 "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        return meses[mes - 1]

    def cambiar_mes_calendario(self, delta, frame_cal):
        self.cal_month += delta
        if self.cal_month > 12:
            self.cal_month = 1
            self.cal_year += 1
        elif self.cal_month < 1:
            self.cal_month = 12
            self.cal_year -= 1
        self.lbl_mes_anio.configure(text=f"{self.obtener_nombre_mes(self.cal_month)} {self.cal_year}")
        self.actualizar_calendario()

    def actualizar_calendario(self):
        for widget in self.dias_frame_cal.winfo_children():
            widget.destroy()

        cal = calendar.monthcalendar(self.cal_year, self.cal_month)
        hoy = date.today()

        for semana in cal:
            frame_semana = ctk.CTkFrame(self.dias_frame_cal, fg_color="transparent")
            frame_semana.pack(fill="x", pady=2)

            for dia in semana:
                if dia == 0:
                    btn = ctk.CTkButton(frame_semana, text="", width=45, height=38,
                                        fg_color="transparent", state="disabled")
                else:
                    fecha = date(self.cal_year, self.cal_month, dia)
                    if fecha == hoy:
                        bg_color = "#F39C12"
                        text_color = "#FFFFFF"
                    else:
                        bg_color = "transparent"
                        text_color = AppTheme.TEXT_PRIMARY

                    btn = ctk.CTkButton(
                        frame_semana, text=str(dia), width=45, height=38,
                        font=("Segoe UI", 11), fg_color=bg_color,
                        hover_color=AppTheme.PRIMARY_LIGHT, text_color=text_color,
                        corner_radius=8,
                        command=lambda d=dia: self.seleccionar_dia_calendario(d)
                    )
                btn.pack(side="left", expand=True, fill="x", padx=2)

    def seleccionar_dia_calendario(self, dia):
        fecha = date(self.cal_year, self.cal_month, dia)

        if self.cal_destino == "medicamento":
            self.txt_caducidad.delete(0, "end")
            self.txt_caducidad.insert(0, fecha.strftime("%d/%m/%Y"))
        elif self.cal_destino == "movimiento":
            self.txt_fecha_mov.delete(0, "end")
            self.txt_fecha_mov.insert(0, fecha.strftime("%d/%m/%Y"))
        elif self.cal_destino == "movimiento_rapido":
            self.txt_fecha_rapido.delete(0, "end")
            self.txt_fecha_rapido.insert(0, fecha.strftime("%d/%m/%Y"))
        elif self.cal_destino == "desde":
            self.fecha_desde = fecha
            self.btn_desde.configure(text=fecha.strftime("%d/%m/%Y"))
            self.cargar_historial()
        elif self.cal_destino == "hasta":
            self.fecha_hasta = fecha
            self.btn_hasta.configure(text=fecha.strftime("%d/%m/%Y"))
            self.cargar_historial()

        self.cal_window.destroy()


if __name__ == "__main__":
    app = SistemaView(None, 1, "Farmacia")
    app.mainloop()