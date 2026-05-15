import customtkinter as ctk
from tkinter import messagebox, ttk
from assets.styles.themes import AppTheme
from config.database import Database
from utils.icon_manager import IconManager


class CatalogoMedicamentosView(ctk.CTk):

    def __init__(self, parent, usuario):
        super().__init__()
        IconManager.aplicar_icono(self)

        self.parent = parent
        self.usuario = usuario
        self.medicamento_editando = None
        self._debug = True

        AppTheme.aplicar_tema()

        self.title("MediKinv - Catálogo de Medicamentos")

        # Configuración responsiva
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = int(screen_width * 0.9)
        window_height = int(screen_height * 0.85)
        min_width = 900
        min_height = 600

        final_width = max(window_width, min_width)
        final_height = max(window_height, min_height)

        self.geometry(f"{final_width}x{final_height}")
        self.minsize(min_width, min_height)

        # Maximizar por defecto
        try:
            self.state('zoomed')
        except:
            pass

        self.configure(fg_color=AppTheme.SECONDARY)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.bind('<Escape>', self.toggle_fullscreen)
        self.bind('<Configure>', self.on_window_resize)

        self.after(100, self.cargar_datos_iniciales)

    def debug_print(self, msg):
        if self._debug:
            print(f"[DEBUG CATALOGO] {msg}")

    def on_window_resize(self, event=None):
        if event and event.widget == self:
            self.after(100, self.ajustar_tabla)

    def ajustar_tabla(self):
        if hasattr(self, 'tree'):
            altura = self.winfo_height() - 250
            altura_filas = max(15, altura // 30)
            self.tree.configure(height=altura_filas)

            ancho = self.winfo_width()
            self.ajustar_columnas(ancho)

    def ajustar_columnas(self, ancho):
        if ancho < 1000:
            self.tree.column("ID", width=50)
            self.tree.column("Nombre", width=180)
            self.tree.column("Forma", width=140)
            self.tree.column("Concentración", width=140)
            self.tree.column("Controlado", width=80)
        elif ancho < 1300:
            self.tree.column("ID", width=60)
            self.tree.column("Nombre", width=250)
            self.tree.column("Forma", width=180)
            self.tree.column("Concentración", width=180)
            self.tree.column("Controlado", width=100)
        else:
            self.tree.column("ID", width=70)
            self.tree.column("Nombre", width=350)
            self.tree.column("Forma", width=220)
            self.tree.column("Concentración", width=220)
            self.tree.column("Controlado", width=120)

    def cargar_datos_iniciales(self):
        self.crear_interfaz()
        self.cargar_todos()
        self.cargar_filtros()
        self.ajustar_tabla()

    def toggle_fullscreen(self, event=None):
        fullscreen = self.attributes('-fullscreen')
        self.attributes('-fullscreen', not fullscreen)
        if fullscreen:
            # Restaurar tamaño responsivo
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            window_width = int(screen_width * 0.9)
            window_height = int(screen_height * 0.85)
            self.geometry(f"{window_width}x{window_height}")
            self.update_idletasks()
        self.after(100, self.ajustar_tabla)

    def crear_interfaz(self):
        self.debug_print("Creando interfaz...")

        self.frame_principal = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.frame_principal.pack(fill="both", expand=True)

        self.crear_header()

        content = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)

        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)

        # Panel izquierdo (tabla y filtros) ocupa todo el ancho
        self.panel_izquierdo = ctk.CTkFrame(content, fg_color=AppTheme.SURFACE, corner_radius=15)
        self.panel_izquierdo.grid(row=0, column=0, sticky="nsew")

        self.crear_panel_tabla()
        self.debug_print("Interfaz creada")

    def crear_header(self):
        header = ctk.CTkFrame(self.frame_principal, fg_color=AppTheme.PRIMARY, height=70, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        titulo = ctk.CTkLabel(header, text="📋 Catálogo de Medicamentos",
                              font=("Segoe UI", 20, "bold"), text_color=AppTheme.TEXT_ON_PRIMARY)
        titulo.pack(side="left", padx=25)

        frame_botones = ctk.CTkFrame(header, fg_color="transparent")
        frame_botones.pack(side="right", padx=25)

        btn_fullscreen = ctk.CTkButton(frame_botones, text="⛶", width=40, height=35,
                                       font=("Segoe UI", 14), fg_color="transparent",
                                       hover_color=AppTheme.PRIMARY_LIGHT, text_color=AppTheme.TEXT_ON_PRIMARY,
                                       corner_radius=8, command=self.toggle_fullscreen)
        btn_fullscreen.pack(side="left", padx=5)

        btn_volver = ctk.CTkButton(frame_botones, text="← Áreas", width=80, height=35,
                                   font=("Segoe UI", 11), fg_color="transparent",
                                   hover_color=AppTheme.PRIMARY_LIGHT, text_color=AppTheme.TEXT_ON_PRIMARY,
                                   corner_radius=8, command=self.volver_areas)
        btn_volver.pack(side="left", padx=5)

        btn_cerrar = ctk.CTkButton(frame_botones, text="🚪 Salir", width=80, height=35,
                                   font=("Segoe UI", 11), fg_color=AppTheme.PRIMARY_LIGHT,
                                   hover_color=AppTheme.PRIMARY_DARK, text_color=AppTheme.TEXT_ON_PRIMARY,
                                   corner_radius=8, command=self.cerrar_sesion)
        btn_cerrar.pack(side="left", padx=5)

    def volver_areas(self):
        try:
            from views.area_select_view import AreaSelectView
            area_view = AreaSelectView(self.usuario)
            self.destroy()
            area_view.mainloop()
        except Exception as e:
            print(f"Error al volver: {e}")
            messagebox.showerror("Error", f"No se pudo volver: {e}")

    def cerrar_sesion(self):
        try:
            from views.login_view import LoginView
            login = LoginView()
            self.destroy()
            login.mainloop()
        except Exception as e:
            print(f"Error al cerrar sesión: {e}")

    def crear_panel_tabla(self):
        # Frame de botones de acción (Nuevo, Editar, Eliminar) estilo CRUD
        frame_acciones = ctk.CTkFrame(self.panel_izquierdo, fg_color="transparent")
        frame_acciones.pack(fill="x", padx=20, pady=(20, 10))

        btn_nuevo = ctk.CTkButton(frame_acciones, text="➕ NUEVO", width=120, height=40,
                                  font=("Segoe UI", 12, "bold"), fg_color=AppTheme.PRIMARY,
                                  hover_color=AppTheme.PRIMARY_DARK, corner_radius=10,
                                  command=self.abrir_formulario_medicamento)
        btn_nuevo.pack(side="left", padx=5)

        btn_editar = ctk.CTkButton(frame_acciones, text="✏️ EDITAR", width=120, height=40,
                                   font=("Segoe UI", 12, "bold"), fg_color="#4a5568",
                                   hover_color="#2d3748", corner_radius=10,
                                   command=self.editar_seleccionado)
        btn_editar.pack(side="left", padx=5)

        btn_eliminar = ctk.CTkButton(frame_acciones, text="🗑 ELIMINAR", width=120, height=40,
                                     font=("Segoe UI", 12, "bold"), fg_color="#e53e3e",
                                     hover_color="#c53030", corner_radius=10,
                                     command=self.eliminar_seleccionado)
        btn_eliminar.pack(side="left", padx=5)

        # Filtros
        frame_filtros = ctk.CTkFrame(self.panel_izquierdo, fg_color="transparent")
        frame_filtros.pack(fill="x", padx=20, pady=(0, 15))

        # Fila búsqueda
        self.busqueda_entry = ctk.CTkEntry(frame_filtros, placeholder_text="🔍 Buscar por nombre...",
                                           height=38, font=("Segoe UI", 12), corner_radius=10)
        self.busqueda_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.busqueda_entry.bind("<KeyRelease>", self.aplicar_filtros)

        btn_limpiar = ctk.CTkButton(frame_filtros, text="🔄 Limpiar", width=100, height=38,
                                    font=("Segoe UI", 11), fg_color="#4a5568", hover_color="#2d3748",
                                    corner_radius=10, command=self.limpiar_filtros)
        btn_limpiar.pack(side="left")

        # Segunda fila de filtros
        frame_filtros2 = ctk.CTkFrame(self.panel_izquierdo, fg_color="transparent")
        frame_filtros2.pack(fill="x", padx=20, pady=(0, 15))

        ctk.CTkLabel(frame_filtros2, text="Tipo:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(0, 5))
        self.combo_controlado = ctk.CTkComboBox(frame_filtros2, values=["Todos", "Controlados", "No controlados"],
                                                width=120, height=35, font=("Segoe UI", 11),
                                                corner_radius=10, state="readonly", command=self.aplicar_filtros)
        self.combo_controlado.pack(side="left", padx=5)
        self.combo_controlado.set("Todos")

        ctk.CTkLabel(frame_filtros2, text="Forma:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(15, 5))
        self.combo_forma = ctk.CTkComboBox(frame_filtros2, values=["Todos"], width=150, height=35,
                                           font=("Segoe UI", 11), corner_radius=10, state="readonly",
                                           command=self.aplicar_filtros)
        self.combo_forma.pack(side="left", padx=5)

        ctk.CTkLabel(frame_filtros2, text="Concentración:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(15, 5))
        self.txt_concentracion_filtro = ctk.CTkEntry(frame_filtros2, width=200, height=35,
                                                     font=("Segoe UI", 11), corner_radius=10,
                                                     placeholder_text="Escribe para filtrar...")
        self.txt_concentracion_filtro.pack(side="left", padx=5)
        self.txt_concentracion_filtro.bind("<KeyRelease>", self.aplicar_filtros)

        # Tabla
        frame_tabla = ctk.CTkFrame(self.panel_izquierdo)
        frame_tabla.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Catalogo.Treeview", font=("Segoe UI", 10), rowheight=34,
                        background="white", fieldbackground="white", foreground=AppTheme.TEXT_PRIMARY)
        style.configure("Catalogo.Treeview.Heading", font=("Segoe UI", 11, "bold"),
                        background=AppTheme.PRIMARY, foreground=AppTheme.TEXT_ON_PRIMARY, relief="flat")

        scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical")
        scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal")

        self.tree = ttk.Treeview(frame_tabla, columns=("ID", "Nombre", "Forma", "Concentración", "Controlado"),
                                 show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set,
                                 style="Catalogo.Treeview", height=20)

        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)

        self.tree.heading("ID", text="ID")
        self.tree.heading("Nombre", text="Nombre")
        self.tree.heading("Forma", text="Forma Farmacéutica")
        self.tree.heading("Concentración", text="Concentración")
        self.tree.heading("Controlado", text="Controlado")

        self.tree.column("ID", width=60, anchor="center")
        self.tree.column("Nombre", width=300)
        self.tree.column("Forma", width=200)
        self.tree.column("Concentración", width=200)
        self.tree.column("Controlado", width=100, anchor="center")

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.on_doble_click)

    def on_doble_click(self, event):
        self.editar_seleccionado()

    def editar_seleccionado(self):
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un medicamento para editar")
            return
        valores = self.tree.item(seleccion[0])['values']
        self.abrir_formulario_medicamento(medicamento_id=valores[0])

    def eliminar_seleccionado(self):
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un medicamento para eliminar")
            return
        valores = self.tree.item(seleccion[0])['values']
        medicamento_id = valores[0]
        nombre = valores[1]
        if messagebox.askyesno("Confirmar", f"¿Eliminar el medicamento '{nombre}'?\nEsta acción no se puede deshacer."):
            try:
                success = Database.execute_non_query("DELETE FROM catalogo_medicamentos WHERE id_catalogo = ?",
                                                     (medicamento_id,))
                if success:
                    messagebox.showinfo("Éxito", "Medicamento eliminado correctamente")
                    self.cargar_todos()
                    self.cargar_filtros()
                else:
                    messagebox.showerror("Error", "No se pudo eliminar el medicamento")
            except Exception as e:
                messagebox.showerror("Error", f"Error al eliminar: {e}")

    # ---------- FORMULARIO EMERGENTE (CRUD) ----------
    def abrir_formulario_medicamento(self, medicamento_id=None):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Nuevo Medicamento" if not medicamento_id else "Editar Medicamento")
        dialog.geometry("550x600")
        dialog.resizable(False, False)
        dialog.grab_set()

        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 275
        y = self.winfo_y() + (self.winfo_height() // 2) - 300
        dialog.geometry(f"+{x}+{y}")

        scroll = ctk.CTkScrollableFrame(dialog, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=20)

        titulo = ctk.CTkLabel(scroll, text="AGREGAR MEDICAMENTO" if not medicamento_id else "EDITAR MEDICAMENTO",
                              font=("Segoe UI", 18, "bold"), text_color=AppTheme.PRIMARY)
        titulo.pack(pady=(0, 15))

        ctk.CTkLabel(scroll, text="Nombre del medicamento *", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(5, 2))
        txt_nombre = ctk.CTkEntry(scroll, height=40, font=("Segoe UI", 11), corner_radius=10)
        txt_nombre.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(scroll, text="Forma farmacéutica", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(5, 2))
        txt_forma = ctk.CTkEntry(scroll, placeholder_text="Ej: Tableta, Jarabe, Cápsula...", height=40,
                                 font=("Segoe UI", 11), corner_radius=10)
        txt_forma.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(scroll, text="Concentración", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(5, 2))
        txt_concentracion = ctk.CTkEntry(scroll, placeholder_text="Ej: 500 mg, 10 mg/ml...", height=40,
                                         font=("Segoe UI", 11), corner_radius=10)
        txt_concentracion.pack(fill="x", pady=(0, 10))

        controlado_var = ctk.BooleanVar(value=False)
        check_controlado = ctk.CTkCheckBox(scroll, text="⚠️ Medicamento controlado (requiere receta especial)",
                                           variable=controlado_var, font=("Segoe UI", 12),
                                           text_color=AppTheme.TEXT_PRIMARY)
        check_controlado.pack(anchor="w", pady=10)

        # Si es edición, cargar datos
        if medicamento_id:
            query = "SELECT id_catalogo, nombre, forma_farmaceutica, concentracion, es_controlado FROM catalogo_medicamentos WHERE id_catalogo = ?"
            rows = Database.execute_query(query, (medicamento_id,))
            if rows:
                row = rows[0]
                txt_nombre.insert(0, row[1])
                if row[2]:
                    txt_forma.insert(0, row[2])
                if row[3]:
                    txt_concentracion.insert(0, row[3])
                controlado_var.set(row[4])

        def guardar():
            nombre = txt_nombre.get().strip()
            forma = txt_forma.get().strip() or None
            concentracion = txt_concentracion.get().strip() or None
            es_controlado = controlado_var.get()

            if not nombre:
                messagebox.showerror("Error", "El nombre del medicamento es obligatorio")
                return

            if medicamento_id:
                # Actualizar
                query = """UPDATE catalogo_medicamentos 
                           SET nombre = ?, forma_farmaceutica = ?, concentracion = ?, es_controlado = ?
                           WHERE id_catalogo = ?"""
                params = (nombre, forma, concentracion, es_controlado, medicamento_id)
                success = Database.execute_non_query(query, params)
                msg = "actualizado"
            else:
                # Insertar
                query = """INSERT INTO catalogo_medicamentos (nombre, forma_farmaceutica, concentracion, es_controlado)
                           VALUES (?, ?, ?, ?)"""
                params = (nombre, forma, concentracion, es_controlado)
                success = Database.execute_non_query(query, params)
                msg = "agregado"

            if success:
                messagebox.showinfo("Éxito", f"Medicamento {msg} correctamente")
                self.cargar_todos()
                self.cargar_filtros()
                dialog.destroy()
            else:
                messagebox.showerror("Error", f"No se pudo {msg} el medicamento")

        frame_botones = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_botones.pack(fill="x", pady=15)

        btn_guardar = ctk.CTkButton(frame_botones, text="💾 GUARDAR", height=45,
                                    font=("Segoe UI", 13, "bold"), fg_color=AppTheme.PRIMARY,
                                    hover_color=AppTheme.PRIMARY_DARK, corner_radius=10,
                                    command=guardar)
        btn_guardar.pack(fill="x", pady=2)

        btn_cancelar = ctk.CTkButton(frame_botones, text="❌ CANCELAR", height=40,
                                     font=("Segoe UI", 12), fg_color="#4a5568",
                                     hover_color="#2d3748", text_color=AppTheme.TEXT_ON_PRIMARY,
                                     corner_radius=10, command=dialog.destroy)
        btn_cancelar.pack(fill="x", pady=2)

    # ---------- MÉTODOS DE CARGA Y FILTRADO ----------
    def cargar_filtros(self):
        try:
            rows = Database.execute_query(
                "SELECT DISTINCT forma_farmaceutica FROM catalogo_medicamentos WHERE forma_farmaceutica IS NOT NULL ORDER BY forma_farmaceutica")
            formas = ["Todos"] + [row[0] for row in rows] if rows else ["Todos"]
            self.combo_forma.configure(values=formas)
            self.combo_forma.set("Todos")
        except Exception as e:
            self.debug_print(f"Error al cargar filtros: {e}")

    def aplicar_filtros(self, event=None):
        busqueda = self.busqueda_entry.get().strip()
        tipo_controlado = self.combo_controlado.get()
        forma = self.combo_forma.get()
        concentracion = self.txt_concentracion_filtro.get().strip()

        query = """
            SELECT id_catalogo, nombre, forma_farmaceutica, concentracion, es_controlado
            FROM catalogo_medicamentos
            WHERE 1=1
        """
        params = []

        if busqueda:
            query += " AND (nombre LIKE ? OR forma_farmaceutica LIKE ? OR concentracion LIKE ?)"
            params.extend([f"%{busqueda}%", f"%{busqueda}%", f"%{busqueda}%"])

        if tipo_controlado == "Controlados":
            query += " AND es_controlado = 1"
        elif tipo_controlado == "No controlados":
            query += " AND es_controlado = 0"

        if forma != "Todos":
            query += " AND forma_farmaceutica = ?"
            params.append(forma)

        if concentracion:
            query += " AND concentracion LIKE ?"
            params.append(f"%{concentracion}%")

        query += " ORDER BY id_catalogo ASC"
        self.actualizar_treeview(query, params if params else None)

    def limpiar_filtros(self):
        self.busqueda_entry.delete(0, "end")
        self.combo_controlado.set("Todos")
        self.combo_forma.set("Todos")
        self.txt_concentracion_filtro.delete(0, "end")
        self.cargar_todos()

    def cargar_todos(self):
        self.actualizar_treeview(
            "SELECT id_catalogo, nombre, forma_farmaceutica, concentracion, es_controlado FROM catalogo_medicamentos ORDER BY id_catalogo ASC")

    def actualizar_treeview(self, query, params=None):
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            rows = Database.execute_query(query, params) if params else Database.execute_query(query)
            if rows is None:
                return

            for row in rows:
                controlado = "✓ Sí" if row[4] else "✗ No"
                tags = ('controlado',) if row[4] else ()
                self.tree.insert("", "end", values=(row[0], row[1], row[2], row[3], controlado), tags=tags)

            self.tree.tag_configure('controlado', background='#E1BEE7', foreground='#4A0072')
            self.ajustar_tabla()
        except Exception as e:
            self.debug_print(f"Error al actualizar treeview: {e}")
            messagebox.showerror("Error", f"No se pudieron cargar los datos:\n{e}")

    def on_closing(self):
        self.quit()
        self.destroy()