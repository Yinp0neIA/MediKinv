import customtkinter as ctk
from tkinter import messagebox, ttk
import tkinter as tk

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
        self.geometry("1400x800")
        self.configure(fg_color=AppTheme.SECONDARY)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.bind('<Escape>', self.toggle_fullscreen)

        self.after(100, self.maximizar_ventana)

    def debug_print(self, msg):
        if self._debug:
            print(f"[DEBUG CATALOGO] {msg}")

    def maximizar_ventana(self):
        try:
            self.attributes('-fullscreen', True)
            self.debug_print("Ventana en pantalla completa")
        except Exception as e:
            self.debug_print(f"No se pudo activar pantalla completa: {e}")

        self.crear_interfaz()
        self.cargar_todos()
        self.cargar_filtros()

    def toggle_fullscreen(self, event=None):
        fullscreen = self.attributes('-fullscreen')
        self.attributes('-fullscreen', not fullscreen)
        if fullscreen:
            self.geometry("1400x800")
            self.update_idletasks()
            self.debug_print("Modo ventana normal")
        else:
            self.debug_print("Modo pantalla completa")

    def crear_interfaz(self):
        self.debug_print("Creando interfaz...")

        self.frame_principal = ctk.CTkFrame(self, fg_color="transparent", corner_radius=0)
        self.frame_principal.pack(fill="both", expand=True)

        self.crear_header()

        content = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)

        content.grid_columnconfigure(0, weight=2)
        content.grid_columnconfigure(1, weight=1)
        content.grid_rowconfigure(0, weight=1)

        self.panel_izquierdo = ctk.CTkFrame(content, fg_color=AppTheme.SURFACE, corner_radius=15)
        self.panel_izquierdo.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.panel_derecho = ctk.CTkFrame(content, fg_color=AppTheme.SURFACE, corner_radius=15)
        self.panel_derecho.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self.crear_panel_tabla()
        self.crear_panel_formulario()

        self.debug_print("Interfaz creada")

    def crear_header(self):
        header = ctk.CTkFrame(self.frame_principal, fg_color=AppTheme.PRIMARY, height=80, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        titulo = ctk.CTkLabel(header, text="📋 Catálogo de Medicamentos",
                              font=("Segoe UI", 24, "bold"), text_color=AppTheme.TEXT_ON_PRIMARY)
        titulo.pack(side="left", padx=30, pady=20)

        frame_botones = ctk.CTkFrame(header, fg_color="transparent")
        frame_botones.pack(side="right", padx=30, pady=20)

        btn_fullscreen = ctk.CTkButton(frame_botones, text="⛶ Pantalla Completa", width=150, height=38,
                                       font=("Segoe UI", 11), fg_color="transparent",
                                       hover_color=AppTheme.PRIMARY_LIGHT, text_color=AppTheme.TEXT_ON_PRIMARY,
                                       corner_radius=12, command=self.toggle_fullscreen)
        btn_fullscreen.pack(side="left", padx=10)

        btn_volver = ctk.CTkButton(frame_botones, text="← Volver a Áreas", width=130, height=38,
                                   font=("Segoe UI", 11), fg_color="transparent",
                                   hover_color=AppTheme.PRIMARY_LIGHT, text_color=AppTheme.TEXT_ON_PRIMARY,
                                   corner_radius=12, command=self.volver_areas)
        btn_volver.pack(side="left", padx=10)

        btn_cerrar = ctk.CTkButton(frame_botones, text="Cerrar Sesión", width=120, height=38,
                                   font=("Segoe UI", 11), fg_color=AppTheme.PRIMARY_LIGHT,
                                   hover_color=AppTheme.PRIMARY_DARK, text_color=AppTheme.TEXT_ON_PRIMARY,
                                   corner_radius=12, command=self.cerrar_sesion)
        btn_cerrar.pack(side="left", padx=10)

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

    def cargar_filtros(self):
        try:
            rows = Database.execute_query(
                "SELECT DISTINCT forma_farmaceutica FROM catalogo_medicamentos WHERE forma_farmaceutica IS NOT NULL ORDER BY forma_farmaceutica")
            formas = ["Todos"] + [row[0] for row in rows] if rows else ["Todos"]
            self.combo_forma.configure(values=formas)
            self.combo_forma.set("Todos")

            rows = Database.execute_query(
                "SELECT DISTINCT concentracion FROM catalogo_medicamentos WHERE concentracion IS NOT NULL ORDER BY concentracion")
            self.concentraciones_lista = [row[0] for row in rows] if rows else []

            if hasattr(self, 'txt_concentracion_filtro'):
                pass
        except Exception as e:
            self.debug_print(f"Error al cargar filtros: {e}")

    def crear_panel_tabla(self):
        frame_controles = ctk.CTkFrame(self.panel_izquierdo, fg_color="transparent")
        frame_controles.pack(fill="x", padx=20, pady=(20, 10))

        titulo = ctk.CTkLabel(frame_controles, text="📋 Lista de Medicamentos",
                              font=("Segoe UI", 20, "bold"), text_color=AppTheme.PRIMARY)
        titulo.pack(anchor="w", pady=(0, 15))

        frame_busqueda = ctk.CTkFrame(frame_controles, fg_color="transparent")
        frame_busqueda.pack(fill="x", pady=(0, 10))

        self.busqueda_entry = ctk.CTkEntry(frame_busqueda, placeholder_text="🔍 Buscar por nombre...",
                                           height=40, font=("Segoe UI", 12), corner_radius=12,
                                           border_width=1, border_color=AppTheme.PRIMARY_LIGHT)
        self.busqueda_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.busqueda_entry.bind("<KeyRelease>", self.aplicar_filtros)

        btn_limpiar = ctk.CTkButton(frame_busqueda, text="🔄 Limpiar", width=100, height=40,
                                    font=("Segoe UI", 12), fg_color="#4a5568", hover_color="#2d3748",
                                    corner_radius=12, command=self.limpiar_filtros)
        btn_limpiar.pack(side="left")

        frame_filtros = ctk.CTkFrame(frame_controles, fg_color="transparent")
        frame_filtros.pack(fill="x", pady=(10, 0))

        ctk.CTkLabel(frame_filtros, text="Tipo:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(0, 5))
        self.combo_controlado = ctk.CTkComboBox(frame_filtros, values=["Todos", "Controlados", "No controlados"],
                                                width=120, height=35, font=("Segoe UI", 11),
                                                corner_radius=12, state="readonly", command=self.aplicar_filtros)
        self.combo_controlado.pack(side="left", padx=5)
        self.combo_controlado.set("Todos")

        ctk.CTkLabel(frame_filtros, text="Forma:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(15, 5))
        self.combo_forma = ctk.CTkComboBox(frame_filtros, values=["Todos"], width=150, height=35,
                                           font=("Segoe UI", 11), corner_radius=12, state="readonly",
                                           command=self.aplicar_filtros)
        self.combo_forma.pack(side="left", padx=5)

        ctk.CTkLabel(frame_filtros, text="Concentración:", font=("Segoe UI", 11, "bold")).pack(side="left", padx=(15, 5))
        self.txt_concentracion_filtro = ctk.CTkEntry(frame_filtros, width=200, height=35,
                                                     font=("Segoe UI", 11), corner_radius=12,
                                                     placeholder_text="Escribe para filtrar concentración...")
        self.txt_concentracion_filtro.pack(side="left", padx=5)
        self.txt_concentracion_filtro.bind("<KeyRelease>", self.on_concentracion_change_filtro)

        frame_tabla = ctk.CTkFrame(self.panel_izquierdo)
        frame_tabla.pack(fill="both", expand=True, padx=20, pady=(10, 20))

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

        self.tree.bind("<<TreeviewSelect>>", self.on_medicamento_seleccionado)

    def on_concentracion_change_filtro(self, event):
        texto = self.txt_concentracion_filtro.get().strip()
        if texto:
            self.cargar_sugerencias_concentracion(texto)
        self.aplicar_filtros()

    def cargar_sugerencias_concentracion(self, texto):
        try:
            query = """
                SELECT DISTINCT TOP 10 concentracion 
                FROM catalogo_medicamentos 
                WHERE concentracion LIKE ? 
                ORDER BY concentracion
            """
            rows = Database.execute_query(query, (f"%{texto}%",))
            if rows:
                self.debug_print(f"Sugerencias encontradas: {len(rows)}")
        except Exception as e:
            self.debug_print(f"Error en sugerencias: {e}")

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
        except Exception as e:
            self.debug_print(f"Error al actualizar treeview: {e}")
            messagebox.showerror("Error", f"No se pudieron cargar los datos:\n{e}")

    def crear_panel_formulario(self):
        scroll = ctk.CTkScrollableFrame(self.panel_derecho, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=20)

        self.form_titulo = ctk.CTkLabel(scroll, text="➕ AGREGAR MEDICAMENTO",
                                        font=("Segoe UI", 20, "bold"), text_color=AppTheme.PRIMARY)
        self.form_titulo.pack(pady=(0, 20))

        ctk.CTkLabel(scroll, text="Nombre del medicamento *", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_nombre = ctk.CTkEntry(scroll, height=45, font=("Segoe UI", 12), corner_radius=12)
        self.txt_nombre.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Forma farmacéutica", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_forma = ctk.CTkEntry(scroll, placeholder_text="Ej: Tableta, Jarabe, Cápsula...", height=45,
                                      font=("Segoe UI", 12), corner_radius=12)
        self.txt_forma.pack(fill="x", pady=(0, 5))

        ctk.CTkLabel(scroll, text="Concentración", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(10, 2))
        self.txt_concentracion = ctk.CTkEntry(scroll, placeholder_text="Ej: 500 mg, 10 mg/ml...", height=45,
                                              font=("Segoe UI", 12), corner_radius=12)
        self.txt_concentracion.pack(fill="x", pady=(0, 5))

        self.controlado_var = ctk.BooleanVar(value=False)
        self.check_controlado = ctk.CTkCheckBox(scroll, text="⚠️ Medicamento controlado (requiere receta especial)",
                                                variable=self.controlado_var, font=("Segoe UI", 12),
                                                text_color=AppTheme.TEXT_PRIMARY)
        self.check_controlado.pack(anchor="w", pady=15)

        ctk.CTkFrame(scroll, height=2, fg_color=AppTheme.PRIMARY).pack(fill="x", pady=15)

        frame_botones = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_botones.pack(fill="x", pady=10)

        self.btn_guardar = ctk.CTkButton(frame_botones, text="💾 GUARDAR", height=50,
                                         font=("Segoe UI", 14, "bold"), fg_color=AppTheme.PRIMARY,
                                         hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                         command=self.guardar_medicamento)
        self.btn_guardar.pack(fill="x", pady=5)

        self.btn_actualizar = ctk.CTkButton(frame_botones, text="🔄 ACTUALIZAR", height=50,
                                            font=("Segoe UI", 14, "bold"), fg_color=AppTheme.PRIMARY,
                                            hover_color=AppTheme.PRIMARY_DARK, corner_radius=12,
                                            command=self.actualizar_medicamento, state="disabled")
        self.btn_actualizar.pack(fill="x", pady=5)

        self.btn_eliminar = ctk.CTkButton(frame_botones, text="🗑️ ELIMINAR", height=50,
                                          font=("Segoe UI", 14, "bold"), fg_color="#e53e3e",
                                          hover_color="#c53030", corner_radius=12,
                                          command=self.eliminar_medicamento, state="disabled")
        self.btn_eliminar.pack(fill="x", pady=5)

        self.btn_cancelar = ctk.CTkButton(frame_botones, text="CANCELAR", height=45,
                                          font=("Segoe UI", 13), fg_color="#4a5568",
                                          hover_color="#2d3748", corner_radius=12,
                                          command=self.limpiar_formulario)
        self.btn_cancelar.pack(fill="x", pady=5)

        lbl_info = ctk.CTkLabel(scroll,
                                text="* Campos obligatorios\n\nSeleccione un medicamento de la tabla para editarlo o eliminarlo\n\nLos medicamentos controlados aparecen en morado\n\nPresione ESC para salir de pantalla completa\n\n💡 En el filtro de concentración puedes escribir directamente",
                                font=("Segoe UI", 11), text_color=AppTheme.TEXT_SECONDARY, justify="left")
        lbl_info.pack(pady=(20, 0))

    def on_medicamento_seleccionado(self, event):
        seleccion = self.tree.selection()
        if seleccion:
            valores = self.tree.item(seleccion[0])['values']
            id_medicamento = valores[0]
            query = "SELECT id_catalogo, nombre, forma_farmaceutica, concentracion, es_controlado FROM catalogo_medicamentos WHERE id_catalogo = ?"
            rows = Database.execute_query(query, (id_medicamento,))

            if rows and len(rows) > 0:
                self.medicamento_editando = {
                    'id': rows[0][0],
                    'nombre': rows[0][1],
                    'forma': rows[0][2],
                    'concentracion': rows[0][3],
                    'controlado': rows[0][4]
                }

                self.txt_nombre.delete(0, "end")
                self.txt_nombre.insert(0, self.medicamento_editando['nombre'])

                self.txt_forma.delete(0, "end")
                if self.medicamento_editando['forma']:
                    self.txt_forma.insert(0, self.medicamento_editando['forma'])

                self.txt_concentracion.delete(0, "end")
                if self.medicamento_editando['concentracion']:
                    self.txt_concentracion.insert(0, self.medicamento_editando['concentracion'])

                self.controlado_var.set(self.medicamento_editando['controlado'])

                self.form_titulo.configure(text="✏️ EDITAR MEDICAMENTO")
                self.btn_guardar.configure(state="disabled")
                self.btn_actualizar.configure(state="normal")
                self.btn_eliminar.configure(state="normal")

    def guardar_medicamento(self):
        nombre = self.txt_nombre.get().strip()
        forma = self.txt_forma.get().strip() or None
        concentracion = self.txt_concentracion.get().strip() or None
        es_controlado = self.controlado_var.get()

        if not nombre:
            messagebox.showwarning("Advertencia", "El nombre del medicamento es obligatorio")
            return

        try:
            query = """
                INSERT INTO catalogo_medicamentos (nombre, forma_farmaceutica, concentracion, es_controlado)
                VALUES (?, ?, ?, ?)
            """
            params = (nombre, forma, concentracion, es_controlado)
            success = Database.execute_non_query(query, params)

            if success:
                messagebox.showinfo("Éxito", "Medicamento agregado correctamente")
                self.limpiar_formulario()
                self.cargar_todos()
                self.cargar_filtros()
            else:
                messagebox.showerror("Error", "No se pudo agregar el medicamento")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{e}")

    def actualizar_medicamento(self):
        if not self.medicamento_editando:
            messagebox.showwarning("Advertencia", "No hay medicamento seleccionado para actualizar")
            return

        nombre = self.txt_nombre.get().strip()
        forma = self.txt_forma.get().strip() or None
        concentracion = self.txt_concentracion.get().strip() or None
        es_controlado = self.controlado_var.get()

        if not nombre:
            messagebox.showwarning("Advertencia", "El nombre del medicamento es obligatorio")
            return

        try:
            query = """
                UPDATE catalogo_medicamentos
                SET nombre = ?, forma_farmaceutica = ?, concentracion = ?, es_controlado = ?
                WHERE id_catalogo = ?
            """
            params = (nombre, forma, concentracion, es_controlado, self.medicamento_editando['id'])
            success = Database.execute_non_query(query, params)

            if success:
                messagebox.showinfo("Éxito", "Medicamento actualizado correctamente")
                self.limpiar_formulario()
                self.cargar_todos()
                self.cargar_filtros()
            else:
                messagebox.showerror("Error", "No se pudo actualizar el medicamento")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo actualizar:\n{e}")

    def eliminar_medicamento(self):
        if not self.medicamento_editando:
            messagebox.showwarning("Advertencia", "No hay medicamento seleccionado para eliminar")
            return

        try:
            success = Database.execute_non_query("DELETE FROM catalogo_medicamentos WHERE id_catalogo = ?",
                                                 (self.medicamento_editando['id'],))
            if success:
                messagebox.showinfo("Éxito", "Medicamento eliminado correctamente")
                self.limpiar_formulario()
                self.cargar_todos()
                self.cargar_filtros()
            else:
                messagebox.showerror("Error", "No se pudo eliminar el medicamento")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo eliminar:\n{e}")

    def limpiar_formulario(self):
        self.medicamento_editando = None
        self.form_titulo.configure(text="➕ AGREGAR MEDICAMENTO")
        self.txt_nombre.delete(0, "end")
        self.txt_forma.delete(0, "end")
        self.txt_concentracion.delete(0, "end")
        self.controlado_var.set(False)
        self.btn_guardar.configure(state="normal")
        self.btn_actualizar.configure(state="disabled")
        self.btn_eliminar.configure(state="disabled")

        for item in self.tree.selection():
            self.tree.selection_remove(item)

    def on_closing(self):
        self.quit()
        self.destroy()