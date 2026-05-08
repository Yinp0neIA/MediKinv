from config.database import Database

class Medicamento:
    """Modelo único de Medicamento para ambas áreas"""

    @staticmethod
    def obtener_catalogo(busqueda=""):
        """Obtiene todos los medicamentos del catálogo (para combos y búsqueda)"""
        query = """
            SELECT id_catalogo, nombre, forma_farmaceutica, concentracion,
                   CONCAT(ISNULL(forma_farmaceutica, ''), ' ', ISNULL(concentracion, '')) as presentacion
            FROM catalogo_medicamentos
            WHERE 1=1
        """
        params = []
        if busqueda:
            query += " AND nombre LIKE ?"
            params.append(f"%{busqueda}%")
        query += " ORDER BY nombre"
        result = Database.execute_query(query, tuple(params) if params else None)
        catalogo = []
        if result:
            for row in result:
                catalogo.append({
                    'id_catalogo': row[0],
                    'nombre': row[1],
                    'forma_farmaceutica': row[2] or "",
                    'concentracion': row[3] or "",
                    'presentacion': row[4].strip() or "---"
                })
        return catalogo

    @staticmethod
    def obtener_por_area(area_id, busqueda=""):
        """Obtiene los medicamentos del inventario para un área.
           Si se proporciona busqueda, busca en nombre del medicamento O en observaciones."""
        tabla_inventario = "inventario_farmacia" if area_id == 1 else "inventario_odontologia"
        query = f"""
            SELECT i.id_inventario as id_medicamento, 
                   c.nombre, 
                   CONCAT(ISNULL(c.forma_farmaceutica, ''), ' ', ISNULL(c.concentracion, '')) as presentacion,
                   i.ubicacion, i.stock, i.caducidad, i.es_donacion, i.donante, i.observaciones
            FROM {tabla_inventario} i
            JOIN catalogo_medicamentos c ON i.id_catalogo = c.id_catalogo
            WHERE 1=1
        """
        params = []
        if busqueda:
            query += " AND (c.nombre LIKE ? OR i.observaciones LIKE ?)"
            params.append(f"%{busqueda}%")
            params.append(f"%{busqueda}%")
        query += " ORDER BY c.nombre"
        result = Database.execute_query(query, tuple(params) if params else None)
        medicamentos = []
        if result:
            for row in result:
                presentacion = row[2].strip() if row[2] else "---"
                medicamentos.append({
                    'id_medicamento': row[0],
                    'nombre': row[1],
                    'presentacion': presentacion,
                    'ubicacion': row[3] or "",
                    'stock': row[4],
                    'caducidad': row[5],
                    'es_donacion': row[6],
                    'donante': row[7] or "",
                    'observaciones': row[8] or ""
                })
        return medicamentos

    @staticmethod
    def obtener_por_id(medicamento_id):
        # Buscar en farmacia
        query_farma = """
            SELECT i.id_inventario, i.id_catalogo, i.ubicacion, i.caducidad,
                   i.stock, i.es_donacion, i.donante, i.observaciones,
                   c.nombre, c.forma_farmaceutica, c.concentracion
            FROM inventario_farmacia i
            JOIN catalogo_medicamentos c ON i.id_catalogo = c.id_catalogo
            WHERE i.id_inventario = ?
        """
        result = Database.execute_query(query_farma, (medicamento_id,))
        if result:
            row = result[0]
            forma = row[9] or ""
            conc = row[10] or ""
            presentacion = f"{forma} {conc}".strip()
            if not presentacion:
                presentacion = "---"
            return {
                'id_medicamento': row[0],
                'id_catalogo': row[1],
                'ubicacion': row[2] or "",
                'caducidad': row[3],
                'stock': row[4],
                'es_donacion': row[5],
                'donante': row[6] or "",
                'observaciones': row[7] or "",
                'nombre': row[8],
                'presentacion': presentacion,
                'forma_farmaceutica': forma,
                'concentracion': conc,
                'area_id': 1
            }
        # Buscar en odontología
        query_odonto = """
            SELECT i.id_inventario, i.id_catalogo, i.ubicacion, i.caducidad,
                   i.stock, i.es_donacion, i.donante, i.observaciones,
                   c.nombre, c.forma_farmaceutica, c.concentracion
            FROM inventario_odontologia i
            JOIN catalogo_medicamentos c ON i.id_catalogo = c.id_catalogo
            WHERE i.id_inventario = ?
        """
        result = Database.execute_query(query_odonto, (medicamento_id,))
        if result:
            row = result[0]
            forma = row[9] or ""
            conc = row[10] or ""
            presentacion = f"{forma} {conc}".strip()
            if not presentacion:
                presentacion = "---"
            return {
                'id_medicamento': row[0],
                'id_catalogo': row[1],
                'ubicacion': row[2] or "",
                'caducidad': row[3],
                'stock': row[4],
                'es_donacion': row[5],
                'donante': row[6] or "",
                'observaciones': row[7] or "",
                'nombre': row[8],
                'presentacion': presentacion,
                'forma_farmaceutica': forma,
                'concentracion': conc,
                'area_id': 2
            }
        return None

    @staticmethod
    def _separar_presentacion(presentacion_str):
        """Separa la presentación en forma farmacéutica y concentración.
           Toma la primera palabra como forma y el resto como concentración."""
        partes = presentacion_str.strip().split(maxsplit=1)
        if len(partes) == 1:
            return partes[0], ""
        else:
            return partes[0], partes[1]

    @staticmethod
    def crear(area_id, nombre, presentacion, ubicacion, caducidad,
              observaciones, stock, es_donacion, donante=""):
        """Crea un nuevo medicamento en el inventario y, si no existe en el catálogo, lo crea.
           (Este método se conserva pero no se usará en el flujo principal de selección desde catálogo)"""
        forma_farmaceutica, concentracion = Medicamento._separar_presentacion(presentacion)

        # Buscar en catálogo por coincidencia exacta
        query_catalogo = """
            SELECT id_catalogo FROM catalogo_medicamentos
            WHERE nombre = ? AND ISNULL(forma_farmaceutica, '') = ? AND ISNULL(concentracion, '') = ?
        """
        result = Database.execute_query(query_catalogo, (nombre, forma_farmaceutica, concentracion))
        if result:
            id_catalogo = result[0][0]
        else:
            # Si no existe, se crea (esto puede ser útil en otras partes, pero en el formulario principal se evita)
            insert_catalogo = """
                INSERT INTO catalogo_medicamentos (nombre, forma_farmaceutica, concentracion)
                VALUES (?, ?, ?)
            """
            Database.execute_non_query(insert_catalogo, (nombre, forma_farmaceutica, concentracion))
            result_id = Database.execute_query("SELECT SCOPE_IDENTITY() as id")
            id_catalogo = result_id[0][0] if result_id else None

        if not id_catalogo:
            return False

        tabla = "inventario_farmacia" if area_id == 1 else "inventario_odontologia"
        query = f"""
            INSERT INTO {tabla} 
            (id_catalogo, ubicacion, caducidad, stock, es_donacion, donante, observaciones)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        return Database.execute_non_query(query, (
            id_catalogo, ubicacion, caducidad, stock, es_donacion, donante, observaciones
        ))

    @staticmethod
    def crear_desde_catalogo(area_id, id_catalogo, ubicacion, caducidad,
                              stock, es_donacion, donante, observaciones):
        """Agrega un medicamento al inventario usando un id_catalogo existente.
           No crea nuevas entradas en el catálogo."""
        tabla = "inventario_farmacia" if area_id == 1 else "inventario_odontologia"
        query = f"""
            INSERT INTO {tabla} 
            (id_catalogo, ubicacion, caducidad, stock, es_donacion, donante, observaciones)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        return Database.execute_non_query(query, (
            id_catalogo, ubicacion, caducidad, stock, es_donacion, donante, observaciones
        ))

    @staticmethod
    def actualizar(id_medicamento, nombre, presentacion, ubicacion,
                   caducidad, observaciones, stock, es_donacion, donante=""):
        """Actualiza un medicamento existente (incluyendo su entrada en catálogo)"""
        med = Medicamento.obtener_por_id(id_medicamento)
        if not med:
            return False

        forma_farmaceutica, concentracion = Medicamento._separar_presentacion(presentacion)

        # Actualizar catálogo
        update_catalogo = """
            UPDATE catalogo_medicamentos
            SET nombre = ?, forma_farmaceutica = ?, concentracion = ?
            WHERE id_catalogo = ?
        """
        Database.execute_non_query(update_catalogo, (nombre, forma_farmaceutica, concentracion, med['id_catalogo']))

        tabla = "inventario_farmacia" if med['area_id'] == 1 else "inventario_odontologia"
        query = f"""
            UPDATE {tabla}
            SET ubicacion = ?, caducidad = ?, stock = ?, es_donacion = ?, donante = ?, observaciones = ?
            WHERE id_inventario = ?
        """
        return Database.execute_non_query(query, (
            ubicacion, caducidad, stock, es_donacion, donante, observaciones, id_medicamento
        ))

    @staticmethod
    def eliminar(id_medicamento):
        med = Medicamento.obtener_por_id(id_medicamento)
        if not med:
            return False
        tabla = "inventario_farmacia" if med['area_id'] == 1 else "inventario_odontologia"
        query = f"DELETE FROM {tabla} WHERE id_inventario = ?"
        return Database.execute_non_query(query, (id_medicamento,))