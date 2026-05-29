from config.database import Database
from datetime import date


class Movimiento:
    """Modelo de Movimientos para ambas áreas"""

    @staticmethod
    def registrar(id_inventario, tipo, cantidad, quien_recibe, quien_retira, motivo, fecha=None):
        """Registra un nuevo movimiento"""
        if fecha is None:
            fecha = date.today()

        # Determinar la tabla según el área (necesitamos saber el área del medicamento)
        from models.medicamento import Medicamento
        med = Medicamento.obtener_por_id(id_inventario)
        if not med:
            print(f"Medicamento {id_inventario} no encontrado")
            return False

        tabla = "movimientos_farmacia" if med['area_id'] == 1 else "movimientos_odontologia"
        query = f"""
            INSERT INTO {tabla} 
            (id_inventario, tipo, cantidad, quien_recibe, quien_retira, motivo, fecha)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        return Database.execute_non_query(query, (
            id_inventario, tipo, cantidad, quien_recibe, quien_retira, motivo, fecha
        ))

    @staticmethod
    def obtener_todos_por_area(area_id, fecha_desde=None, fecha_hasta=None, tipo_filtro=None):
        """Obtiene todos los movimientos de un área con filtros opcionales"""
        try:
            tabla_movimientos = "movimientos_farmacia" if area_id == 1 else "movimientos_odontologia"
            tabla_inventario = "inventario_farmacia" if area_id == 1 else "inventario_odontologia"

            query = f"""
                SELECT m.id_movimiento, m.id_inventario, m.tipo, m.cantidad, 
                       m.quien_recibe, m.quien_retira, m.motivo, m.fecha,
                       c.nombre as medicamento_nombre
                FROM {tabla_movimientos} m
                JOIN {tabla_inventario} i ON m.id_inventario = i.id_inventario
                JOIN catalogo_medicamentos c ON i.id_catalogo = c.id_catalogo
                WHERE 1=1
            """
            params = []

            if fecha_desde:
                query += " AND m.fecha >= ?"
                params.append(fecha_desde)
            if fecha_hasta:
                query += " AND m.fecha <= ?"
                params.append(fecha_hasta)
            if tipo_filtro and tipo_filtro != "Todos":
                query += " AND m.tipo = ?"
                params.append(tipo_filtro.lower())

            query += " ORDER BY m.fecha DESC, m.id_movimiento DESC"

            result = Database.execute_query(query, tuple(params) if params else None)
            movimientos = []
            if result:
                for row in result:
                    movimientos.append({
                        'id_movimiento': row[0],
                        'id_inventario': row[1],
                        'tipo': row[2],
                        'cantidad': row[3],
                        'quien_recibe': row[4] or "",
                        'quien_retira': row[5] or "",
                        'motivo': row[6] or "",
                        'fecha': row[7],
                        'medicamento_nombre': row[8]
                    })
            return movimientos
        except Exception as e:
            print(f"Error obteniendo movimientos: {e}")
            return []

    @staticmethod
    def obtener_por_id(movimiento_id, area_id):
        """Obtiene un movimiento específico por su ID y área"""
        try:
            tabla = "movimientos_farmacia" if area_id == 1 else "movimientos_odontologia"
            query = f"""
                SELECT id_movimiento, id_inventario, tipo, cantidad, 
                       quien_recibe, quien_retira, motivo, fecha
                FROM {tabla}
                WHERE id_movimiento = ?
            """
            result = Database.execute_query(query, (movimiento_id,))
            if result:
                row = result[0]
                return {
                    'id_movimiento': row[0],
                    'id_medicamento': row[1],
                    'tipo': row[2],
                    'cantidad': row[3],
                    'quien_recibe': row[4] or "",
                    'quien_retira': row[5] or "",
                    'motivo': row[6] or "",
                    'fecha': row[7]
                }
            return None
        except Exception as e:
            print(f"Error obteniendo movimiento por ID: {e}")
            return None

    @staticmethod
    def actualizar(movimiento_id, area_id, id_inventario, tipo, cantidad,
                   quien_recibe, quien_retira, motivo, fecha):
        """Actualiza un movimiento existente"""
        try:
            tabla = "movimientos_farmacia" if area_id == 1 else "movimientos_odontologia"
            query = f"""
                UPDATE {tabla}
                SET id_inventario = ?, tipo = ?, cantidad = ?, 
                    quien_recibe = ?, quien_retira = ?, motivo = ?, fecha = ?
                WHERE id_movimiento = ?
            """
            return Database.execute_non_query(query, (
                id_inventario, tipo, cantidad, quien_recibe, quien_retira, motivo, fecha, movimiento_id
            ))
        except Exception as e:
            print(f"Error actualizando movimiento: {e}")
            return False

    @staticmethod
    def eliminar(movimiento_id, area_id):
        """Elimina un movimiento"""
        try:
            tabla = "movimientos_farmacia" if area_id == 1 else "movimientos_odontologia"
            query = f"DELETE FROM {tabla} WHERE id_movimiento = ?"
            return Database.execute_non_query(query, (movimiento_id,))
        except Exception as e:
            print(f"Error eliminando movimiento: {e}")
            return False

    @staticmethod
    def obtener_por_medicamento(medicamento_id, area_id, limite=50):
        """Obtiene los últimos movimientos de un medicamento específico"""
        try:
            tabla = "movimientos_farmacia" if area_id == 1 else "movimientos_odontologia"
            query = f"""
                SELECT id_movimiento, tipo, cantidad, quien_recibe, quien_retira, motivo, fecha
                FROM {tabla}
                WHERE id_inventario = ?
                ORDER BY fecha DESC, id_movimiento DESC
                LIMIT ?
            """
            result = Database.execute_query(query, (medicamento_id, limite))
            movimientos = []
            if result:
                for row in result:
                    movimientos.append({
                        'id_movimiento': row[0],
                        'tipo': row[1],
                        'cantidad': row[2],
                        'quien_recibe': row[3] or "",
                        'quien_retira': row[4] or "",
                        'motivo': row[5] or "",
                        'fecha': row[6]
                    })
            return movimientos
        except Exception as e:
            print(f"Error obteniendo movimientos por medicamento: {e}")
            return []