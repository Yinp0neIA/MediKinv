from config.database import Database
from datetime import datetime, date


class Movimiento:
    """Modelo de Movimiento - Adaptado para movimientos separados"""

    @staticmethod
    def registrar(medicamento_id, tipo, cantidad, quien_recibe,
                  quien_retira, motivo, fecha=None):
        """Registra un movimiento en la tabla correspondiente"""
        from models.medicamento import Medicamento
        medicamento = Medicamento.obtener_por_id(medicamento_id)
        if not medicamento:
            return False

        if fecha is None:
            fecha = date.today()

        tabla_movimientos = "movimientos_farmacia" if medicamento['area_id'] == 1 else "movimientos_odontologia"

        query = f"""
                INSERT INTO {tabla_movimientos} 
                (id_inventario, tipo, cantidad, quien_recibe, quien_retira, motivo, fecha)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """
        return Database.execute_non_query(query, (
            medicamento_id, tipo, cantidad, quien_recibe, quien_retira, motivo, fecha
        ))

    @staticmethod
    def obtener_por_medicamento(medicamento_id):
        """Obtiene los movimientos de un medicamento específico"""
        from models.medicamento import Medicamento
        medicamento = Medicamento.obtener_por_id(medicamento_id)
        if not medicamento:
            return []

        tabla_movimientos = "movimientos_farmacia" if medicamento['area_id'] == 1 else "movimientos_odontologia"

        query = f"""
                SELECT id_movimiento, tipo, cantidad, quien_recibe, 
                       quien_retira, motivo, fecha
                FROM {tabla_movimientos} 
                WHERE id_inventario = ?
                ORDER BY fecha DESC, id_movimiento DESC
            """
        result = Database.execute_query(query, (medicamento_id,))

        movimientos = []
        if result:
            for row in result:
                movimientos.append({
                    'id_movimiento': row[0],
                    'tipo': row[1],
                    'cantidad': row[2],
                    'quien_recibe': row[3] if row[3] else "",
                    'quien_retira': row[4] if row[4] else "",
                    'motivo': row[5] if row[5] else "",
                    'fecha': row[6],
                    'id_medicamento': medicamento_id
                })
        return movimientos

    @staticmethod
    def obtener_todos_por_area(area_id, fecha_desde=None, fecha_hasta=None, tipo=None):
        """Obtiene todos los movimientos de un área específica"""
        tabla_movimientos = "movimientos_farmacia" if area_id == 1 else "movimientos_odontologia"
        tabla_inventario = "inventario_farmacia" if area_id == 1 else "inventario_odontologia"

        query = f"""
                SELECT m.id_movimiento, m.tipo, m.cantidad, m.quien_recibe, 
                       m.quien_retira, m.motivo, m.fecha, c.nombre as medicamento_nombre, m.id_inventario
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
        if tipo and tipo != "Todos":
            query += " AND m.tipo = ?"
            params.append(tipo.lower())

        query += " ORDER BY m.fecha DESC, m.id_movimiento DESC"

        result = Database.execute_query(query, tuple(params) if params else None)

        movimientos = []
        if result:
            for row in result:
                movimientos.append({
                    'id_movimiento': row[0],
                    'tipo': row[1],
                    'cantidad': row[2],
                    'quien_recibe': row[3] if row[3] else "",
                    'quien_retira': row[4] if row[4] else "",
                    'motivo': row[5] if row[5] else "",
                    'fecha': row[6],
                    'medicamento_nombre': row[7],
                    'id_medicamento': row[8]
                })
        return movimientos

    @staticmethod
    def obtener_por_id(id_movimiento, area_id):
        """Obtiene un movimiento por su ID"""
        tabla_movimientos = "movimientos_farmacia" if area_id == 1 else "movimientos_odontologia"
        query = f"""
                SELECT id_movimiento, id_inventario, tipo, cantidad, quien_recibe, 
                       quien_retira, motivo, fecha
                FROM {tabla_movimientos} 
                WHERE id_movimiento = ?
            """
        result = Database.execute_query(query, (id_movimiento,))
        if not result:
            return None
        row = result[0]
        return {
            'id_movimiento': row[0],
            'id_medicamento': row[1],
            'tipo': row[2],
            'cantidad': row[3],
            'quien_recibe': row[4] if row[4] else "",
            'quien_retira': row[5] if row[5] else "",
            'motivo': row[6] if row[6] else "",
            'fecha': row[7]
        }