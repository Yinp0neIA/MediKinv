from config.database import Database


class Usuario:
    """Modelo de Usuario"""

    def __init__(self, id_usuario=None, nombre_usuario=None, contrasena=None):
        self.id_usuario = id_usuario
        self.nombre_usuario = nombre_usuario
        self.contrasena = contrasena

    @staticmethod
    def validar_credenciales(usuario, contrasena):
        query = """
            SELECT id_usuario, nombre_usuario 
            FROM usuarios 
            WHERE nombre_usuario = ? AND contrasena = ?
        """
        result = Database.execute_query(query, (usuario, contrasena))

        if result and len(result) > 0:
            usuario_data = result[0]
            return Usuario(
                id_usuario=usuario_data[0],
                nombre_usuario=usuario_data[1]
            )
        return None

    @staticmethod
    def registrar_usuario(usuario, contrasena):
        query_check = "SELECT COUNT(*) FROM usuarios WHERE nombre_usuario = ?"
        result = Database.execute_query(query_check, (usuario,))

        if result and result[0][0] > 0:
            return False, "El nombre de usuario ya existe"

        query_insert = """
            INSERT INTO usuarios (nombre_usuario, contrasena)
            VALUES (?, ?)
        """
        success = Database.execute_non_query(query_insert, (usuario, contrasena))

        if success:
            return True, "Usuario registrado exitosamente"
        return False, "Error al registrar usuario"

    @staticmethod
    def obtener_areas():
        query = "SELECT id_areas, nombre_area FROM areas"
        result = Database.execute_query(query)

        areas = []
        if result:
            for row in result:
                areas.append({
                    'id': row[0],
                    'nombre': row[1]
                })
        return areas