import pyodbc
import os
from dotenv import load_dotenv

# Cargar variables de entorno
load_dotenv()


class DatabaseConfig:
    """Configuración de la base de datos"""

    SERVER = os.getenv('DB_SERVER', 'localhost\\SQLEXPRESS')
    DATABASE = os.getenv('DB_NAME', 'farmaciaUAEH')
    TRUSTED_CONNECTION = os.getenv('DB_TRUSTED', 'yes')

    @classmethod
    def get_connection_string(cls):
        return (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={cls.SERVER};'
            f'DATABASE={cls.DATABASE};'
            f'Trusted_Connection={cls.TRUSTED_CONNECTION};'
        )

    @classmethod
    def get_connection(cls):
        try:
            conn = pyodbc.connect(cls.get_connection_string(), timeout=5)
            return conn
        except Exception as e:
            print(f"Error de conexión: {e}")
            return None


class Database:
    """Clase para operaciones de base de datos"""

    @staticmethod
    def execute_query(query, params=None):
        conn = DatabaseConfig.get_connection()
        if not conn:
            return None

        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            result = cursor.fetchall()
            return result
        except Exception as e:
            print(f"Error en consulta: {e}")
            return None
        finally:
            conn.close()

    @staticmethod
    def execute_non_query(query, params=None):
        conn = DatabaseConfig.get_connection()
        if not conn:
            return False

        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
            return True
        except Exception as e:
            print(f"Error en operación: {e}")
            conn.rollback()
            return False
        finally:
            conn.close()