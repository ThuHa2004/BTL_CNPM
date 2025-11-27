import pyodbc
import os
from dotenv import load_dotenv

load_dotenv()


def get_connection():
    """
    Tạo kết nối đến SQL Server
    Hỗ trợ cả Windows Authentication và SQL Server Authentication
    """
    server = os.getenv('DB_SERVER', 'LAPTOP-B1UCCBCI\SQL_TTTHA')
    database = os.getenv('DB_DATABASE', 'BTL_CNPMDDKM')
    username = os.getenv('DB_USERNAME', 'sa')
    password = os.getenv('DB_PASSWORD', '123')
    trusted_connection = os.getenv('DB_TRUSTED_CONNECTION', 'yes').lower()

    # Nếu dùng Windows Authentication
    if not username or trusted_connection == 'yes':
        connection_string = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'Trusted_Connection=yes;'
        )
        print("[INFO] Kết nối bằng Windows Authentication")
    else:
        # Nếu dùng SQL Server Authentication
        connection_string = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password};'
        )
        print("[INFO] Kết nối bằng SQL Server Authentication")

    try:
        conn = pyodbc.connect(connection_string)
        print(f"[SUCCESS] Kết nối database thành công: {database}")
        return conn
    except pyodbc.Error as e:
        print(f"[ERROR] Lỗi kết nối database: {str(e)}")
        raise


def execute_query(query, params=None):
    """
    Thực thi query và trả về kết quả
    """
    conn = get_connection()
    cursor = conn.cursor()

    try:
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)

        # Nếu là SELECT query
        if query.strip().upper().startswith('SELECT'):
            results = cursor.fetchall()
            return results
        else:
            # Nếu là INSERT, UPDATE, DELETE
            conn.commit()
            return cursor.rowcount
    except Exception as e:
        conn.rollback()
        print(f"[ERROR] Lỗi thực thi query: {str(e)}")
        raise
    finally:
        cursor.close()
        conn.close()


def test_connection():
    """
    Test kết nối database
    """
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT @@VERSION")
        version = cursor.fetchone()[0]
        print(f"\n[SUCCESS] Kết nối thành công!")
        print(f"SQL Server Version: {version[:100]}...")
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        print(f"\n[ERROR] Kết nối thất bại: {str(e)}")
        return False


if __name__ == "__main__":
    print("========================================")
    print("  TEST KẾT NỐI DATABASE")
    print("========================================\n")
    test_connection()