"""
Скрипт для перевірки підключення до MSSQL бази даних
Запустіть цей скрипт перед запуском веб-додатку для перевірки налаштувань
"""

import pyodbc
import sys

# Імпорт конфігурації
try:
    from config import DB_CONFIG
except ImportError:
    print("❌ Файл config.py не знайдено!")
    print("Створіть файл config.py з налаштуваннями підключення.")
    sys.exit(1)

def test_connection():
    """Тестування підключення до БД"""
    print("🔍 Перевірка підключення до MSSQL...")
    print(f"Сервер: {DB_CONFIG['server']}")
    print(f"База даних: {DB_CONFIG['database']}")
    print(f"Драйвер: {DB_CONFIG['driver']}")
    print("-" * 50)
    
    try:
        # Формуємо рядок підключення
        trusted = DB_CONFIG.get('trusted_connection', True)
        if trusted is True or trusted == 'yes':
            conn_str = (
                f"DRIVER={DB_CONFIG['driver']};"
                f"SERVER={DB_CONFIG['server']};"
                f"DATABASE={DB_CONFIG['database']};"
                f"Trusted_Connection=yes;"
            )
            print("Аутентифікація: Windows (Trusted Connection)")
        else:
            conn_str = (
                f"DRIVER={DB_CONFIG['driver']};"
                f"SERVER={DB_CONFIG['server']};"
                f"DATABASE={DB_CONFIG['database']};"
                f"UID={DB_CONFIG['uid']};"
                f"PWD={DB_CONFIG['pwd']};"
            )
            print(f"Аутентифікація: SQL Server (користувач: {DB_CONFIG['uid']})")
        
        print("\nСпроба підключення...")
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        # Отримуємо версію SQL Server
        cursor.execute("SELECT @@VERSION")
        version = cursor.fetchone()[0]
        print("\n✅ Підключення успішно!")
        print(f"\nВерсія SQL Server:\n{version}")
        
        # Отримуємо список таблиць
        cursor.execute("""
            SELECT TABLE_NAME 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_TYPE = 'BASE TABLE'
            ORDER BY TABLE_NAME
        """)
        tables = cursor.fetchall()
        
        print(f"\n📊 Знайдено таблиць: {len(tables)}")
        if tables:
            print("\nСписок таблиць:")
            for i, table in enumerate(tables, 1):
                print(f"  {i}. {table[0]}")
        
        conn.close()
        print("\n✅ Всі перевірки пройдено успішно!")
        print("Можна запускати веб-додаток: python app.py")
        return True
        
    except pyodbc.Error as e:
        print(f"\n❌ Помилка підключення: {e}")
        print("\nМожливі причини:")
        print("1. SQL Server не запущено")
        print("2. Неправильне ім'я сервера або бази даних")
        print("3. Не встановлено ODBC Driver для SQL Server")
        print("4. Немає прав доступу до бази даних")
        print("\nПеревірте налаштування в config.py")
        return False
    except Exception as e:
        print(f"\n❌ Неочікувана помилка: {e}")
        return False

if __name__ == '__main__':
    test_connection()
