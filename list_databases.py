"""
Скрипт для перегляду списку всіх доступних баз даних на SQL Server
Запустіть цей скрипт, щоб дізнатися ім'я вашої бази даних
"""

import pyodbc
import sys

# Налаштування підключення (підключаємося до master для отримання списку БД)
SERVER = 'localhost\\SQLEXPRESS'
DRIVER = '{ODBC Driver 17 for SQL Server}'

# Пробуємо різні драйвери, якщо перший не працює
DRIVERS_TO_TRY = [
    '{ODBC Driver 17 for SQL Server}',
    '{ODBC Driver 18 for SQL Server}',
    '{SQL Server Native Client 11.0}',
    '{SQL Server}',
]

def list_databases():
    """Отримати список всіх баз даних"""
    print("🔍 Пошук доступних баз даних на сервері...")
    print(f"Сервер: {SERVER}")
    print("-" * 60)
    
    for driver in DRIVERS_TO_TRY:
        try:
            print(f"\nСпроба підключення з драйвером: {driver}")
            conn_str = (
                f"DRIVER={driver};"
                f"SERVER={SERVER};"
                f"DATABASE=master;"
                f"Trusted_Connection=yes;"
            )
            
            conn = pyodbc.connect(conn_str, timeout=5)
            cursor = conn.cursor()
            
            print("✅ Підключення успішно!")
            print(f"Використовуваний драйвер: {driver}\n")
            
            # Отримуємо список всіх баз даних
            cursor.execute("""
                SELECT name, database_id, create_date
                FROM sys.databases
                WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb')
                ORDER BY name
            """)
            
            databases = cursor.fetchall()
            
            if databases:
                print(f"📊 Знайдено користувацьких баз даних: {len(databases)}\n")
                print("Список баз даних:")
                print("-" * 60)
                for i, db in enumerate(databases, 1):
                    print(f"{i}. {db[0]}")
                    print(f"   ID: {db[1]}, Створена: {db[2]}")
                    print()
            else:
                print("⚠️  Користувацьких баз даних не знайдено")
                print("Доступні тільки системні бази: master, tempdb, model, msdb")
            
            # Також показуємо системні бази
            cursor.execute("""
                SELECT name FROM sys.databases
                WHERE name IN ('master', 'tempdb', 'model', 'msdb')
                ORDER BY name
            """)
            system_dbs = [row[0] for row in cursor.fetchall()]
            if system_dbs:
                print("\nСистемні бази даних:")
                for db in system_dbs:
                    print(f"  - {db}")
            
            conn.close()
            
            print("\n" + "=" * 60)
            print("💡 Скопіюйте ім'я потрібної бази даних та вкажіть її в config.py")
            print("   в параметрі 'database'")
            print("=" * 60)
            
            return True
            
        except pyodbc.Error as e:
            print(f"❌ Помилка з драйвером {driver}: {e}")
            continue
        except Exception as e:
            print(f"❌ Неочікувана помилка: {e}")
            continue
    
    print("\n" + "=" * 60)
    print("❌ Не вдалося підключитися жодним драйвером")
    print("\nМожливі рішення:")
    print("1. Переконайтеся, що SQL Server запущено")
    print("2. Перевірте правильність імені сервера")
    print("3. Встановіть ODBC Driver для SQL Server:")
    print("   https://docs.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server")
    print("=" * 60)
    return False

if __name__ == '__main__':
    list_databases()
