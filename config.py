# Конфігурація підключення до MSSQL
# Змініть ці параметри на ваші реальні дані

DB_CONFIG = {
    # Ім'я сервера SQL Server
    # Приклади:
    # - 'localhost' - для локального сервера за замовчуванням
    # - 'localhost\\SQLEXPRESS' - для SQL Server Express
    # - 'DESKTOP-XXXXX\\SQLEXPRESS' - для іменованого екземпляра
    # Сервер: localhost\SQLEXPRESS (MSSQL16.SQLEXPRESS)
    'server': 'localhost\\SQLEXPRESS',
    
    # Ім'я бази даних
    'database': 'bps',
    
    # Драйвер ODBC
    # Для Windows зазвичай використовується один з:
    # - '{ODBC Driver 17 for SQL Server}' - рекомендується
    # - '{ODBC Driver 18 for SQL Server}' - нова версія
    # - '{SQL Server}' - старий драйвер
    'driver': '{ODBC Driver 17 for SQL Server}',
    
    # Використовувати Windows аутентифікацію (Trusted Connection)
    # Якщо True - використовується поточний користувач Windows
    # Якщо False - потрібно вказати uid і pwd нижче
    'trusted_connection': True,
    
    # Розкоментуйте та заповніть, якщо використовуєте SQL Server аутентифікацію:
    # 'uid': 'your_username',
    # 'pwd': 'your_password',
}
