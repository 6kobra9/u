from flask import Flask, render_template, request, jsonify
import pyodbc
import os
from datetime import datetime
import pandas as pd
import re

# Импорт конфигурации
try:
    from config import DB_CONFIG
except ImportError:
    # Если config.py не найден, используем значения по умолчанию
    DB_CONFIG = {
        'server': 'localhost',
        'database': 'your_database',
        'driver': '{ODBC Driver 17 for SQL Server}',
        'trusted_connection': True,
    }

app = Flask(__name__)

def get_db_connection():
    """Создает подключение к базе данных MSSQL"""
    # Список драйверов для попытки подключения
    drivers_to_try = [
        DB_CONFIG.get('driver', '{ODBC Driver 17 for SQL Server}'),
        '{ODBC Driver 17 for SQL Server}',
        '{ODBC Driver 18 for SQL Server}',
        '{SQL Server Native Client 11.0}',
        '{SQL Server}',
    ]
    
    trusted = DB_CONFIG.get('trusted_connection', True)
    
    for driver in drivers_to_try:
        try:
            # Формируем строку подключения
            if trusted is True or trusted == 'yes':
                conn_str = (
                    f"DRIVER={driver};"
                    f"SERVER={DB_CONFIG['server']};"
                    f"DATABASE={DB_CONFIG['database']};"
                    f"Trusted_Connection=yes;"
                )
            else:
                conn_str = (
                    f"DRIVER={driver};"
                    f"SERVER={DB_CONFIG['server']};"
                    f"DATABASE={DB_CONFIG['database']};"
                    f"UID={DB_CONFIG['uid']};"
                    f"PWD={DB_CONFIG['pwd']};"
                )
            
            conn = pyodbc.connect(conn_str, timeout=5)
            return conn
        except pyodbc.Error:
            continue  # Пробуем следующий драйвер
        except Exception:
            continue  # Пробуем следующий драйвер
    
    # Якщо жоден драйвер не спрацював
    raise Exception(f"Не вдалося підключитися до БД. Перевірте:\n"
                   f"1. SQL Server запущено\n"
                   f"2. Правильність налаштувань у config.py\n"
                   f"3. Встановлено ODBC Driver для SQL Server")

@app.route('/')
def index():
    """Главная страница"""
    return render_template('index.html')

@app.route('/api/test-connection')
def test_connection():
    """Проверка подключения к БД"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT @@VERSION")
        version = cursor.fetchone()[0]
        conn.close()
        return jsonify({'success': True, 'message': 'Підключення успішно', 'version': version})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/tables')
def get_tables():
    """Получить список всех таблиц в БД"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        # Получаем список таблиц
        cursor.execute("""
            SELECT TABLE_NAME 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_TYPE = 'BASE TABLE'
            ORDER BY TABLE_NAME
        """)
        tables = [row[0] for row in cursor.fetchall()]
        conn.close()
        return jsonify({'success': True, 'tables': tables})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/table/<table_name>')
def get_table_data(table_name):
    """Получить данные из таблицы"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Получаем информацию о колонках
        cursor.execute(f"""
            SELECT COLUMN_NAME, DATA_TYPE 
            FROM INFORMATION_SCHEMA.COLUMNS 
            WHERE TABLE_NAME = '{table_name}'
            ORDER BY ORDINAL_POSITION
        """)
        columns_info = cursor.fetchall()
        columns = [col[0] for col in columns_info]
        
        # Получаем данные
        cursor.execute(f"SELECT * FROM [{table_name}]")
        rows = cursor.fetchall()
        conn.close()
        
        data = []
        for row in rows:
            row_dict = {}
            for i, col in enumerate(columns):
                value = row[i]
                # Преобразуем datetime в строку
                if isinstance(value, datetime):
                    value = value.isoformat()
                row_dict[col] = value
            data.append(row_dict)
        
        return jsonify({
            'success': True, 
            'columns': columns,
            'columns_info': {col[0]: col[1] for col in columns_info},
            'data': data
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/table-data/<table_name>')
def get_table_data_for_select(table_name):
    """Получить данные из таблицы для выпадающих списков"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Для unit показуємо off_name замість name
        if table_name.lower() == 'unit':
            cursor.execute("SELECT id, off_name FROM [unit] ORDER BY off_name")
            rows = cursor.fetchall()
            data = [{'id': row[0], 'off_name': row[1] if row[1] is not None else ''} for row in rows]
        else:
            cursor.execute(f"SELECT id, name FROM [{table_name}] ORDER BY name")
            rows = cursor.fetchall()
            data = [{'id': row[0], 'name': row[1]} for row in rows]
        conn.close()
        
        return jsonify({
            'success': True,
            'data': data
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/query', methods=['POST'])
def execute_query():
    """Выполнить SQL запрос"""
    try:
        data = request.json
        query = data.get('query', '').strip()
        
        if not query:
            return jsonify({'success': False, 'error': 'Запит не може бути порожнім'}), 400
        
        # Безпека: дозволяємо тільки SELECT запити для читання
        query_upper = query.upper().strip()
        # Дозволяємо SELECT та WITH (для CTE)
        if not (query_upper.startswith('SELECT') or query_upper.startswith('WITH')):
            return jsonify({'success': False, 'error': 'Дозволені тільки SELECT запити'}), 400
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query)
        
        # Получаем названия колонок
        columns = [column[0] for column in cursor.description]
        
        # Получаем все данные
        rows = cursor.fetchall()
        conn.close()
        
        data = []
        for row in rows:
            row_dict = {}
            for i, col in enumerate(columns):
                value = row[i]
                if isinstance(value, datetime):
                    value = value.isoformat()
                row_dict[col] = value
            data.append(row_dict)
        
        return jsonify({'success': True, 'columns': columns, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/structure-tree')
def get_structure_tree():
    """Повертає ієрархію підпорядкування (id, parent_id, name) для деревоподібної схеми."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT s.unit_id, s.parent_id, u.off_name AS name
            FROM subordination s
            JOIN unit u ON u.id = s.unit_id
            ORDER BY s.parent_id, s.unit_id
        """)
        rows = cursor.fetchall()
        conn.close()
        data = []
        for row in rows:
            data.append({
                'id': row[0],
                'parent_id': row[1],
                'name': row[2] or ''
            })
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/insert', methods=['POST'])
def insert_data():
    """Вставить данные в таблицу"""
    try:
        data = request.json
        table_name = data.get('table')
        values = data.get('values', {})
        
        if not table_name or not values:
            return jsonify({'success': False, 'error': 'Не вказано таблицю або значення'}), 400
        
        columns = ', '.join([f"[{k}]" for k in values.keys()])
        placeholders = ', '.join(['?' for _ in values])
        sql = f"INSERT INTO [{table_name}] ({columns}) VALUES ({placeholders})"
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(sql, list(values.values()))
        conn.commit()
        conn.close()
        
        return jsonify({'success': True, 'message': 'Дані успішно додано'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/get-row', methods=['POST'])
def get_row():
    """Отримати повну запис з таблиці за ключовими полями"""
    try:
        data = request.json
        table_name = data.get('table')
        where = data.get('where', {})
        
        if not table_name or not where:
            return jsonify({'success': False, 'error': 'Не вказано таблицю або умови пошуку'}), 400
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Формуємо WHERE умову
        where_clause = ' AND '.join([f"[{k}] = ?" for k in where.keys()])
        sql = f"SELECT * FROM [{table_name}] WHERE {where_clause}"
        
        cursor.execute(sql, list(where.values()))
        row = cursor.fetchone()
        
        if not row:
            conn.close()
            return jsonify({'success': False, 'error': 'Запис не знайдено'}), 404
        
        # Отримуємо назви колонок
        columns = [column[0] for column in cursor.description]
        
        # Формуємо словник з даними
        row_dict = {}
        for i, col in enumerate(columns):
            value = row[i]
            if isinstance(value, datetime):
                value = value.isoformat()
            row_dict[col] = value
        
        conn.close()
        
        return jsonify({
            'success': True,
            'data': row_dict
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/update', methods=['POST'])
def update_data():
    """Обновить данные в таблице"""
    try:
        data = request.json
        table_name = data.get('table')
        values = data.get('values', {})
        where = data.get('where', {})
        
        if not table_name or not values or not where:
            return jsonify({'success': False, 'error': 'Недостатньо даних для оновлення'}), 400
        
        set_clause = ', '.join([f"[{k}] = ?" for k in values.keys()])
        where_clause = ' AND '.join([f"[{k}] = ?" for k in where.keys()])
        sql = f"UPDATE [{table_name}] SET {set_clause} WHERE {where_clause}"
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(sql, list(values.values()) + list(where.values()))
        conn.commit()
        conn.close()
        
        return jsonify({'success': True, 'message': 'Дані успішно оновлено'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/update-nrc-structure', methods=['POST'])
def update_nrc_structure():
    """
    Оновлення логічних полів структури НРК:
    - type_bps
    - unit_structure_id
    
    Оновлюємо таблицю bps_structure за unit_id.
    Поля l1, l2 зараз відображаються тільки для перегляду.
    """
    try:
        data = request.json or {}
        unit_id = data.get('unit_id')
        type_bps = data.get('type_bps')
        unit_structure_id = data.get('unit_structure_id')

        if not unit_id:
            return jsonify({'success': False, 'error': 'Не вказано unit_id'}), 400

        conn = get_db_connection()
        cursor = conn.cursor()

        # Формуємо SET тільки з переданих полів
        set_parts = []
        params = []

        if type_bps is not None:
            set_parts.append("[type_bps] = ?")
            params.append(type_bps)

        if unit_structure_id is not None and unit_structure_id != '':
            set_parts.append("[unit_structure_id] = ?")
            params.append(unit_structure_id)

        if not set_parts:
            conn.close()
            return jsonify({'success': False, 'error': 'Немає даних для оновлення'}), 400

        set_clause = ', '.join(set_parts)
        sql = f"UPDATE [bps_structure] SET {set_clause} WHERE [unit_id] = ?"
        params.append(unit_id)

        cursor.execute(sql, params)
        conn.commit()
        conn.close()

        return jsonify({'success': True, 'message': 'Структуру НРК успішно оновлено'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
@app.route('/api/delete', methods=['POST'])
def delete_data():
    """Удалить данные из таблицы"""
    try:
        data = request.json
        table_name = data.get('table')
        where = data.get('where', {})
        
        if not table_name or not where:
            return jsonify({'success': False, 'error': 'Недостатньо даних для видалення'}), 400
        
        where_clause = ' AND '.join([f"[{k}] = ?" for k in where.keys()])
        sql = f"DELETE FROM [{table_name}] WHERE {where_clause}"
        
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(sql, list(where.values()))
        conn.commit()
        conn.close()
        
        return jsonify({'success': True, 'message': 'Дані успішно видалено'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/table-info/<table_name>')
def get_table_info(table_name):
    """Получить подробную информацию о таблице"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Получаем информацию о колонках
        cursor.execute(f"""
            SELECT 
                COLUMN_NAME,
                DATA_TYPE,
                IS_NULLABLE,
                COLUMN_DEFAULT,
                CHARACTER_MAXIMUM_LENGTH
            FROM INFORMATION_SCHEMA.COLUMNS 
            WHERE TABLE_NAME = '{table_name}'
            ORDER BY ORDINAL_POSITION
        """)
        
        columns_info = []
        for row in cursor.fetchall():
            columns_info.append({
                'name': row[0],
                'type': row[1],
                'nullable': row[2],
                'default': row[3],
                'max_length': row[4]
            })
        
        # Получаем количество записей
        cursor.execute(f"SELECT COUNT(*) FROM [{table_name}]")
        row_count = cursor.fetchone()[0]
        
        conn.close()
        
        return jsonify({
            'success': True,
            'columns': columns_info,
            'row_count': row_count
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/load-materials-to-table', methods=['POST'])
def load_materials_to_table():
    """Завантажити дані з Material у нову таблицю"""
    try:
        data = request.json
        table_name = data.get('table_name', '').strip()
        
        if not table_name:
            return jsonify({'success': False, 'error': 'Не вказано ім\'я таблиці'}), 400
        
        # Валідація імені таблиці (захист від SQL injection)
        if not all(c.isalnum() or c in ['_', '-'] for c in table_name):
            return jsonify({'success': False, 'error': 'Невірне ім\'я таблиці. Дозволені тільки букви, цифри, підкреслення та дефіси'}), 400
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Перевіряємо, чи існує таблиця Material
        cursor.execute("""
            SELECT COUNT(*) 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_NAME = 'Material'
        """)
        if cursor.fetchone()[0] == 0:
            conn.close()
            return jsonify({'success': False, 'error': 'Таблиця Material не знайдена'}), 404
        
        # Отримуємо структуру таблиці Material
        cursor.execute("""
            SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
            FROM INFORMATION_SCHEMA.COLUMNS 
            WHERE TABLE_NAME = 'Material'
            ORDER BY ORDINAL_POSITION
        """)
        columns_info = cursor.fetchall()
        
        if not columns_info:
            conn.close()
            return jsonify({'success': False, 'error': 'Не вдалося отримати структуру таблиці Material'}), 500
        
        # Перевіряємо, чи існує вже таблиця з таким ім'ям
        cursor.execute(f"""
            SELECT COUNT(*) 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_NAME = '{table_name}'
        """)
        table_exists = cursor.fetchone()[0] > 0
        
        if table_exists:
            # Якщо таблиця існує, видаляємо її
            cursor.execute(f"DROP TABLE [{table_name}]")
            conn.commit()
        
        # Створюємо SQL для створення таблиці
        column_definitions = []
        for col_info in columns_info:
            col_name = col_info[0]
            col_type = col_info[1]
            col_length = col_info[2]
            is_nullable = col_info[3]
            
            # Формуємо тип даних
            if col_type in ['varchar', 'nvarchar', 'char', 'nchar']:
                if col_length:
                    type_def = f"{col_type}({col_length})"
                else:
                    type_def = f"{col_type}(MAX)"
            elif col_type in ['decimal', 'numeric']:
                type_def = col_type
            else:
                type_def = col_type
            
            nullable = "NULL" if is_nullable == "YES" else "NOT NULL"
            column_definitions.append(f"[{col_name}] {type_def} {nullable}")
        
        create_table_sql = f"""
            CREATE TABLE [{table_name}] (
                {', '.join(column_definitions)}
            )
        """
        
        cursor.execute(create_table_sql)
        conn.commit()
        
        # Копіюємо дані з Material в нову таблицю
        column_names = [col[0] for col in columns_info]
        columns_str = ', '.join([f"[{col}]" for col in column_names])
        
        insert_sql = f"""
            INSERT INTO [{table_name}] ({columns_str})
            SELECT {columns_str}
            FROM Material
        """
        
        cursor.execute(insert_sql)
        rows_loaded = cursor.rowcount
        conn.commit()
        conn.close()
        
        return jsonify({
            'success': True,
            'message': f'Дані успішно завантажено в таблицю "{table_name}"',
            'table_name': table_name,
            'rows_loaded': rows_loaded
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/upload-excel', methods=['POST'])
def upload_excel():
    """Завантажити дані з Excel файлу в нову таблицю"""
    try:
        # Перевіряємо, чи є файл у запиті
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'Файл не знайдено'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'Файл не вибрано'}), 400
        
        # Отримуємо ім'я таблиці з параметрів запиту або з імені файлу
        table_name = request.args.get('table_name', '')
        if not table_name:
            # Якщо ім'я не передано, використовуємо ім'я файлу без розширення
            table_name = os.path.splitext(file.filename)[0]
        
        # Очищаємо ім'я таблиці від недопустимих символів
        table_name = re.sub(r'[^a-zA-Z0-9_]', '_', table_name)
        if not table_name or table_name[0].isdigit():
            table_name = 'table_' + table_name
        
        # Читаємо Excel файл
        try:
            df = pd.read_excel(file)
        except Exception as e:
            return jsonify({'success': False, 'error': f'Помилка читання Excel файлу: {str(e)}'}), 400
        
        if df.empty:
            return jsonify({'success': False, 'error': 'Файл порожній'}), 400
        
        # Очищаємо назви колонок від недопустимих символів для SQL Server
        df.columns = [re.sub(r'[^a-zA-Z0-9_]', '_', str(col)) for col in df.columns]
        # Замінюємо порожні назви колонок
        df.columns = [f'Column_{i+1}' if not col or col.startswith('_') else col for i, col in enumerate(df.columns)]
        
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Перевіряємо, чи існує вже таблиця з таким ім'ям
        cursor.execute(f"""
            SELECT COUNT(*) 
            FROM INFORMATION_SCHEMA.TABLES 
            WHERE TABLE_NAME = '{table_name}'
        """)
        table_exists = cursor.fetchone()[0] > 0
        
        if table_exists:
            # Якщо таблиця існує, видаляємо її
            cursor.execute(f"DROP TABLE [{table_name}]")
            conn.commit()
        
        # Визначаємо типи даних для колонок
        column_definitions = []
        for col in df.columns:
            # Визначаємо тип даних на основі перших непустих значень
            col_data = df[col].dropna()
            if len(col_data) == 0:
                sql_type = 'NVARCHAR(MAX)'
            elif pd.api.types.is_integer_dtype(df[col]):
                sql_type = 'BIGINT'
            elif pd.api.types.is_float_dtype(df[col]):
                sql_type = 'FLOAT'
            elif pd.api.types.is_datetime64_any_dtype(df[col]):
                sql_type = 'DATETIME'
            elif pd.api.types.is_bool_dtype(df[col]):
                sql_type = 'BIT'
            else:
                # Для текстових даних визначаємо максимальну довжину
                max_len = df[col].astype(str).str.len().max()
                if max_len > 4000:
                    sql_type = 'NVARCHAR(MAX)'
                else:
                    sql_type = f'NVARCHAR({min(max_len * 2, 4000)})'
            
            column_definitions.append(f"[{col}] {sql_type} NULL")
        
        # Створюємо таблицю
        create_table_sql = f"""
            CREATE TABLE [{table_name}] (
                {', '.join(column_definitions)}
            )
        """
        
        cursor.execute(create_table_sql)
        conn.commit()
        
        # Вставляємо дані
        rows_loaded = 0
        for index, row in df.iterrows():
            try:
                # Підготовка значень для вставки
                values = []
                placeholders = []
                for col in df.columns:
                    value = row[col]
                    if pd.isna(value):
                        values.append(None)
                    elif pd.api.types.is_datetime64_any_dtype(df[col]):
                        # Конвертуємо datetime в Python datetime для pyodbc
                        if pd.notna(value):
                            if isinstance(value, pd.Timestamp):
                                values.append(value.to_pydatetime())
                            else:
                                values.append(value)
                        else:
                            values.append(None)
                    elif pd.api.types.is_bool_dtype(df[col]):
                        values.append(bool(value) if pd.notna(value) else None)
                    elif pd.api.types.is_integer_dtype(df[col]):
                        values.append(int(value) if pd.notna(value) else None)
                    elif pd.api.types.is_float_dtype(df[col]):
                        values.append(float(value) if pd.notna(value) else None)
                    else:
                        values.append(str(value) if pd.notna(value) else None)
                    placeholders.append('?')
                
                insert_sql = f"""
                    INSERT INTO [{table_name}] ({', '.join([f'[{col}]' for col in df.columns])})
                    VALUES ({', '.join(placeholders)})
                """
                
                cursor.execute(insert_sql, values)
                rows_loaded += 1
            except Exception as e: 
                continue
        
        conn.commit()
        conn.close()
        
        return jsonify({
            'success': True,
            'message': f'Дані успішно завантажено в таблицю "{table_name}"',
            'table_name': table_name,
            'rows_loaded': rows_loaded,
            'columns_count': len(df.columns)
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

