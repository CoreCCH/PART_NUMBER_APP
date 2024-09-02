import mysql.connector
import numpy as np

__host= '192.168.3.113' # 'localhost'192.168.3.196
__user= 'user'
__port= '3306'
__password= '50778700'
__SQLName = ('name_rule',)


# def get_table_data(table_name, database_name= __SQLName[0],host=__host, port=__port, user=__user, password=__password):
#     try:
#         connection = mysql.connector.connect(host=host,port=port,user=user,password=password,database=database_name)
#         cursor = connection.cursor()
#         query = f"SELECT * FROM `{table_name}`"
#         cursor.execute(query)
#         records = cursor.fetchall()

#     except mysql.connector.Error as err:
#         print(f"Error: {err}")
#         return (f"Error: {err}")

#     else:
#         cursor.close()
#         connection.close()

#         return records

def get_table_data(table_name, 
                   options=None, 
                   order_by=None, 
                   order='ASC', 
                   database_name=__SQLName[0], 
                   host=__host, 
                   port=__port, 
                   user=__user, 
                   password=__password):
    """
    Fetch data from a specified table with optional filtering and sorting.
    
    :param table_name: Name of the table to fetch data from.
    :param options: A dictionary of column-value pairs for filtering (e.g., {'age': 30}).
    :param order_by: The column name to order by.
    :param order: The sorting order, 'ASC' for ascending or 'DESC' for descending.
    :param database_name: Name of the database.
    :param host: Host of the database server.
    :param port: Port number of the database server.
    :param user: Username for the database.
    :param password: Password for the database.
    :return: Fetched records or an error message.
    """
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # Start forming the query
        query = f"SELECT * FROM `{table_name}`"

        # Add filtering options if provided
        if options:
            conditions = " AND ".join([f"{key}='{value}'" for key, value in options.items()])
            query += f" WHERE {conditions}"

        # Add ordering if provided
        if order_by:
            query += f" ORDER BY `{order_by}` {order}"

        cursor.execute(query)
        records = cursor.fetchall()

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        cursor.close()
        connection.close()
        return records
    
def fetch_data_from_item_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `item_list` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            base_query += f" AND `{key}` = %s"
            params.append(value)

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(base_query, params)
        result = cursor.fetchall()
        # for row in result:
            # print(row)

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result    

def fetch_data_from_category_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `category_list` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            base_query += f" AND `{key}` = %s"
            params.append(value)

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(base_query, params)
        result = cursor.fetchall()
        # for row in result:
            # print(row)

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result    

def fetch_data_from_type_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `type_list` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            base_query += f" AND `{key}` = %s"
            params.append(value)

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(base_query, params)
        result = cursor.fetchall()
        # for row in result:
            # print(row)

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result    

def fetch_data_from_col_module_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `col_module` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            base_query += f" AND `{key}` = %s"
            params.append(value)

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(base_query, params)
        result = cursor.fetchall()
        # for row in result:
            # print(row)

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result   

def fetch_data_from_table(filters: dict, table_name: str, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = f"SELECT * FROM `{table_name}` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            base_query += f" AND `{key}` = %s"
            params.append(value)

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(base_query, params)
        result = cursor.fetchall()
        # for row in result:
            # print(row)

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result   

def insert_or_update_table_data(table_name, columns, data, unique_key, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    if not data:
        print("No data to insert or update.")
        return
    
    columns_str = ", ".join([f"`{col}`" for col in columns])
    placeholders_str = ", ".join(["%s"] * len(columns))
    update_clause = ", ".join([f"`{col}` = VALUES(`{col}`)" for col in columns if col != unique_key])
    
    query = f"""
    INSERT INTO `{table_name}` ({columns_str}) 
    VALUES ({placeholders_str}) 
    ON DUPLICATE KEY UPDATE {update_clause}
    """

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        for row in data:
            params = tuple(row)
            cursor.execute(query, params)
        
        connection.commit()
    
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        if connection:
            cursor.close()
            connection.close()
        return True

def get_column_names(table_name, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    query = """
    SELECT COLUMN_NAME
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_SCHEMA = %s
      AND TABLE_NAME = %s;
    """
    
    try:
        # 连接到数据库
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(query, (database_name, table_name))
        
        # 获取列名
        columns = cursor.fetchall()
        column_names = [col[0] for col in columns]
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return []

    else:
        if connection:
            cursor.close()
            connection.close()

        return column_names

def delete_by_keys(table_name, key_column, key_values, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # 创建 IN 子句所需的参数占位符
        format_strings = ','.join(['%s'] * len(key_values))
        query = f"DELETE FROM `{table_name}` WHERE `{key_column}` IN ({format_strings})"
        cursor.execute(query, tuple(key_values))
        
        connection.commit()
        # print(f"Records with {key_column} in {key_values} deleted successfully from {table_name}")
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"
    
    else:
        cursor.close()
        connection.close()
        return True

def delete_by_conditions(table_name, conditions, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # Build the WHERE clause based on the conditions provided
        query = f"DELETE FROM `{table_name}` WHERE "
        params = []

        condition_clauses = []
        for column, value in conditions.items():
            condition_clauses.append(f"`{column}` = %s")
            params.append(value)

        query += " AND ".join(condition_clauses)

        # Execute the query
        cursor.execute(query, tuple(params))
        
        connection.commit()
        print(f"Records with conditions {conditions} deleted successfully from {table_name}")
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"
    
    else:
        cursor.close()
        connection.close()
        return True

def drop_table(table_names, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        with connection.cursor() as cursor:
            for table_name in table_names:
                sql = f"DROP TABLE IF EXISTS {table_name}"
                cursor.execute(sql)
            connection.commit()
    
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"
    
    else:
        cursor.close()
        connection.close()

        return True
 
def add_data_to_col_module_table(col_module_info: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    col_module_info = [None if isinstance(item, float) and np.isnan(item) else item for item in col_module_info]

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # 構建插入語句
        query = """
        INSERT INTO `col_module`(`Num1`, `Num2`, 
        `Num3`, `Num4`, `Num5`, `Num6`, 
        `Num7`, `Num8`, `Num9`, `Num10`, 
        `Num11`, `Count`) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        # 執行插入查詢
        cursor.execute(query, tuple(col_module_info))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True

# print(fetch_data_from_item_table({}))
# print(fetch_data_from_table({"NUM1":"SMT", "NUM2":"橋堆"}, 'col_module'))
