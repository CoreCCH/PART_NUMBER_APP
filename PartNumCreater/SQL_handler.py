import mysql.connector
import numpy as np
from datetime import datetime

__host= '192.168.3.113' # ''192.168.3.196
__user= 'user'
__port= '3306'
__password= '50778700'
__SQLName = ('test_pcb',)

def get_PCBmanagement(host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(host=host,port=port,user=user,password=password)
        cursor = connection.cursor()
        cursor.execute("SHOW DATABASES;")
        records = cursor.fetchall()
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"
    
    else:
        cursor.close()
        connection.close()

        if __SQLName in records:
            return __SQLName[0]
        else:
            return "SQL does not create"

def get_table_data(table_name, database_name= __SQLName[0],host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(host=host,port=port,user=user,password=password,database=database_name)
        cursor = connection.cursor()
        query = f"SELECT * FROM `{table_name}`"
        cursor.execute(query)
        records = cursor.fetchall()

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return (f"Error: {err}")

    else:
        cursor.close()
        connection.close()

        return records

def add_data_to_Employees_table(user_info: list ,database_name= __SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(host=host,port=port,user=user,password=password,database=database_name)
        cursor = connection.cursor()
        query = "INSERT INTO `employees` (EmployeeName, EmployeePassWord, Email) VALUES (%s, %s, %s)"
        cursor.execute(query, (user_info[0], user_info[1], user_info[2]))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return (f"Error: {err}")

    else:
        cursor.close()
        connection.commit()
        connection.close()

        return True
    
def add_data_to_folder_table(foler_name: str ,database_name= __SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(host=host,port=port,user=user,password=password,database=database_name)
        cursor = connection.cursor()
        query = "INSERT INTO `folder_list` (Folder_Name) VALUES (%s)"
        cursor.execute(query, (foler_name,))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return (f"Error: {err}")

    else:
        cursor.close()
        connection.commit()
        connection.close()

        return True
    
def add_data_to_folder_BOM_table(Folder_BOM_list: list ,database_name= __SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(host=host,port=port,user=user,password=password,database=database_name)
        cursor = connection.cursor()
        query = "INSERT INTO `folder_correspond_bom` (Folder_Name, BOM_Name) VALUES (%s, %s)"
        cursor.execute(query, (Folder_BOM_list[0], Folder_BOM_list[1]))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return (f"Error: {err}")

    else:
        cursor.close()
        connection.commit()
        connection.close()

        return True

def add_data_to_manage_table(Manage_list: list, database_name= __SQLName[0], host=__host, port=__port, user=__user, password=__password):
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
        INSERT INTO `manage` (
            UniCode, Motion, Quantity, SourceLocation, DestinationLocation, PickingList, WorkOrder, UpdateTime, EditBy 
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        # 執行插入查詢
        cursor.execute(query, tuple(Manage_list))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True

def validate_user_login(user_info: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # 構建查詢語句
        query = "SELECT * FROM `employees` WHERE `EmployeeName` = %s AND `EmployeePassWord` = %s"
        
        # 執行 SELECT 查詢
        cursor.execute(query, (user_info[0], user_info[1]))

        # 獲取查詢結果
        result = cursor.fetchone()

        

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return (f"Error: {err}")

    else:
        cursor.close()
        connection.close()

        if result:
            return True
        else:
            return False

def add_data_to_Material_table(material_info: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    material_info = [None if isinstance(item, float) and np.isnan(item) else item for item in material_info]

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
        INSERT INTO `Material` (
            ERP_Code, ECount_Code, Part_Name, Item, Type, Size_Type, 
            Percentage_Package, Capacitance_Resistance_Name, 
            Voltage_PinSize_Frequency, Manufacturer, Supplier, 
            CreatedAt, PartNumber, EditBy
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        # 執行插入查詢
        cursor.execute(query, tuple(material_info))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True
    
def add_data_to_Inventory_table(inventory_info: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
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
        INSERT INTO `Inventory` (
            InventoryID, ERP_Code, TotalQuantity, Location, EntryDate, EditBy
        ) VALUES (%s, %s, %s, %s, %s, %s)
        """
        
        # 執行插入查詢
        cursor.execute(query, tuple(inventory_info))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True

def add_data_to_Boxes_table(box_info: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
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
        INSERT INTO `Boxes` (
            BoxID, InventoryID, Serial_Number, Quantity, UniCode, Supplier_Batch, In_Time, Print_Count
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """
        
        # 執行插入查詢
        cursor.execute(query, tuple(box_info))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True

def add_data_to_Operation_table(operation_info: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
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
        INSERT INTO `Operation` (
            BoxID, Motion, Quantity_Old, Quantity_New, Old_BoxID, LastUpdated, EditBy
        ) VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        
        # 執行插入查詢
        cursor.execute(query, tuple(operation_info))

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True

def add_data_to_Forecast_Group_table(group_list: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
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
        INSERT INTO `forecast_list` (
            Group_Name, BOM_Name, Quantity, Date, EditBy
        ) VALUES (%s, %s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE 
            Quantity = VALUES(Quantity)
        """
        
        # 執行插入查詢，使用 executemany 插入多筆資料
        cursor.executemany(query, group_list)

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True

def fetch_data_from_Employees_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `Employees` WHERE 1=1"
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
        for row in result:
            pass

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()

    return result

def fetch_data_from_Material_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `Material` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            base_query += f" AND `{key}` LIKE %s"
            params.append(f"%{value}%")

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
        for row in result:
            print(row)

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result

def fetch_data_from_Inventory_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `Inventory` WHERE 1=1"
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

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result

def fetch_data_from_Boxes_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `Boxes` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            if isinstance(value, (list, tuple)):  # 檢查是否為列表或元組
                if value:  # 確保列表或元組不為空
                    placeholders = ','.join(['%s'] * len(value))  # 創建佔位符
                    base_query += f" AND `{key}` IN ({placeholders})"
                    params.extend(value)  # 展開列表並添加到參數中
            else:
                base_query += f" AND `{key}` = %s"
                params.append(value)

    base_query += " ORDER BY `In_Time` DESC"

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

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result

def fetch_data_from_Operation_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `Operation` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            if key == "LastUpdated":
                base_query += " AND DATE(`LastUpdated`) = %s"
                params.append(value)
            elif isinstance(value, (list, tuple)):  # 如果值是列表或元组
                placeholders = ','.join(['%s'] * len(value))  # 为IN语句生成占位符
                base_query += f" AND `{key}` IN ({placeholders})"
                params.extend(value)  # 扩展列表并添加到参数中
            else:
                base_query += f" AND `{key}` = %s"
                params.append(value)
        
    base_query += " ORDER BY `LastUpdated` DESC"

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

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result

def fetch_data_from_Manage_table(filters: dict, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    base_query = "SELECT * FROM `Manage` WHERE 1=1"
    params = []

    for key, value in filters.items():
        if value != "":
            if key == "ERPCode":
                # 如果是ERPCode，則僅比較前11個字符
                base_query += f" AND LEFT(`{"UniCode"}`, 11) = %s"
                params.append(value[:11])  # 取值的前11個字符作為參數
            elif key == "UniCode":
                # 如果是UniCode，則完全匹配
                base_query += f" AND `{key}` = %s"
                params.append(value)
            elif key == "UpdateTime":
                base_query += " AND DATE(`UpdateTime`) = %s"
                params.append(value)
            elif isinstance(value, (list, tuple)):  # 如果值是列表或元组
                placeholders = ','.join(['%s'] * len(value))  # 为IN语句生成占位符
                base_query += f" AND `{key}` IN ({placeholders})"
                params.extend(value)  # 扩展列表并添加到参数中
            else:
                base_query += f" AND `{key}` = %s"
                params.append(value)
        
    base_query += " ORDER BY `UpdateTime` DESC"

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

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

    return result

def update_total_quantity_if_duplicate(inventory_id, total_quantity, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    # If inventory id is in Inventory table, add total quantity.
    select_query = "SELECT TotalQuantity FROM Inventory WHERE InventoryID = %s"
    update_query = "UPDATE Inventory SET TotalQuantity = TotalQuantity + %s WHERE InventoryID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        
        # Check for duplicates
        cursor.execute(select_query, (inventory_id,))
        result = cursor.fetchone()

        if result:
            # If duplicate exists, update the total quantity
            cursor.execute(update_query, (total_quantity, inventory_id))
            connection.commit()
            print(f"TotalQuantity updated for InventoryID: {inventory_id}")
        else:
            print(f"No record found for InventoryID: {inventory_id}")
            return f"No record found for InventoryID: {inventory_id}"
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()
        return True

def update_total_quantity_by_inventory_id(inventory_id, total_quantity, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    # If inventory id is in Inventory table, add total quantity.
    select_query = "SELECT TotalQuantity FROM Inventory WHERE InventoryID = %s"
    update_query = "UPDATE Inventory SET TotalQuantity = %s WHERE InventoryID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        
        # Check for duplicates
        cursor.execute(select_query, (inventory_id,))
        result = cursor.fetchone()

        if result:
            # If duplicate exists, update the total quantity
            cursor.execute(update_query, (total_quantity, inventory_id))
            connection.commit()
            print(f"TotalQuantity updated for InventoryID: {inventory_id}")
        else:
            print(f"No record found for InventoryID: {inventory_id}")
            return f"No record found for InventoryID: {inventory_id}"
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()
        return True

def increment_print_count_by_box_id(box_id,  database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    select_query = "SELECT Print_Count FROM Boxes WHERE BoxID = %s"
    update_query = "UPDATE Boxes SET Print_Count = Print_Count + 1 WHERE BoxID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # Fetch the current Print_Count
        cursor.execute(select_query, (box_id,))
        result = cursor.fetchone()

        if result is None:
            print(f"No record found with BoxID {box_id}")
            return f"No record found with BoxID {box_id}"

        current_print_count = result[0]
        print(f"Current Print_Count for BoxID {box_id}: {current_print_count}")

        # Update the Print_Count
        cursor.execute(update_query, (box_id,))
        connection.commit()
        print(f"Print_Count for BoxID {box_id} has been incremented by 1")

        return f"Print_Count for BoxID {box_id} has been incremented by 1"

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()

def get_max_serial_number(inventory_id, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    # Give inventory id and return max serial number
    query = "SELECT MAX(Serial_Number) FROM Boxes WHERE InventoryID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(query, (inventory_id,))
        result = cursor.fetchone()

        max_serial_number = (result[0]+1) if result[0] is not None else 1
        print(f"Max Serial Number for InventoryID {inventory_id}: {max_serial_number}")
        return max_serial_number

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()

def get_total_quantity_by_inventory_id(inventory_id, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    query = "SELECT SUM(Quantity) FROM Boxes WHERE InventoryID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        cursor.execute(query, (inventory_id,))
        result = cursor.fetchone()

        total_quantity = result[0] if result[0] is not None else 0
        print(f"Total Quantity for InventoryID {inventory_id}: {total_quantity}")
        return total_quantity

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()

def update_quantity_by_box_id(box_id, new_quantity, time = datetime.now(),database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    update_query = "UPDATE Boxes SET Quantity = %s, In_Time = %s WHERE BoxID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # Execute the update query
        cursor.execute(update_query, (new_quantity, time, box_id))
        connection.commit()

        if cursor.rowcount == 0:
            print(f"No record found with BoxID {box_id}")
            return f"No record found with BoxID {box_id}"
        else:
            print(f"Quantity for BoxID {box_id} has been updated to {new_quantity} and In_Time has been updated to {time}")
            

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()
        return True

def update_Ecount_by_ERP_Code(ERP_Code, Ecount_Code, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    # If inventory id is in Inventory table, add total quantity.
    select_query = "SELECT Ecount_Code FROM Material WHERE ERP_Code = %s"
    update_query = "UPDATE Material SET Ecount_Code = %s WHERE ERP_Code = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        
        # Check for duplicates
        cursor.execute(select_query, (ERP_Code,))
        result = cursor.fetchone()

        if result:
            # If duplicate exists, update the total quantity
            cursor.execute(update_query, (Ecount_Code, ERP_Code))
            connection.commit()
            print(f"Eount_Code updated for ERP_Code: {ERP_Code}")
        else:
            print(f"No record found for ERP_Code: {ERP_Code}")
            return f"No record found for ERP_Code: {ERP_Code}"
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()
        return True
    
def update_PN_by_ERP_Code(ERP_Code, PN, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    # If inventory id is in Inventory table, add total quantity.
    select_query = "SELECT PartNumber FROM Material WHERE ERP_Code = %s"
    update_query = "UPDATE Material SET PartNumber = %s WHERE ERP_Code = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        
        # Check for duplicates
        cursor.execute(select_query, (ERP_Code,))
        result = cursor.fetchone()

        if result:
            # If duplicate exists, update the total quantity
            cursor.execute(update_query, (PN, ERP_Code))
            connection.commit()
            print(f"Eount_Code updated for ERP_Code: {ERP_Code}")
        else:
            print(f"No record found for ERP_Code: {ERP_Code}")
            return f"No record found for ERP_Code: {ERP_Code}"
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()
        return True

def update_UniCode_by_BoxID(BoxID,UniCode, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    select_query = "SELECT UniCode FROM boxes WHERE BoxID = %s"
    update_query = "UPDATE boxes SET UniCode = %s WHERE BoxID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        
        # Check for duplicates
        cursor.execute(select_query, (BoxID,))
        result = cursor.fetchone()

        if result:
            # If duplicate exists, update the total quantity
            cursor.execute(update_query, (UniCode, BoxID))
            connection.commit()
            print(f"Eount_Code updated for ERP_Code: {BoxID}")
        else:
            print(f"No record found for ERP_Code: {BoxID}")
            return f"No record found for ERP_Code: {BoxID}"
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()
        return True
    
def update_Time_by_BoxID(BoxID,Time, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    select_query = "SELECT In_Time FROM boxes WHERE BoxID = %s"
    update_query = "UPDATE boxes SET In_Time = %s WHERE BoxID = %s"

    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()
        
        # Check for duplicates
        cursor.execute(select_query, (BoxID,))
        result = cursor.fetchone()

        if result:
            # If duplicate exists, update the total quantity
            cursor.execute(update_query, (Time, BoxID))
            connection.commit()
            print(f"Eount_Code updated for ERP_Code: {BoxID}")
        else:
            print(f"No record found for ERP_Code: {BoxID}")
            return f"No record found for ERP_Code: {BoxID}"
        
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()
        return True

def update_Location(database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    # 連接到資料庫
    connection = mysql.connector.connect(
        host=host,
        port=port,
        user=user,
        password=password,
        database=database_name
    )

    try:
        with connection.cursor() as cursor:
            # 更新 SourceLocation 欄位
            sql_update_source = """
            UPDATE manage
            SET SourceLocation = '台南半成品倉 (貼片加工)'
            WHERE SourceLocation = '台南半成品倉 (貼片';
            """
            cursor.execute(sql_update_source)
            
            # 更新 DestinationLocation 欄位
            sql_update_destination = """
            UPDATE manage
            SET DestinationLocation = '台南半成品倉 (貼片加工)'
            WHERE DestinationLocation = '台南半成品倉 (貼片';
            """
            cursor.execute(sql_update_destination)

            # 提交更改
            connection.commit()

    finally:
        # 關閉資料庫連接
        connection.close()

def create_and_populate_table(table_name, data_list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    # Define the create table query
    create_table_query = f"""
    CREATE TABLE IF NOT EXISTS `{table_name}` (
        ERP_Code CHAR(11) UNIQUE,
        Quantity INT,
        PM INT AUTO_INCREMENT PRIMARY KEY
    );
    """
    # Establish connection to the MySQL database
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # Create the table
        cursor.execute(create_table_query)

        # Insert data into the table
        insert_query = f"INSERT INTO `{table_name}` (ERP_Code, Quantity) VALUES (%s, %s)"
        for record in data_list:
            for erp_code, quantity in record.items():
                cursor.execute(insert_query, (erp_code, quantity))
        
        # Commit the transaction
        connection.commit()
        
        print(f"Table '{table_name}' created and data inserted successfully.")

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"
    else:
        if connection.is_connected():
            cursor.close()
            connection.close()
        return True

def list_tables(database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        # Establish connection to the MySQL database
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # Execute the query to list all tables
        cursor.execute("SHOW TABLES")

        # Fetch all results
        tables = cursor.fetchall()

        # Print all table names
        
        return [table[0] for table in tables]

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

def add_data_to_Bom_Operation_table(operation_info: list, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
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
        INSERT INTO `operation` (
            Motion, Type, Name, Op_Time, EditBy
        ) VALUES (%s, %s, %s, %s, %s)
        """
        
        # 執行插入查詢
        if all(isinstance(row, (list, tuple)) and len(row) == 5 for row in operation_info):
            # 执行批量插入查询
            cursor.executemany(query, operation_info)
            connection.commit()  # 提交所有更改
        else:
            raise ValueError("operation_info must be a list of tuples or lists with 5 elements each.")
    

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        cursor.close()
        connection.close()

        return True

def update_quantity_in_Boxes_table(BoxID, new_quantity, database_name=__SQLName[0], host=__host, port=__port, user=__user, password=__password):
    try:
        connection = mysql.connector.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database_name
        )
        cursor = connection.cursor()

        # 構建更新語句
        query = """
        UPDATE `Boxes`
        SET Quantity = %s
        WHERE BoxID = %s
        """
        
        # 執行更新查詢
        cursor.execute(query, (new_quantity, BoxID))

        # 檢查是否有行被更新
        if cursor.rowcount == 0:
            return f"Error: BoxID {BoxID} not found."

    except mysql.connector.Error as err:
        print(f"Error: {err}")
        return f"Error: {err}"

    else:
        connection.commit()
        cursor.close()
        connection.close()

        return True

# 測試用例
# operation_info = [
#     'BOX12345678901234567890', 'IN', 30, 40, '2024-07-19 10:30:00', None
# ]
# print(add_data_to_Operation_table(operation_info))

# box_info = [
#     'BOX12345678901234567890', 'INV123456789012345678', 12345, 50, 5
# ]
# print(add_data_to_Boxes_table(box_info))

# inventory_info = [
#     'INV1234567890123456', 'ERP12345678', 100, 'Warehouse A', '2024-07-19 10:00:00', None
# ]
# print(add_data_to_Inventory_table(inventory_info))

# 測試用例中的 filters 字典可以包含需要過濾的欄位和值
# filters = {
#     'ERP_Code': '',
#     'Item': 'SMT',  # 如果不想過濾 'Item'，可以設為空字符串
#     'Type': '',
#     'Size_Type': '',
#     'Percentage_Package': '',
#     'Capacitance_Resistance_Name': '',
#     'Voltage_PinSize_Frequency': '',
#     'Manufacturer': '',
#     'Supplier': '',
#     'PartNumber': '',
#     'EditBy': ''
# }

# print(fetch_data_from_Material_table(filters))

# 測試用戶登錄驗證
# print(validate_user_login(['James', '50778700']))

# material_info = [
#     'ERP23456789', 'SMT', 'Capacitor', '0805', '5%', '10uF', 
#     '16V', 'ABC Corp', 'XYZ Supplier', 'Part123', None
# ]

# print(add_data_to_Material_table(material_info))

# print(add_data_to_Employees_table(['James', '50778700']))
# get_table_data('Boxes')
# filters = {'Email': "james.chiu@orient-suntech.com"}
# print(fetch_data_from_Employees_table(filters=filters))
# filters = {"EntryDate": "2024-07-20"}
# fetch_data_from_Inventory_table(filters)
# print(get_table_data('Material'))


