import SQL_handler
import Name_Rule_SQL_handler
import pandas as pd
from component import get_place_from_code, find_stockroom_name

def load_data_from_excel(file_path: str, sheet_name="Sheet1"):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name)

    for index, row in df.iterrows():
        material_info = [
            row['ERP Code'], row['品項編號'], row['品項名稱'], row['項目'], row['種類'],
            row['尺寸/種類'], row['%數/封裝'], row['容值/阻值/名稱'], row['電壓/腳位大小/頻率'],
            row['廠商'], row['供應商'], row['產生時間'], row['Part Number'], 1
        ]
        result = SQL_handler.add_data_to_Material_table(material_info)
        print(f"Inserted row {index + 1}: {result}")

def load_Inventory_data_from_excel(file_path: str, sheet_name="Sheet1"):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name)

    for index, row in df.iterrows():
        material_info = [
            row['InventoryID'], row['ERP_code'], row['TotalQuantity'], row['Location'],
            row['EntryDate'],1
        ]
        result = SQL_handler.add_data_to_Inventory_table(material_info)
        print(f"Inserted row {index + 1}: {row['ERP_code']}: {result}")

def Update_Boxes_data_from_excel(file_path: str, sheet_name="Sheet1"):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name)

    for index, row in df.iterrows():
        # material_info = [
        #     row['BoxID'], row['InventoryID'], row['Number'], row['Quantity'],'', '','',
        #     row['Print_Count']
        # ]
        SQL_handler.update_quantity_in_Boxes_table(BoxID=row['BoxID'], new_quantity=row['Quantity'])

def load_Boxes_data_from_excel(file_path: str, sheet_name="Sheet1"):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name)

    for index, row in df.iterrows():
        material_info = [
            row['BoxID'], row['InventoryID'], row['Number'], row['Quantity'],'', '','',
            row['Print_Count']
        ]
        result = SQL_handler.add_data_to_Boxes_table(material_info)
        print(f"Inserted row {index + 1}: {row['InventoryID']}: {result}")

def Update_Boxes_data_from_excel(file_path: str, sheet_name="Sheet1"):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name)

    for index, row in df.iterrows():
        # material_info = [
        #     row['BoxID'], row['InventoryID'], row['Number'], row['Quantity'],'', '','',
        #     row['Print_Count']
        # ]
        SQL_handler.update_quantity_in_Boxes_table(BoxID=row['BoxID'], new_quantity=row['Quantity'])

def load_Operation_data_from_excel(file_path: str, sheet_name="Sheet1"):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name)

    for index, row in df.iterrows():
        if row['Status'] == "入庫":
            row['Status'] = 'IN'
        elif row['Status'] == "出庫":
            row['Status'] = 'OUT'
        elif row['Status'] == "退庫":
            row['Status'] = 'RETURNED'
        elif row['Status'] == "轉倉":
            row['Status'] = 'SWITCH'
        elif row['Status'] == "刪除":
            row['Status'] = 'DELETE'
        material_info = [
            row['BoxID'], row['Status'], row['Quantity_Old'], 
            row['Quantity_New'], '', row['LastUpdated'], 1
        ]
        
        result = SQL_handler.add_data_to_Operation_table(material_info)
        print(f"Inserted row {index + 1}: {result}")

def get_BoxID_had_print_count():
    item_data = []

    items = (Name_Rule_SQL_handler.get_table_data(table_name="boxes", database_name="pcbmanagement"))
    item_list = [item[0] for item in items if item[0][-3:-1] == "55"]
    item_old_list = [item[0] for item in items if item[0][-3:-1] == "51"]

    
    # for data in item_list:
    #     for data_old in item_old_list:
    #         if data[:11] == data_old[:11]:
    #             item_data.append([data, data[:12] + data_old[12:18] + data[18:]])
                

    # for idx, val in enumerate(item_data):
    #     # print(val[0],val[1][:-3] + "51" + val[1][-1])
    #     SQL_handler.update_UniCode_by_BoxID(val[0],val[1][:-3] + "51" + val[1][-1],database_name="pcbmanagement")

    # box_list = [item[0][:-3] + "51" + item[0][-1] for item in items if item[0][-3:-1] == "52"]
    # print(box_list)
    # for idx, val in enumerate(item_list):
    #     SQL_handler.update_UniCode_by_BoxID(val,box_list[idx],database_name="pcbmanagement")

    for idx, val in enumerate(item_list):
        SQL_handler.update_UniCode_by_BoxID(val,val,database_name="pcbmanagement")

def old_operation_to_new():
    old_operation_data = SQL_handler.get_table_data(table_name="operation", database_name="pcbmanagement")
    idx = 0
    while(1):
        if (idx == len(old_operation_data)-1):
            break
    # for idx in range(len(old_operation_data)-1):
        box_id = old_operation_data[idx][1]
        motion = old_operation_data[idx][2]
        unicode = SQL_handler.fetch_data_from_Boxes_table({"BoxID":box_id})[0][4]
        update_time = old_operation_data[idx][6]
        edit_by = old_operation_data[idx][7]
        work_order = None
        pickingL_list = None

        # get stock
        from barcode_generator import placecode, stockroom
        if(get_place_from_code(int(box_id[17]),placecode) != None):
            place = get_place_from_code(int(box_id[17]),placecode)
            if(find_stockroom_name(place ,int(box_id[18]), stockroom) != None):
                destination_stcok_room = find_stockroom_name(place ,int(box_id[18]), stockroom)

        # merge in + out to switch
        quantity = abs(old_operation_data[idx][3] - old_operation_data[idx][4])
        if ((motion == "IN" and old_operation_data[idx+1][2] == "OUT") and 
            unicode == SQL_handler.fetch_data_from_Boxes_table({"BoxID":old_operation_data[idx+1][1]})[0][4] and
            quantity == abs(old_operation_data[idx+1][3] - old_operation_data[idx+1][4])):
            # recognize that pair of in and out is switch
            source_box_id = old_operation_data[idx+1][1]
            if(get_place_from_code(int(source_box_id[17]),placecode) != None):
                place = get_place_from_code(int(source_box_id[17]),placecode)
                if(find_stockroom_name(place ,int(source_box_id[18]), stockroom) != None):
                    source_stcok_room = find_stockroom_name(place ,int(source_box_id[18]), stockroom)
            motion = "SWITCH"
            idx=idx+1
        elif motion == "SWITCH":
            if(get_place_from_code(int(unicode[17]),placecode) != None):
                place = get_place_from_code(int(unicode[17]),placecode)
                if(find_stockroom_name(place ,int(unicode[18]), stockroom) != None):
                    source_stcok_room = find_stockroom_name(place ,int(unicode[18]), stockroom)
        elif motion == "OUT":
            source_stcok_room = destination_stcok_room
            destination_stcok_room = None
        elif motion == "DELETE":
            destination_stcok_room = None
        elif motion == "RETURNED":
            source_stcok_room = None
            for data in old_operation_data:
                if data[2] == "IN" and SQL_handler.fetch_data_from_Boxes_table({"BoxID":data[1]})[0][4] == unicode and data[1] != unicode:    
                    if(get_place_from_code(int(data[1][17]),placecode) != None):
                        place = get_place_from_code(int(data[1][17]),placecode)
                        if(find_stockroom_name(place ,int(data[1][18]), stockroom) != None):
                            source_stcok_room = find_stockroom_name(place ,int(data[1][18]), stockroom)
                    break
        else:
            source_stcok_room = None
        
        print([unicode, motion, quantity, source_stcok_room, destination_stcok_room, pickingL_list, work_order, update_time, edit_by])
        SQL_handler.add_data_to_manage_table([unicode, motion, quantity, source_stcok_room, destination_stcok_room, pickingL_list, work_order, update_time, edit_by])
        idx=idx+1
    
def update_manage_Loaction():
    SQL_handler.update_Location()
    operation_data = SQL_handler.get_table_data(table_name="manage", database_name="pcbmanagement")
    print(operation_data)
    

# def set_Time_by_BoxID():
#     items = (Name_Rule_SQL_handler.get_table_data(table_name="boxes", database_name="pcbmanagement"))
#     item_list = [item[0] for item in items if item[0][-3:-1] == "51"]
#     print(item_list)
    # from datetime import datetime
    # now = datetime.now()
    # for idx, val in enumerate(item_list):
    #     SQL_handler.update_Time_by_BoxID(val,now, database_name="pcbmanagement")
old_operation_to_new()
# update_manage_Loaction()
# Update_Boxes_data_from_excel('stock.xlsx', 'Boxes')
# set_Time_by_BoxID()
# get_BoxID_had_print_count()
# load_data_from_excel('output.xlsx')
# load_Inventory_data_from_excel('stock.xlsx', 'Inventory')
# load_Boxes_data_from_excel('stock.xlsx', 'Boxes')
# load_Operation_data_from_excel('stock.xlsx', 'Operation')