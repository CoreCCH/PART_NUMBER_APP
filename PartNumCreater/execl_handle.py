import pandas as pd
import copy


data = {}
list_data = []

def excel_file_read(file_name: str, sheet_name: str):
    # 讀取 Excel 文件
    return pd.read_excel(file_name, sheet_name)

#get KEY array
def get_key_array(data: dict):
    unique_items = set()
    # Extract and add items from 'SMT'
    for key in data.values():
        for item in key:
            unique_items.update(item.keys())

    # Convert the set to a list
    unique_list = list(unique_items)
    return unique_list

#return capacity
def add_capacity(df2):
    return_dict = {}
    X7R_list = []
    X5R_list = []
    NPO_list = []


    X7R_column_values_code = df2.iloc[:, 7]
    X5R_column_values_code = df2.iloc[:, 9]
    NPO_column_values_code = df2.iloc[:, 11]

    for index in range(X7R_column_values_code.size):
        if (X7R_column_values_code[index] != "4.編號"  and pd.notna(X7R_column_values_code[index])):
            X7R_list.append(str(X7R_column_values_code[index]))
            X5R_list.append(X5R_column_values_code[index])
            NPO_list.append(NPO_column_values_code[index])

    return_dict.update({'X7R':X7R_list})
    return_dict.update({'X5R':X5R_list})
    return_dict.update({'NPO':NPO_list})

    return return_dict

#return array
def excel_list_read(df):
    return_value = []

    column_values = df.iloc[:, 8]
    column_values_code = df.iloc[:, 7]

    for len in range(column_values.size):
        if (column_values[len] == "項目"):
            if ({column_values[len+1]:column_values_code[len+1]} not in return_value):
                return_value.append({column_values[len+1]:column_values_code[len+1]})
    
    return return_value

def excel_type_read(df):
    type_value = []
    key_list = []
    return_value = {}
    list_data = ""

    column_values = df.iloc[:, 10]
    column_values_code = df.iloc[:, 9]
    column_values_list = df.iloc[:, 8]

    for len in range(column_values.size):
        if (column_values_list[len] == "項目"):
            if(list_data != column_values_list[len+1]):
                list_data = column_values_list[len+1]
                type_value.clear()
        if (column_values[len] == "種類"):
            if ({column_values[len+1].strip():str(column_values_code[len+1])} not in type_value):
                type_value.append({column_values[len+1].strip():str(column_values_code[len+1])})
                return_value.update({list_data:copy.deepcopy(type_value)})

    key_list = get_key_array(return_value)

    return key_list, return_value

def excel_type_read_supplier(df):
    column_values = df.iloc[:, 1]
    company_dict = [company.split()[0] for company in column_values]
    return company_dict
 
def execl_size_read(df, key_list: list):
    size_value = []
    storage_list = []
    storage_list = []
    return_value = {}
    type_data = ""

    column_values = df.iloc[:, 12]
    column_values_code = df.iloc[:, 11]
    column_values_type = df.iloc[:, 10]

    for length in range(column_values.size):
        if (column_values_type[length] == "種類"):
            if(type_data != column_values_type[length+1]):
                type_data = column_values_type[length+1]

        if (column_values[length] != "零件尺寸" and column_values[length] != "種類" and pd.notna(column_values[length])):
            storage_list.append([type_data,column_values[length],length])
    
    for item in key_list:
        size_value.clear()
        for i in range(len(storage_list)):
            if (storage_list[i][0].strip() == item.strip()):
                if ({str(column_values[storage_list[i][2]]).strip():str(column_values_code[storage_list[i][2]])} not in size_value):
                    size_value.append({str(column_values[storage_list[i][2]]).strip():str(column_values_code[storage_list[i][2]])})
                    return_value.update({item:copy.deepcopy(size_value)})
        
    return return_value

def execl_percentage_read(df, key_list: list):
    percentage_value = []
    storage_list = []
    return_value = {}
    type_data = ""

    column_values = df.iloc[:, 14]
    column_values_code = df.iloc[:, 13]
    column_values_type = df.iloc[:, 10]

    for length in range(column_values.size):
        if (column_values_type[length] == "種類"):
            if(type_data != column_values_type[length+1]):
                type_data = column_values_type[length+1]

        if (column_values[length] != "%數" and pd.notna(column_values[length])):
            storage_list.append([type_data,column_values[length],length])

    for item in key_list:
        percentage_value.clear()
        for i in range(len(storage_list)):
            if (storage_list[i][0].strip() == item.strip()):
                percentage_value.append({str(column_values[storage_list[i][2]]).strip().replace('\u3000','').replace(' ',''):str(column_values_code[storage_list[i][2]])})
                return_value.update({item:copy.deepcopy(percentage_value)})

    return return_value

def execl_capacity_read(df, key_list: list):
    capacity_value = []
    storage_list = []
    return_value = {}
    type_data = ""

    column_values = df.iloc[:, 16]
    column_values_code = df.iloc[:, 15]
    column_values_type = df.iloc[:, 10]

    for length in range(column_values.size):
        if (column_values_type[length] == "種類"):
            if(type_data != column_values_type[length+1]):
                type_data = column_values_type[length+1]

        if (column_values[length] != "電容值" and column_values[length] != "電阻值" and column_values[length] != "PCB名稱" and column_values[length] != "零件名稱" and pd.notna(column_values[length])):
            storage_list.append([type_data,column_values[length],length])

    for item in key_list:
        capacity_value.clear()
        for i in range(len(storage_list)):
            if (storage_list[i][0].strip() == item.strip()):
                capacity_value.append({str(column_values[storage_list[i][2]]).strip():str(column_values_code[storage_list[i][2]])})
                return_value.update({item:copy.deepcopy(capacity_value)})

    return return_value

def execl_voltage_read(df, key_list: list):
    voltage_value = []
    storage_list = []
    return_value = {}
    type_data = ""

    column_values = df.iloc[:, 18]
    column_values_code = df.iloc[:, 17]
    column_values_type = df.iloc[:, 10]

    for length in range(column_values.size):
        if (column_values_type[length] == "種類"):
            if(type_data != column_values_type[length+1]):
                type_data = column_values_type[length+1]

        if (column_values[length] != "電壓" and pd.notna(column_values[length])):
            storage_list.append([type_data,column_values[length],length])

    for item in key_list:
        voltage_value.clear()
        for i in range(len(storage_list)):
            if (storage_list[i][0].strip() == item.strip()):
                voltage_value.append({str(column_values[storage_list[i][2]]).strip():str(column_values_code[storage_list[i][2]])})
                return_value.update({item:copy.deepcopy(voltage_value)})

    return return_value

def execl_manufacturer_read(df, key_list: list):
    manufacturer_value = []
    storage_list = []
    return_value = {}
    type_data = ""

    column_values = df.iloc[:, 20]
    column_values_code = df.iloc[:, 19]
    column_values_type = df.iloc[:, 10]

    for length in range(column_values.size):
        if (column_values_type[length] == "種類"):
            if(type_data != column_values_type[length+1]):
                type_data = column_values_type[length+1]

        if (column_values[length] != "廠商" and pd.notna(column_values[length])):
            storage_list.append([type_data,column_values[length],length])

    for item in key_list:
        manufacturer_value.clear()
        for i in range(len(storage_list)):
            if (storage_list[i][0].strip() == item.strip()):
                manufacturer_value.append({str(column_values[storage_list[i][2]]).strip():str(column_values_code[storage_list[i][2]])})
                return_value.update({item:copy.deepcopy(manufacturer_value)})

    return return_value

def execl_supplier_read(df, key_list: list):
    manufacturer_value = []
    storage_list = []
    return_value = {}
    type_data = ""

    column_values = df.iloc[:, 22]
    column_values_code = df.iloc[:, 21]
    column_values_type = df.iloc[:, 10]

    for length in range(column_values.size):
        if (column_values_type[length] == "種類"):
            if(type_data != column_values_type[length+1]):
                type_data = column_values_type[length+1]

        if (column_values[length] != "供應商" and pd.notna(column_values[length])):
            storage_list.append([type_data,column_values[length],length])

    for item in key_list:
        manufacturer_value.clear()
        for i in range(len(storage_list)):
            if (storage_list[i][0].strip() == item.strip()):
                manufacturer_value.append({str(column_values[storage_list[i][2]]).strip():str(column_values_code[storage_list[i][2]])})
                return_value.update({item:copy.deepcopy(manufacturer_value)})

    return return_value

def check_output_existing(file_path, headers):
    try:
        existing_df = pd.read_excel(file_path)

        return existing_df
    except:
        return None

def append_data_to_excel(file_path, part_number, values, headers):
    try:
        existing_df = pd.read_excel(file_path)
    except FileNotFoundError:
        existing_df = pd.DataFrame(columns=headers)


    add_row = [part_number] + values
    new_row = pd.DataFrame([add_row], columns=headers)
    
    existing_df = pd.concat([existing_df, new_row], ignore_index=True)

    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            existing_df.to_excel(writer, index=False, header=headers)
    except:
        from component import show_alert
        show_alert("儲存物料["+part_number+"]時發生錯誤\nOutput.xlxs已被使用")
        return False

    return True

def execl_stock_record(part_number, in_stock_date, count, stock_place):
    try:
        existing_df = pd.read_excel("stock_record.xlxs")
    except FileNotFoundError:
        existing_df = pd.DataFrame(columns=["物料編碼","入庫日期","數量","倉庫"])


    
# df = excel_file_read('曜璿東命名規則 20240605-2.xlsx', '命名規則')
# df2 = excel_file_read('曜璿東命名規則 20240605-2.xlsx', '電容種類規則')

# add_capacity(df2)

# print(excel_list_read(df))
# list_data, data = excel_type_read(df)
# print(list_data)
# print(execl_size_read(df, list_data))
# print(execl_percentage_read(df, list_data))
# print(execl_capacity_read(df, list_data))
# execl_voltage_read(df, list_data)
# execl_manufacturer_read(df, list_data)
# print(execl_supplier_read(df, list_data))

# df = excel_file_read('曜璿東命名規則 20240605-2.xlsx', '命名規則')
# df_supplier = excel_file_read('曜璿東命名規則 20240605-2.xlsx', '供應商編碼')
# excel_type_read_supplier(df_supplier)