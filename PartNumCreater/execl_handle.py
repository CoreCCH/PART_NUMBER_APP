import pandas as pd
import copy

def excel_file_read(file_name: str, sheet_name: str):
    # 讀取 Excel 文件
    return pd.read_excel(file_name, sheet_name)

#retrun array
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


    print(return_value)

def execl_size_read(df):
    size_value = []
    return_value = {}
    type_data = ""

    column_values = df.iloc[:, 12]
    column_values_code = df.iloc[:, 11]
    column_values_type = df.iloc[:, 10]

    for len in range(column_values.size):
        if (column_values_type[len] == "種類"):
            if(type_data != column_values_type[len+1]):
                type_data = column_values_type[len+1]
                size_value.clear()
        if (column_values[len] != "零件尺寸" and column_values[len] != "種類" and pd.notna(column_values[len]) and pd.notna(column_values[len])):
            print(column_values[len])
            size_value.append({str(column_values[len]).strip():str(column_values_code[len])})
            return_value.update({type_data:copy.deepcopy(size_value)})    

    print(return_value)

df = excel_file_read('曜璿東命名規則 20240605-2.xlsx', '命名規則')

# excel_list_read(df)
excel_type_read(df)
execl_size_read(df)