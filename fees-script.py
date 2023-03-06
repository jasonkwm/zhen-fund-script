import pandas as pd
import numpy as np

excel_file_name = input("Excelsheet name: ")
namelist_file_name = input("Name list fine name: ")

excel_sheet = pd.ExcelFile(excel_file_name)
name_list = pd.read_csv(namelist_file_name, header=None)[0].tolist()
sheet_list = excel_sheet.sheet_names

df_list = list()
tick = "ü"
id_col = "REGISTERED 42 ID"
col_name = "NEW_COLUMN"

# Stores each sheet in a list of sheets in df_list
for sheet in sheet_list:
    df_list.append(excel_sheet.parse(sheet))

# loop through each Quater
for df in df_list[1:]:
    new_col = []
    # Loop through intra id and create a new list of tick and untick boxes
    for id in df.loc[:, id_col]:
        if id in name_list:
            new_col.append(tick)
            name_list.remove(id)
        else:
            new_col.append(np.nan)
    new_col[0] = col_name
    check = 0
    # determine if column already exist. if column exist then append to it
    # else create a new column
    for col in df.columns:
        if type(df[col][0]) is str and df[col][0].find(col_name) != -1:
            check = 1
            new_col[0] = df[col][0]
            df[col] = new_col
            break
    if check == 0:
        df[col_name] = new_col

print("People that are not found in excel: ", name_list)
file_name = input("File name: ")
# Write to Excel with sheets
with pd.ExcelWriter(file_name + ".xlsx", engine="openpyxl") as writer:
    for i, sheet in enumerate(sheet_list):
        df_list[i].to_excel(writer, sheet)

