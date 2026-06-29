import pandas as pd
import openpyxl
from datetime import datetime, timedelta

file_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda License management system\SAVWIPL Scrip Master TEST.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")
cols = [str(c).strip().upper() for c in df.columns]

print("Excel Columns in File:", cols)

lic_col = -1
date_col = -1
port_col = -1
type_col = -1
val_col = -1
exp_col = -1
noboe_col = -1
used_col = -1
bal_col = -1
job_col = -1

for idx, col in enumerate(cols):
    col_str = str(col).strip().upper()
    if 'LICNO' in col_str or 'LIC REF' in col_str or 'LICENSE NO' in col_str:
        lic_col = idx
    elif 'LIC DATE' in col_str or ('DATE' in col_str and 'EXPIRY' not in col_str):
        date_col = idx
    elif 'PORT' in col_str:
        port_col = idx
    elif 'TYPE' in col_str:
        type_col = idx
    elif 'VALUE' in col_str or 'VAL' in col_str:
        if 'USED' not in col_str and 'BAL' not in col_str:
            val_col = idx
    elif 'EXPIRY' in col_str:
        exp_col = idx
    elif 'COUNT' in col_str or 'NO OF BOE' in col_str or 'BOE' in col_str:
        noboe_col = idx
    elif 'USED' in col_str:
        used_col = idx
    elif 'BAL' in col_str:
        bal_col = idx
    elif 'JOB' in col_str:
        job_col = idx

print("\nMatched Column Indices:")
print(f"lic_col: {lic_col}, date_col: {date_col}, port_col: {port_col}, type_col: {type_col}")
print(f"val_col: {val_col}, exp_col: {exp_col}, noboe_col: {noboe_col}")
print(f"used_col: {used_col}, bal_col: {bal_col}, job_col: {job_col}")

# Let's inspect the first row parsing:
row_list = list(df.iloc[0])
print("\nFirst Row Values in Excel:", row_list)

val_used = row_list[used_col] if used_col != -1 else None
print(f"Raw Used Value: {val_used}, Type: {type(val_used)}")
