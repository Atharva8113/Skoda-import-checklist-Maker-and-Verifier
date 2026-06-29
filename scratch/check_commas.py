import pandas as pd

file_path = "Import Item Report 26-JUN-2026_01_11_PM.xlsx"
df = pd.read_excel(file_path)

numeric_cols = ['Assessable Value (INR)', 'Basic Duty Rate', 'Quantity']
for col in numeric_cols:
    if col in df.columns:
        print(f"\nChecking column: {col}")
        for idx, val in enumerate(df[col]):
            if isinstance(val, str):
                print(f"  Row {idx+2}: Found string '{val}'")
            else:
                print(f"  Row {idx+2}: Value={val} (type={type(val)})")
