import pandas as pd

file_path = "Import Item Report 26-JUN-2026_01_11_PM.xlsx"
df = pd.read_excel(file_path)
print("Columns in file:", df.columns.tolist())
print("\nFirst row of data:")
for col in df.columns:
    print(f"{col}: {df.loc[0, col]} (type: {type(df.loc[0, col])})")
