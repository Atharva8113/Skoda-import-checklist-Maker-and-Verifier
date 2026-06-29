import pandas as pd

file_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda License management system\SAVWIPL Scrip Master TEST.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")
print("Total rows loaded by pandas:", len(df))

# Count non-empty license numbers
non_empty = df[df['LICNO/'].notna()]
print("Non-empty license number rows:", len(non_empty))

# Print the last 10 rows
print("\nLast 10 rows:")
print(df.tail(10).to_dict(orient='records'))
