import openpyxl

scrip_path = "SAVWIPL Scrip Master 12-06-2026.xlsx"
wb = openpyxl.load_workbook(scrip_path, read_only=True, data_only=True)
ws = wb['Data 14072021']
rows = list(ws.iter_rows(values_only=True))
headers = [str(h).strip().upper() for h in rows[0]]

lic_idx = headers.index('LICNO/')
bal_idx = headers.index('BALANCE')
type_idx = headers.index('LICENCE TYPE')
port_idx = headers.index('PORT OF REGISTRATION')

print("Loaded licenses in Excel order:")
count = 0
for idx, r in enumerate(rows[1:], start=2):
    if len(r) > lic_idx and r[lic_idx] is not None:
        lic_no = str(r[lic_idx]).split('.')[0].strip()
        lic_type = str(r[type_idx]).strip().upper() if r[type_idx] is not None else ""
        if 'RODTEP' in lic_type:
            bal = r[bal_idx]
            if bal is not None:
                try:
                    bal_f = float(bal)
                    if bal_f >= 3.00:
                        count += 1
                        print(f"{count}. LIC: {lic_no} | Port: {r[port_idx]} | Bal: {bal_f:,.2f}")
                except ValueError:
                    pass
wb.close()
