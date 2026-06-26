import openpyxl
import pandas as pd
from datetime import datetime, timedelta

jd_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda import checklist Maker and Verifier\IR53719 - AS pune\JobData_IR_53719_blank lic sheet.xlsx"
ir_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda import checklist Maker and Verifier\IR53719 - AS pune\Import Item Report 25-JUN-2026_04_44_PM.xlsx"
sm_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda import checklist Maker and Verifier\SAVWIPL Scrip Master 12-06-2026.xlsx"

# 1. Parse Item Report
df_rep = pd.read_excel(ir_path)
df_rep['calc_duty'] = df_rep['Assessable Value (INR)'] * df_rep['Basic Duty Rate'] / 100.0
required_duty = round(df_rep['calc_duty'].sum(), 2)
scheme_code = str(df_rep.loc[0, 'Exim Scheme Code']).strip()
print(f"Required duty: {required_duty:,.2f}")
print(f"Scheme code: {scheme_code}")

# 2. Load licenses
wb_scrip = openpyxl.load_workbook(sm_path, read_only=True, data_only=True)
ws_scrip = wb_scrip['Data 14072021']
scrip_rows = list(ws_scrip.iter_rows(values_only=True))
header_scrip = [str(h).strip().upper() for h in scrip_rows[0]]

lic_idx = header_scrip.index('LICNO/')
bal_idx = header_scrip.index('BALANCE')
port_idx = header_scrip.index('PORT OF REGISTRATION')
type_idx = header_scrip.index('LICENCE TYPE')
exp_idx = header_scrip.index('EXPIRY DATE')
val_idx = header_scrip.index('VALUE')
date_idx = header_scrip.index('LIC DATE')

mapped_lic_type = "RODTEP" if scheme_code.upper() == "RD" else scheme_code.upper()

active_licenses = []
for idx, r in enumerate(scrip_rows[1:], start=2):
    if len(r) <= max(lic_idx, bal_idx, port_idx, type_idx, exp_idx) or r[lic_idx] is None:
        continue
    lic_type = str(r[type_idx]).strip().upper() if r[type_idx] is not None else ""
    if mapped_lic_type in lic_type:
        bal = r[bal_idx]
        if bal is not None:
            try:
                bal_f = float(bal)
                if bal_f >= 3.00:
                    exp_val = r[exp_idx]
                    if isinstance(exp_val, str):
                        exp_date = datetime.strptime(exp_val.split()[0], "%Y-%m-%d")
                    elif isinstance(exp_val, datetime):
                        exp_date = exp_val
                    else:
                        exp_date = None
                        
                    reg_date_val = r[date_idx]
                    if isinstance(reg_date_val, str):
                        reg_date = datetime.strptime(reg_date_val.split()[0], "%Y-%m-%d")
                    elif isinstance(reg_date_val, datetime):
                        reg_date = reg_date_val
                    else:
                        reg_date = None
                        
                    active_licenses.append({
                        'lic_no': str(r[lic_idx]).split('.')[0].strip(),
                        'port': str(r[port_idx]).strip(),
                        'type': str(r[type_idx]).strip(),
                        'val': float(r[val_idx]) if r[val_idx] is not None else 0.0,
                        'bal': bal_f,
                        'expiry': exp_date,
                        'reg_date': reg_date
                    })
            except ValueError:
                pass
wb_scrip.close()

# Simulate tree table population & auto-checking
cumulative_bal = 0.0
checked_lics = []
unchecked_lics = []

for item in active_licenses:
    if cumulative_bal < required_duty:
        checked_lics.append(item)
        cumulative_bal += item['bal']
    else:
        unchecked_lics.append(item)

print(f"\nTotal active licenses: {len(active_licenses)}")
print(f"Checked licenses count: {len(checked_lics)}")
print(f"Unchecked licenses count: {len(unchecked_lics)}")
print(f"Selected License Balance (cumulative_bal): {cumulative_bal:,.2f}")

print("\nLast 10 Checked Licenses:")
for item in checked_lics[-10:]:
    print(f"LIC: {item['lic_no']} | Port: {item['port']} | Bal: {item['bal']:,.2f}")
