import openpyxl
import pandas as pd
from datetime import datetime

ir_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda import checklist Maker and Verifier\IR53719 - AS pune\Import Item Report 25-JUN-2026_04_44_PM.xlsx"
sm_path = r"c:\Users\Admin\Documents\NAGARKOT\Documentation\Skoda 1702\Skoda import checklist Maker and Verifier\SAVWIPL Scrip Master 12-06-2026.xlsx"

# 1. Parse Item Report
df_rep = pd.read_excel(ir_path)
df_rep['calc_duty'] = df_rep['Assessable Value (INR)'] * df_rep['Basic Duty Rate'] / 100.0
required_duty = round(df_rep['calc_duty'].sum(), 2)
scheme_code = str(df_rep.loc[0, 'Exim Scheme Code']).strip()
item_duties = [round(d, 2) for d in df_rep['calc_duty'].tolist()]

# 2. Load licenses
wb_scrip = openpyxl.load_workbook(sm_path, read_only=True, data_only=True)
ws_scrip = wb_scrip['Data 14072021']
scrip_rows = list(ws_scrip.iter_rows(values_only=True))
header_scrip = [str(h).strip().upper() for h in scrip_rows[0]]

lic_idx = header_scrip.index('LICNO/')
bal_idx = header_scrip.index('BALANCE')
port_idx = header_scrip.index('PORT OF REGISTRATION')
type_idx = header_scrip.index('LICENCE TYPE')

mapped_lic_type = "RODTEP" if scheme_code.upper() == "RD" else scheme_code.upper()

active_licenses = []
for idx, r in enumerate(scrip_rows[1:], start=2):
    if len(r) <= max(lic_idx, bal_idx, port_idx, type_idx) or r[lic_idx] is None:
        continue
    lic_type = str(r[type_idx]).strip().upper() if r[type_idx] is not None else ""
    if mapped_lic_type in lic_type:
        bal = r[bal_idx]
        if bal is not None:
            try:
                bal_f = float(bal)
                if bal_f > 0.00:
                    active_licenses.append({
                        'lic_no': str(r[lic_idx]).split('.')[0].strip(),
                        'port': str(r[port_idx]).strip(),
                        'type': str(r[type_idx]).strip(),
                        'bal': bal_f
                    })
            except ValueError:
                pass
wb_scrip.close()

# Filter/Check licenses
selected_lics = []
cumulative_bal = 0.0
for item in active_licenses:
    if item['bal'] >= 100.00 and cumulative_bal < required_duty:
        selected_lics.append(dict(item))
        cumulative_bal += item['bal']

print(f"Total Selected capacity: {cumulative_bal:,.2f}")
print(f"Required duty: {required_duty:,.2f}")

# Simulate allocation
lics = [dict(l) for l in selected_lics]
lic_ptr = 0
license_debit_totals = {lic['lic_no']: 0.0 for lic in lics}
initial_balances = {lic['lic_no']: lic['bal'] for lic in lics}

for duty in item_duties:
    if duty <= 0.0:
        continue
        
    duty_remaining = duty
    while duty_remaining > 0.005:
        if lic_ptr >= len(lics):
            break
        curr_lic = lics[lic_ptr]
        max_debitable = curr_lic['bal'] - 2.00
        if max_debitable <= 0.01:
            lic_ptr += 1
            continue
            
        if max_debitable >= duty_remaining:
            debit_amt = duty_remaining
            curr_lic['bal'] -= debit_amt
            license_debit_totals[curr_lic['lic_no']] += debit_amt
            duty_remaining = 0.0
        else:
            debit_amt = max_debitable
            curr_lic['bal'] = 2.00
            license_debit_totals[curr_lic['lic_no']] += debit_amt
            duty_remaining -= debit_amt
            lic_ptr += 1

print("\nAllocated Debits:")
total_debited = 0.0
for lic_no, val in license_debit_totals.items():
    if val > 0.00:
        init_bal = initial_balances[lic_no]
        rem_bal = init_bal - val
        total_debited += val
        print(f"LIC: {lic_no} | Init Bal: {init_bal:,.2f} | Debited: {val:,.2f} | Rem Bal: {rem_bal:,.2f}")

print(f"\nTotal Debited: {total_debited:,.2f}")
