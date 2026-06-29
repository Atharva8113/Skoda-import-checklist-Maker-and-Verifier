import openpyxl
import pandas as pd
from datetime import datetime

ir_path = r"Import Item Report 26-JUN-2026_01_11_PM.xlsx"
sm_path = r"SAVWIPL Scrip Master 12-06-2026.xlsx"

# 1. Parse Item Report
df_rep = pd.read_excel(ir_path)
df_rep['calc_duty'] = (df_rep['Assessable Value (INR)'] * df_rep['Basic Duty Rate'] / 100.0).round(2)
required_duty = round(df_rep['calc_duty'].sum(), 2)
item_duties = df_rep['calc_duty'].tolist()

raw_be_no = df_rep.loc[0, 'BE No']
be_no = str(raw_be_no).strip() if not pd.isna(raw_be_no) else "nan"

be_date = df_rep.loc[0, 'BE Date']
if pd.isna(be_date):
    be_date_str = "nan"
elif isinstance(be_date, datetime):
    be_date_str = be_date.strftime('%d-%b-%Y')
else:
    be_date_str = str(be_date).split()[0]

job_no = str(df_rep.loc[0, 'Job No']).strip()

# Fallback auto-detect logic (what Auto-Detect selection will do)
raw_scheme_code = df_rep.loc[0, 'Exim Scheme Code']
if pd.isna(raw_scheme_code) or str(raw_scheme_code).strip().lower() in ('nan', ''):
    scheme_code = "RD"  # Default to RD (RODTEP) if empty
    print("Auto-detected empty Exim Scheme Code! Defaulting to RD.")
else:
    scheme_code = str(raw_scheme_code).strip()

print(f"Job Details:")
print(f"  Job No: {job_no}")
print(f"  BE No: {be_no}")
print(f"  BE Date: {be_date_str}")
print(f"  Scheme: {scheme_code}")
print(f"  Required Duty: {required_duty:,.2f}")

# Map scheme to mapped_lic_type
mapped_lic_type = "RODTEP" if scheme_code.upper() == "RD" else scheme_code.upper()

# Load licenses
wb_scrip = openpyxl.load_workbook(sm_path, read_only=True, data_only=True)
ws_scrip = wb_scrip['Data 14072021']
scrip_rows = list(ws_scrip.iter_rows(values_only=True))
header_scrip = [str(h).strip().upper() for h in scrip_rows[0]]

lic_idx = header_scrip.index('LICNO/')
bal_idx = header_scrip.index('BALANCE')
port_idx = header_scrip.index('PORT OF REGISTRATION')
type_idx = header_scrip.index('LICENCE TYPE')

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

print(f"\nSuccessfully loaded {len(active_licenses)} active {mapped_lic_type} licenses from Scrip Master.")
