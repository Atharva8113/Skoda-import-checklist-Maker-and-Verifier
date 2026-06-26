import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import re
import pandas as pd
import openpyxl
from datetime import datetime, timedelta

# --- MONKEYPATCH FOR OPENPYXL TO WRITE FORMULA CACHED VALUES ---
from xml.etree.ElementTree import Element, SubElement
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.cell.rich_text import CellRichText
from openpyxl.cell._writer import _set_attributes, safe_string, whitespace

# Registry for cached values of formulas (mapped by (sheet_title, cell_coordinate))
CACHED_VALUES = {}

def custom_etree_write_cell(xf, worksheet, cell, styled=None):
    value, attributes = _set_attributes(cell, styled)

    el = Element("c", attributes)
    if value is None or value == "":
        xf.write(el)
        return

    is_formula = (cell.data_type == 'f')

    if is_formula:
        attrib = {}
        if isinstance(value, ArrayFormula):
            attrib = dict(value)
            value = value.text
        elif isinstance(value, DataTableFormula):
            attrib = dict(value)
            value = None

        formula = SubElement(el, 'f', attrib)
        if value is not None and not attrib.get('t') == "dataTable":
            formula.text = value[1:]
            
            # Check if there is a custom cached value
            key = (worksheet.title, cell.coordinate)
            if key in CACHED_VALUES:
                value = CACHED_VALUES[key]
            else:
                value = None

    if cell.data_type == 's':
        if isinstance(value, CellRichText):
            el.append(value.to_tree())
        else:
            inline_string = Element("is")
            text = Element('t')
            text.text = value
            whitespace(text)
            inline_string.append(text)
            el.append(inline_string)
    else:
        cell_content = SubElement(el, 'v')
        if value is not None:
            cell_content.text = safe_string(value)

    xf.write(el)

# Apply monkeypatch
import openpyxl.cell._writer
import openpyxl.worksheet._writer
openpyxl.cell._writer.etree_write_cell = custom_etree_write_cell
openpyxl.worksheet._writer.write_cell = custom_etree_write_cell
# -------------------------------------------------------------

# Design Tokens (Corporate Blue & White Theme)
COLOR_PRIMARY = "#0F3057"       # Corporate Dark Blue
COLOR_SECONDARY = "#00587A"     # Medium Blue
COLOR_BG = "#F4F6F9"            # Light Gray-Blue Background
COLOR_CARD = "#FFFFFF"          # White Card Background
COLOR_TEXT = "#333333"          # Dark Gray Text
COLOR_ACCENT = "#17B978"        # Green accent for success / action
COLOR_ERROR = "#FF4D4D"         # Red accent for error
FONT_FAMILY = "Segoe UI"

class LicenseAutomationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Nagarkot Skoda License Automation Tool")
        self.root.geometry("1100x820")
        self.root.configure(bg=COLOR_BG)
        
        # State Variables
        self.job_data_path = tk.StringVar()
        self.item_report_path = tk.StringVar()
        self.scrip_master_path = tk.StringVar()
        
        self.job_info = {}
        self.active_licenses = []
        self.required_duty = 0.0
        self.selected_duty_covered = 0.0
        self.estimated_debit = 0.0
        self.item_duties = []
        
        self.create_styles()
        self.build_ui()

    def create_styles(self):
        """Configure custom styles for ttk widgets."""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Frame styles
        self.style.configure('TFrame', background=COLOR_BG)
        self.style.configure('Card.TFrame', background=COLOR_CARD, relief='flat', borderwidth=0)
        
        # Label styles
        self.style.configure('TLabel', background=COLOR_BG, foreground=COLOR_TEXT, font=(FONT_FAMILY, 10))
        self.style.configure('Header.TLabel', background=COLOR_PRIMARY, foreground='#FFFFFF', font=(FONT_FAMILY, 16, 'bold'))
        self.style.configure('Section.TLabel', background=COLOR_CARD, foreground=COLOR_PRIMARY, font=(FONT_FAMILY, 12, 'bold'))
        self.style.configure('Summary.TLabel', background=COLOR_CARD, foreground=COLOR_TEXT, font=(FONT_FAMILY, 10, 'bold'))
        
        # Button styles
        self.style.configure('TButton', font=(FONT_FAMILY, 10, 'bold'), borderwidth=0)
        self.style.map('TButton',
            background=[('active', COLOR_SECONDARY), ('!disabled', COLOR_PRIMARY)],
            foreground=[('active', '#FFFFFF'), ('!disabled', '#FFFFFF')]
        )
        
        self.style.configure('Action.TButton', font=(FONT_FAMILY, 11, 'bold'))
        self.style.map('Action.TButton',
            background=[('active', '#149F67'), ('!disabled', COLOR_ACCENT)],
            foreground=[('active', '#FFFFFF'), ('!disabled', '#FFFFFF')]
        )

        # Entry styles
        self.style.configure('TEntry', fieldbackground='#FFFFFF', bordercolor='#CCCCCC', lightcolor='#CCCCCC', darkcolor='#CCCCCC')
        
        # Treeview styles
        self.style.configure('Treeview',
            background=COLOR_CARD,
            fieldbackground=COLOR_CARD,
            foreground=COLOR_TEXT,
            rowheight=25,
            font=(FONT_FAMILY, 10)
        )
        self.style.configure('Treeview.Heading',
            background=COLOR_PRIMARY,
            foreground='#FFFFFF',
            font=(FONT_FAMILY, 10, 'bold'),
            borderwidth=1,
            relief='flat'
        )
        self.style.map('Treeview.Heading',
            background=[('active', COLOR_SECONDARY), ('!disabled', COLOR_PRIMARY)],
            foreground=[('active', '#FFFFFF'), ('!disabled', '#FFFFFF')]
        )

    def build_ui(self):
        """Assemble the main window components."""
        # --- HEADER BLOCK ---
        header_frame = tk.Frame(self.root, bg=COLOR_PRIMARY, height=80)
        header_frame.pack(fill='x', side='top')
        header_frame.pack_propagate(False)
        
        # Logo handling (PIL-free fallback)
        logo_loaded = False
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd(), "logo.png")
        if os.path.exists(logo_path):
            try:
                # PhotoImage natively supports PNG
                self.logo_img = tk.PhotoImage(file=logo_path)
                # Resize if necessary (subsample or zoom, e.g. subsample 3 to reduce size if it's too big)
                if self.logo_img.width() > 150 or self.logo_img.height() > 150:
                    self.logo_img = self.logo_img.subsample(3)
                
                logo_label = tk.Label(header_frame, image=self.logo_img, bg=COLOR_PRIMARY)
                logo_label.pack(side='left', padx=20, pady=5)
                logo_loaded = True
            except Exception as e:
                print(f"Warning: Logo failed to load: {e}")
        
        if not logo_loaded:
            # Stylized text placeholder for logo
            logo_label = tk.Label(header_frame, text="NAGARKOT", fg="#FFFFFF", bg=COLOR_PRIMARY, font=(FONT_FAMILY, 14, 'bold'))
            logo_label.pack(side='left', padx=25, pady=20)
            
        title_label = ttk.Label(header_frame, text="SKODA LICENSE AUTOMATION TOOL", style='Header.TLabel')
        title_label.pack(side='left', expand=True, fill='x', padx=20)
        
        # --- MAIN CONTAINER ---
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill='both', expand=True)
        
        # Left Panel (Inputs and Job Details)
        left_panel = ttk.Frame(main_frame)
        left_panel.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # Card 1: File Pickers
        files_card = ttk.Frame(left_panel, style='Card.TFrame', padding="15")
        files_card.pack(fill='x', side='top', pady=(0, 10))
        
        ttk.Label(files_card, text="File Configuration", style='Section.TLabel').pack(anchor='w', pady=(0, 10))
        
        # JobData Path Input
        row_jd = ttk.Frame(files_card, style='Card.TFrame')
        row_jd.pack(fill='x', pady=5)
        ttk.Label(row_jd, text="JobData Excel:", width=18, anchor='w', style='Summary.TLabel').pack(side='left')
        ttk.Entry(row_jd, textvariable=self.job_data_path).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(row_jd, text="Browse", command=self.browse_job_data).pack(side='left')
        
        # Item Report Path Input
        row_ir = ttk.Frame(files_card, style='Card.TFrame')
        row_ir.pack(fill='x', pady=5)
        ttk.Label(row_ir, text="Item Report Excel:", width=18, anchor='w', style='Summary.TLabel').pack(side='left')
        ttk.Entry(row_ir, textvariable=self.item_report_path).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(row_ir, text="Browse", command=self.browse_item_report).pack(side='left')
        
        # Scrip Master Path Input
        row_sm = ttk.Frame(files_card, style='Card.TFrame')
        row_sm.pack(fill='x', pady=5)
        ttk.Label(row_sm, text="Scrip Master Excel:", width=18, anchor='w', style='Summary.TLabel').pack(side='left')
        ttk.Entry(row_sm, textvariable=self.scrip_master_path).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(row_sm, text="Browse", command=self.browse_scrip_master).pack(side='left')
        
        # Load Data Button
        ttk.Button(files_card, text="LOAD AND ANALYZE FILES", command=self.load_and_analyze_data).pack(fill='x', pady=(15, 0))
        
        # Card 2: Job details summary
        self.summary_card = ttk.Frame(left_panel, style='Card.TFrame', padding="15")
        self.summary_card.pack(fill='both', expand=True)
        
        ttk.Label(self.summary_card, text="Job Details Summary", style='Section.TLabel').pack(anchor='w', pady=(0, 10))
        
        self.summary_text_frame = ttk.Frame(self.summary_card, style='Card.TFrame')
        self.summary_text_frame.pack(fill='both', expand=True)
        
        # Initial Placeholder text in Summary
        self.summary_placeholder = ttk.Label(self.summary_text_frame, text="Please load the JobData and Item Report files to view summary.", foreground="#777777", style='TLabel')
        self.summary_placeholder.pack(anchor='w', pady=10)
        
        # Right Panel (Licenses and Logging)
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        # Card 3: Active Licenses Table
        lic_card = ttk.Frame(right_panel, style='Card.TFrame', padding="15")
        lic_card.pack(fill='both', expand=True, side='top', pady=(0, 10))
        
        ttk.Label(lic_card, text="Active Licenses in Scrip Master", style='Section.TLabel').pack(anchor='w', pady=(0, 5))
        ttk.Label(lic_card, text="Sorts automatically by expiry date. Check/uncheck to select licenses for allocation.", font=(FONT_FAMILY, 9, 'italic'), foreground="#666666").pack(anchor='w', pady=(0, 10))
        
        # Treeview for active licenses
        tree_scroll = ttk.Scrollbar(lic_card)
        tree_scroll.pack(side='right', fill='y')
        
        # Tree view with columns
        self.lic_tree = ttk.Treeview(
            lic_card, 
            columns=('selected', 'lic_no', 'port', 'type', 'balance', 'expiry'), 
            show='headings', 
            yscrollcommand=tree_scroll.set,
            selectmode='none'
        )
        self.lic_tree.pack(fill='both', expand=True)
        tree_scroll.config(command=self.lic_tree.yview)
        
        # Configure Treeview tags for custom coloring and formatting
        self.lic_tree.tag_configure('near_expiry', foreground=COLOR_ERROR)
        self.lic_tree.tag_configure('checked', background="#E8F5E9")
        
        # Column Headings
        self.lic_tree.heading('selected', text="[Select]")
        self.lic_tree.heading('lic_no', text="License No.")
        self.lic_tree.heading('port', text="Port")
        self.lic_tree.heading('type', text="Type")
        self.lic_tree.heading('balance', text="Balance (INR)")
        self.lic_tree.heading('expiry', text="Expiry Date")
        
        # Column Widths
        self.lic_tree.column('selected', width=80, anchor='center')
        self.lic_tree.column('lic_no', width=130, anchor='center')
        self.lic_tree.column('port', width=80, anchor='center')
        self.lic_tree.column('type', width=90, anchor='center')
        self.lic_tree.column('balance', width=130, anchor='e')
        self.lic_tree.column('expiry', width=110, anchor='center')
        
        # Row double-click or single-click to toggle checkbox
        self.lic_tree.bind("<ButtonRelease-1>", self.on_tree_click)
        
        # Card 4: Run Actions and Live log
        run_card = ttk.Frame(right_panel, style='Card.TFrame', padding="15")
        run_card.pack(fill='x', side='bottom')
        
        self.run_btn = ttk.Button(run_card, text="GENERATE logisys EXCEL & DEBIT LICENSES", style='Action.TButton', command=self.run_license_automation, state='disabled')
        self.run_btn.pack(fill='x', pady=(0, 10))
        
        # Log view
        ttk.Label(run_card, text="Execution Logs", style='Section.TLabel').pack(anchor='w', pady=(0, 5))
        log_scroll = ttk.Scrollbar(run_card)
        log_scroll.pack(side='right', fill='y')
        
        self.log_txt = tk.Text(run_card, height=8, wrap='word', yscrollcommand=log_scroll.set, font=('Consolas', 9), bg='#FAF9F6', fg='#333333', borderwidth=1, relief='solid')
        self.log_txt.pack(fill='x')
        log_scroll.config(command=self.log_txt.yview)
        
        # Add basic greeting to log
        self.log("System initialized. Select files and load data to begin.")

        # --- FOOTER BLOCK ---
        footer_frame = tk.Frame(self.root, bg="#E2E2E2", height=30)
        footer_frame.pack(fill='x', side='bottom')
        footer_frame.pack_propagate(False)
        
        footer_label = tk.Label(footer_frame, text="Developed for Skoda Imports Automation | Version 1.0.0", bg="#E2E2E2", fg="#555555", font=(FONT_FAMILY, 9))
        footer_label.pack(pady=5)

    def log(self, msg: str):
        """Append a message to the logging pane."""
        self.log_txt.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_txt.see(tk.END)

    def browse_job_data(self):
        filename = filedialog.askopenfilename(title="Select JobData Excel File", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.job_data_path.set(filename)

    def browse_item_report(self):
        filename = filedialog.askopenfilename(title="Select Import Item Report Excel", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.item_report_path.set(filename)

    def browse_scrip_master(self):
        filename = filedialog.askopenfilename(title="Select Scrip Master Excel", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            self.scrip_master_path.set(filename)

    def on_tree_click(self, event):
        """Toggle checkbox status on treeview row click anywhere on the row."""
        item_id = self.lic_tree.identify_row(event.y)
        if not item_id:
            return
            
        values = list(self.lic_tree.item(item_id, 'values'))
        # Toggle checkbox
        if values[0] == "☑":
            values[0] = "☐"
        else:
            values[0] = "☑"
            
        # Dynamically manage checked and near_expiry tags to retain formatting
        current_tags = list(self.lic_tree.item(item_id, 'tags'))
        if values[0] == "☑":
            if 'checked' not in current_tags:
                current_tags.append('checked')
        else:
            if 'checked' in current_tags:
                current_tags.remove('checked')
                
        self.lic_tree.item(item_id, values=values, tags=tuple(current_tags))
        
        # Recalculate summary metrics
        self.recalc_selected_duty_metrics()

    def recalc_selected_duty_metrics(self):
        """Calculate selected license sum and compare with required duty."""
        self.selected_duty_covered = 0.0
        selected_lics = []
        for item_id in self.lic_tree.get_children():
            values = self.lic_tree.item(item_id, 'values')
            if values[0] == "☑":
                try:
                    bal_val = float(values[4])
                    self.selected_duty_covered += bal_val
                    selected_lics.append({
                        'lic_no': values[1],
                        'bal': bal_val
                    })
                except ValueError:
                    pass
        
        # Simulate allocation to get estimated debit amount
        self.estimated_debit, _ = self.simulate_debit(selected_lics)
        
        # Update summary UI
        self.update_summary_card()

    def simulate_debit(self, selected_lics):
        """Simulate the greedy allocation to find the exact debit amount for each license."""
        if not hasattr(self, 'item_duties') or not self.item_duties:
            return 0.0, {}
            
        # Copy selected lics to avoid modifying state
        lics = [dict(l) for l in selected_lics]
        lic_ptr = 0
        license_debit_totals = {lic['lic_no']: 0.0 for lic in lics}
        
        for duty in self.item_duties:
            if duty <= 0.0:
                continue
                
            duty_remaining = duty
            while duty_remaining > 0.005:
                if lic_ptr >= len(lics):
                    break
                curr_lic = lics[lic_ptr]
                max_debitable = curr_lic['bal'] - 2.00
                if max_debitable <= 0.01:
                    # Skip this license as we cannot debit without leaving less than 2.00 balance
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
                    
        total_estimated_debit = sum(license_debit_totals.values())
        return total_estimated_debit, license_debit_totals

    def load_and_analyze_data(self):
        """Parse files, extract job summary, load licenses, and update UI."""
        jd_path = self.job_data_path.get()
        ir_path = self.item_report_path.get()
        sm_path = self.scrip_master_path.get()
        
        if not (jd_path and ir_path and sm_path):
            messagebox.showerror("Error", "Please configure all three Excel files first.")
            return
            
        if not (os.path.exists(jd_path) and os.path.exists(ir_path) and os.path.exists(sm_path)):
            messagebox.showerror("Error", "One or more files do not exist at the specified paths.")
            return

        self.log("Starting analysis of files...")
        
        try:
            # 1. Parse Item Report for summary details
            self.log("Reading Import Item Report...")
            df_rep = pd.read_excel(ir_path)
            
            # Identify columns
            cols = df_rep.columns.tolist()
            req_cols = ['BE No', 'BE Date', 'Job No', 'Assessable Value (INR)', 'Basic Duty Rate', 'Exim Scheme Code', 'Quantity', 'Unit']
            missing = [r for r in req_cols if r not in cols]
            if missing:
                raise ValueError(f"Import Item Report is missing required column(s): {missing}")
            
            # Compute total required basic duty
            df_rep['calc_duty'] = (df_rep['Assessable Value (INR)'] * df_rep['Basic Duty Rate'] / 100.0).round(2)
            self.required_duty = round(df_rep['calc_duty'].sum(), 2)
            self.item_duties = df_rep['calc_duty'].tolist()
            
            # Extract basic job details
            be_no = str(df_rep.loc[0, 'BE No']).strip()
            be_date = df_rep.loc[0, 'BE Date']
            if isinstance(be_date, datetime):
                be_date_str = be_date.strftime('%d-%b-%Y')
            else:
                be_date_str = str(be_date).split()[0]
                
            job_no = str(df_rep.loc[0, 'Job No']).strip()
            scheme_code = str(df_rep.loc[0, 'Exim Scheme Code']).strip()
            total_items = len(df_rep)
            
            # 2. Extract port of registration from JobData GENERAL sheet
            self.log("Reading JobData general details...")
            wb_jd = openpyxl.load_workbook(jd_path, read_only=True)
            if 'GENERAL' not in wb_jd.sheetnames:
                raise ValueError("GENERAL sheet not found in JobData workbook.")
            ws_gen = wb_jd['GENERAL']
            rows_gen = list(ws_gen.iter_rows(values_only=True))
            headers_gen = [str(h).strip().upper() for h in rows_gen[0]]
            
            if 'CUSTOMSHOUSECODE' not in headers_gen:
                raise ValueError("Column CUSTOMSHOUSECODE not found in JobData GENERAL sheet.")
            port_col_idx = headers_gen.index('CUSTOMSHOUSECODE')
            import_port = str(rows_gen[1][port_col_idx]).strip()
            wb_jd.close()
            
            self.job_info = {
                'be_no': be_no,
                'be_date': be_date,
                'be_date_str': be_date_str,
                'job_no': job_no,
                'scheme': scheme_code,
                'total_items': total_items,
                'import_port': import_port
            }
            
            self.log(f"Job Details: Job No: {job_no}, BE No: {be_no}, Import Port: {import_port}, Scheme: {scheme_code}")
            self.log(f"Total Checklist Items: {total_items}, Required License Duty: {self.required_duty:,.2f} INR")

            # 3. Load Licenses from Scrip Master Data 14072021
            self.log("Loading active licenses from Scrip Master...")
            wb_scrip = openpyxl.load_workbook(sm_path, read_only=True, data_only=True)
            if 'Data 14072021' not in wb_scrip.sheetnames:
                raise ValueError("Sheet 'Data 14072021' not found in Scrip Master.")
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
            
            self.active_licenses = []
            for idx, r in enumerate(scrip_rows[1:], start=2):
                if len(r) <= max(lic_idx, bal_idx, port_idx, type_idx, exp_idx) or r[lic_idx] is None:
                    continue
                lic_type = str(r[type_idx]).strip().upper() if r[type_idx] is not None else ""
                
                # Filter by scheme type
                if mapped_lic_type in lic_type:
                    bal = r[bal_idx]
                    if bal is not None:
                        try:
                            bal_f = float(bal)
                            if bal_f > 0.00:
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
                                    
                                self.active_licenses.append({
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
            
            # Keep original order from Excel sheet (no sorting)
            
            self.log(f"Successfully loaded {len(self.active_licenses)} active {mapped_lic_type} licenses.")
            
            # Populate table
            self.populate_license_table()
            
            # Recalculate summary metrics
            self.recalc_selected_duty_metrics()
            
            # Enable execute button
            self.run_btn.config(state='normal')
            
        except Exception as e:
            self.log(f"Error during analysis: {e}")
            messagebox.showerror("Error", f"Failed to analyze files: {e}")

    def populate_license_table(self):
        """Fill treeview table and auto-check the licenses needed in their Excel order."""
        # Clear existing
        for item_id in self.lic_tree.get_children():
            self.lic_tree.delete(item_id)
            
        cumulative_bal = 0.0
        for i, item in enumerate(self.active_licenses):
            exp_str = item['expiry'].strftime('%d-%b-%Y') if item['expiry'] is not None else "N/A"
            
            # Greedy auto-selection based on order in table covering required duty.
            # We skip licenses with balance < 100.00 INR by default to avoid cluttering with tiny residual balances.
            if item['bal'] >= 100.00 and cumulative_bal < self.required_duty:
                selected_str = "☑"
                cumulative_bal += item['bal']
            else:
                selected_str = "☐"
                
            # Expiry is near (within 30 days) and base value > 500
            is_near = False
            if item['expiry'] is not None:
                days_to_expiry = (item['expiry'] - datetime.now()).days
                if days_to_expiry <= 30 and item['val'] > 500:
                    is_near = True
            
            tags = []
            if selected_str == "☑":
                tags.append('checked')
            if is_near:
                tags.append('near_expiry')
                
            self.lic_tree.insert('', 'end', values=(
                selected_str,
                item['lic_no'],
                item['port'],
                item['type'],
                f"{item['bal']:.2f}",
                exp_str
            ), tags=tuple(tags))

    def update_summary_card(self):
        """Update left summary panel with calculated details and shortfall indicators."""
        # Clear summary
        for widget in self.summary_text_frame.winfo_children():
            widget.destroy()
            
        if not self.job_info:
            return
            
        # Add labels
        grid_params = {'anchor': 'w', 'pady': 3}
        
        ttk.Label(self.summary_text_frame, text=f"Job Number: {self.job_info['job_no']}", style='Summary.TLabel').pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Bill of Entry No: {self.job_info['be_no']}", style='Summary.TLabel').pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Bill of Entry Date: {self.job_info['be_date_str']}", style='Summary.TLabel').pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Import Port: {self.job_info['import_port']}", style='Summary.TLabel').pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Required License Type: {self.job_info['scheme'].upper()} (RODTEP)", style='Summary.TLabel').pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Total Items in Checklist: {self.job_info['total_items']}", style='Summary.TLabel').pack(**grid_params)
        
        separator = ttk.Separator(self.summary_text_frame, orient='horizontal')
        separator.pack(fill='x', pady=10)
        
        # Simplified user-friendly labels
        cash_payment = max(0.0, self.required_duty - self.estimated_debit)
        
        ttk.Label(self.summary_text_frame, text=f"Total Duty to Pay: {self.required_duty:,.2f} INR", font=(FONT_FAMILY, 10, 'bold')).pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Selected License Capacity: {self.selected_duty_covered:,.2f} INR", font=(FONT_FAMILY, 10, 'bold')).pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Duty Covered by License (Exemption): {self.estimated_debit:,.2f} INR", font=(FONT_FAMILY, 10, 'bold')).pack(**grid_params)
        ttk.Label(self.summary_text_frame, text=f"Remaining Duty to Pay in Cash: {cash_payment:,.2f} INR", font=(FONT_FAMILY, 10, 'bold')).pack(**grid_params)
        
        # Check shortfall based on estimated debit
        if cash_payment <= 0.02:
            status_text = "Status: Fully Covered (No cash payment required)"
            status_color = COLOR_ACCENT
        else:
            status_text = f"Status: Partially Covered (Cash payment of {cash_payment:,.2f} INR required)"
            status_color = COLOR_ERROR
            
        ttk.Label(self.summary_text_frame, text=status_text, foreground=status_color, font=(FONT_FAMILY, 11, 'bold')).pack(**grid_params)

    def run_license_automation(self):
        """Execute license allocation, modify JobData, modify Scrip Master, and notify user."""
        jd_path = self.job_data_path.get()
        ir_path = self.item_report_path.get()
        sm_path = self.scrip_master_path.get()
        
        # Get selected license numbers from treeview
        selected_lic_nos = []
        for item_id in self.lic_tree.get_children():
            values = self.lic_tree.item(item_id, 'values')
            if values[0] == "☑":
                selected_lic_nos.append(values[1])
                
        if not selected_lic_nos:
            messagebox.showerror("Error", "No licenses selected. Please check at least one license.")
            return
            
        if self.estimated_debit < self.required_duty - 0.02:
            cash_needed = self.required_duty - self.estimated_debit
            ans = messagebox.askyesno("Warning", f"The selected licenses do not cover the total required duty. You will need to pay {cash_needed:,.2f} INR in cash to customs. Do you want to proceed?")
            if not ans:
                return

        self.log("Initiating license debiting process...")
        
        try:
            # 1. Build selected licenses list (pull matching references from self.active_licenses)
            selected_lics = []
            for lic_no in selected_lic_nos:
                for item in self.active_licenses:
                    if item['lic_no'] == lic_no:
                        selected_lics.append(dict(item))
                        break
            
            # 2. Read Item Report
            df_rep = pd.read_excel(ir_path)
            
            # 3. Read JobData ITEMS to establish correct serial indexes
            self.log("Opening JobData Excel for writing (preserving formulas)...")
            wb_jd = openpyxl.load_workbook(jd_path, data_only=False)
            ws_inv = wb_jd['INVOICES']
            inv_rows = list(ws_inv.iter_rows(values_only=True))
            inv_header = [str(h).strip() for h in inv_rows[0]]
            inv_sr_idx = inv_header.index('InvSrNo')
            inv_no_idx = inv_header.index('Invoice_No')
            
            inv_map = {}
            for r in inv_rows[1:]:
                if r[inv_sr_idx] is not None and r[inv_no_idx] is not None:
                    inv_map[str(r[inv_no_idx]).strip()] = str(r[inv_sr_idx]).strip()
            
            ws_items = wb_jd['ITEMS']
            items_rows = list(ws_items.iter_rows(values_only=True))
            items_header = [str(h).strip() for h in items_rows[0]]
            item_inv_sr_idx = items_header.index('InvSrNo')
            item_sr_idx = items_header.index('ItemSrNo')
            desc_idx = items_header.index('Product_Description')
            qty_idx = items_header.index('QTY')
            cth_idx = items_header.index('CTH')
            
            # 4. Perform Greedy Allocation Row-by-Row
            self.log("Running allocation algorithm...")
            lic_ptr = 0
            job_license_rows = []
            license_debit_totals = {lic['lic_no']: 0.0 for lic in selected_lics}
            
            for idx, r_row in df_rep.iterrows():
                rep_inv_no = str(r_row['Invoice No']).strip()
                rep_desc = str(r_row['Product Desc']).strip()
                rep_qty = float(r_row['Quantity'])
                rep_cth = str(r_row['CTH']).strip()
                
                av = float(r_row['Assessable Value (INR)'])
                rate = float(r_row['Basic Duty Rate'])
                
                # Match item in JobData ITEMS sheet to get InvSrNo and ItemSrNo
                inv_sr = inv_map.get(rep_inv_no)
                item_sr = None
                
                if inv_sr:
                    for item_row in items_rows[1:]:
                        if item_row[item_inv_sr_idx] is not None:
                            itm_inv_sr = str(item_row[item_inv_sr_idx]).strip()
                            if itm_inv_sr == inv_sr:
                                itm_desc = str(item_row[desc_idx]).strip() if item_row[desc_idx] is not None else ""
                                itm_qty = float(item_row[qty_idx]) if item_row[qty_idx] is not None else 0.0
                                itm_cth = str(item_row[cth_idx]).strip() if item_row[cth_idx] is not None else ""
                                
                                # Match criteria
                                if itm_desc == rep_desc and abs(itm_qty - rep_qty) < 0.01 and itm_cth == rep_cth:
                                    item_sr = str(item_row[item_sr_idx]).strip()
                                    break
                
                if not item_sr:
                    self.log(f"Warning: Could not match item row {idx} (Desc: {rep_desc[:30]}...) to ITEMS sheet. Skipping.")
                    continue
                
                # Calculate required duty (round to 2 decimals)
                duty = round(av * rate / 100.0, 2)
                
                # Allocate duty
                duty_remaining = duty
                
                # Edge case: 0 duty items still need mapping to the current license!
                if duty == 0.0:
                    if lic_ptr < len(selected_lics):
                        curr_lic = selected_lics[lic_ptr]
                        job_license_rows.append({
                            'Inv_SrNo': inv_sr,
                            'Item_SrNo': item_sr,
                            'License_No': curr_lic['lic_no'],
                            'License_Date': curr_lic['reg_date'].strftime('%d-%b-%Y') if curr_lic['reg_date'] else "",
                            'License_RegNo': curr_lic['lic_no'],
                            'License_RegDate': curr_lic['reg_date'].strftime('%d-%b-%Y') if curr_lic['reg_date'] else "",
                            'Reg_Port': curr_lic['port'],
                            'CIF_Value': av,
                            'DebitDeutyValue': 0.00,
                            'DebitQuantity': rep_qty,
                            'DebitQuantityUnitCode': r_row['Unit']
                        })
                    continue

                while duty_remaining > 0:
                    if lic_ptr >= len(selected_lics):
                        self.log(f"Warning: Insufficient license balance. Item {idx+1} (Duty: {duty_remaining:.2f}) left uncovered.")
                        break
                        
                    curr_lic = selected_lics[lic_ptr]
                    max_debitable = curr_lic['bal'] - 2.00
                    if max_debitable <= 0.01:
                        # Move to next license as this one cannot be debited without falling below 2.00 INR balance
                        lic_ptr += 1
                        if lic_ptr < len(selected_lics):
                            self.log(f"License depleted or insufficient (below 2.00 balance). Switching to License: {selected_lics[lic_ptr]['lic_no']}")
                        continue
                        
                    if max_debitable >= duty_remaining:
                        # Full debit from the license
                        debit_amt = duty_remaining
                        curr_lic['bal'] -= debit_amt
                        ratio = debit_amt / duty
                        job_license_rows.append({
                            'Inv_SrNo': inv_sr,
                            'Item_SrNo': item_sr,
                            'License_No': curr_lic['lic_no'],
                            'License_Date': curr_lic['reg_date'].strftime('%d-%b-%Y') if curr_lic['reg_date'] else "",
                            'License_RegNo': curr_lic['lic_no'],
                            'License_RegDate': curr_lic['reg_date'].strftime('%d-%b-%Y') if curr_lic['reg_date'] else "",
                            'Reg_Port': curr_lic['port'],
                            'CIF_Value': round(av * ratio, 2),
                            'DebitDeutyValue': round(debit_amt, 2),
                            'DebitQuantity': round(rep_qty * ratio, 3),
                            'DebitQuantityUnitCode': r_row['Unit']
                        })
                        license_debit_totals[curr_lic['lic_no']] += debit_amt
                        duty_remaining = 0
                    else:
                        # Debit up to max_debitable
                        debit_amt = max_debitable
                        curr_lic['bal'] = 2.00
                        ratio = debit_amt / duty
                        job_license_rows.append({
                            'Inv_SrNo': inv_sr,
                            'Item_SrNo': item_sr,
                            'License_No': curr_lic['lic_no'],
                            'License_Date': curr_lic['reg_date'].strftime('%d-%b-%Y') if curr_lic['reg_date'] else "",
                            'License_RegNo': curr_lic['lic_no'],
                            'License_RegDate': curr_lic['reg_date'].strftime('%d-%b-%Y') if curr_lic['reg_date'] else "",
                            'Reg_Port': curr_lic['port'],
                            'CIF_Value': round(av * ratio, 2),
                            'DebitDeutyValue': round(debit_amt, 2),
                            'DebitQuantity': round(rep_qty * ratio, 3),
                            'DebitQuantityUnitCode': r_row['Unit']
                        })
                        license_debit_totals[curr_lic['lic_no']] += debit_amt
                        duty_remaining -= debit_amt
                        
                        # Move to next license
                        lic_ptr += 1
                        if lic_ptr < len(selected_lics):
                            self.log(f"License depleted (left with exactly 2.00 balance). Switching to License: {selected_lics[lic_ptr]['lic_no']}")

            # 5. Populate JobData LICENSE Sheet
            self.log("Writing to JobData LICENSE sheet...")
            if 'LICENSE' not in wb_jd.sheetnames:
                ws_lic = wb_jd.create_sheet('LICENSE')
            else:
                ws_lic = wb_jd['LICENSE']
                
            # Clear existing data except header
            ws_lic.delete_rows(2, ws_lic.max_row)
            
            headers = [
                'Inv_SrNo', 'Item_SrNo', 'License_RefNo', 'License_No', 'License_Date',
                'License_RegNo', 'License_RegDate', 'Reg_Port', 'License_ItemSrNo',
                'CIF_Value', 'DebitDeutyValue', 'DebitQuantity', 'DebitQuantityUnitCode'
            ]
            for col_num, header in enumerate(headers, 1):
                ws_lic.cell(row=1, column=col_num, value=header)
                
            # Write allocations
            for row_idx, data in enumerate(job_license_rows, start=2):
                ws_lic.cell(row=row_idx, column=1, value=data['Inv_SrNo'])
                ws_lic.cell(row=row_idx, column=2, value=data['Item_SrNo'])
                ws_lic.cell(row=row_idx, column=3, value=None)
                ws_lic.cell(row=row_idx, column=4, value=data['License_No'])
                ws_lic.cell(row=row_idx, column=5, value=data['License_Date'])
                ws_lic.cell(row=row_idx, column=6, value=data['License_RegNo'])
                ws_lic.cell(row=row_idx, column=7, value=data['License_RegDate'])
                ws_lic.cell(row=row_idx, column=8, value=data['Reg_Port'])
                ws_lic.cell(row=row_idx, column=9, value=1)
                ws_lic.cell(row=row_idx, column=10, value=data['CIF_Value'])
                ws_lic.cell(row=row_idx, column=11, value=data['DebitDeutyValue'])
                ws_lic.cell(row=row_idx, column=12, value=data['DebitQuantity'])
                ws_lic.cell(row=row_idx, column=13, value=data['DebitQuantityUnitCode'])
            
            # Save updated JobData
            dir_name = os.path.dirname(jd_path)
            date_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # Sanitize job number to prevent slash directory errors
            clean_job_no_str = re.sub(r'[\\/*?:"<>|]', '_', self.job_info['job_no'])
            job_name = f"JobData_{clean_job_no_str}_Processed_{date_stamp}.xlsx"
            output_jd_path = os.path.join(dir_name, job_name)
            
            self.log("Saving processed JobData Excel...")
            wb_jd.save(output_jd_path)
            wb_jd.close()
            self.log(f"JobData saved successfully to: {os.path.basename(output_jd_path)}")

            # 6. Update Scrip Master BOE sheet
            self.log("Opening Scrip Master for writing (preserving formulas)...")
            wb_sm = openpyxl.load_workbook(sm_path, data_only=False)
            if 'BOE' not in wb_sm.sheetnames:
                ws_boe = wb_sm.create_sheet('BOE')
                ws_boe.append(['LICNO.', 'VALUE', 'BOENO', 'BOE DATE', 'Job No', 'Column1'])
            else:
                ws_boe = wb_sm['BOE']
            
            clean_job_no = None
            try:
                job_match = re.search(r'/(\d+)/', self.job_info['job_no'])
                if job_match:
                    clean_job_no = int(job_match.group(1))
                else:
                    digits = ''.join(c for c in self.job_info['job_no'] if c.isdigit())
                    clean_job_no = int(digits) if digits else 0
            except Exception:
                clean_job_no = 0
                
            clean_be_no = int(self.job_info['be_no']) if self.job_info['be_no'].isdigit() else 0
            
            be_date_val = self.job_info['be_date']
            if isinstance(be_date_val, str):
                try:
                    be_date_val = datetime.strptime(be_date_val.split()[0], "%Y-%m-%d")
                except Exception:
                    pass

            for lic_no, val in license_debit_totals.items():
                if val > 0.01:
                    new_row = [
                        int(lic_no) if lic_no.isdigit() else lic_no,
                        round(val, 2),
                        clean_be_no,
                        be_date_val,
                        clean_job_no,
                        None
                    ]
                    ws_boe.append(new_row)
                    self.log(f"Appended debit of {val:,.2f} INR to BOE for License {lic_no}")
            
            # Recalculate calculations for licenses
            self.log("Recalculating Scrip Master formula values...")
            recalculate_scrip_master_formulas(wb_sm)
            
            self.log("Saving updated Scrip Master Excel...")
            wb_sm.save(sm_path)
            wb_sm.close()
            self.log("Scrip Master saved and updated successfully.")

            # Show Success Dialog
            messagebox.showinfo("Success", f"License automation finished successfully!\n\n1. JobData updated and saved to:\n{os.path.basename(output_jd_path)}\n\n2. Scrip Master updated in-place.")
            
            # Reload everything to refresh the balance list
            self.load_and_analyze_data()
            
        except Exception as e:
            self.log(f"Error executing license allocation: {e}")
            messagebox.showerror("Error", f"Failed to complete license allocation: {e}")


def recalculate_scrip_master_formulas(wb):
    """Recalculate formulas in Data 14072021 based on BOE sheet entries and store in CACHED_VALUES."""
    if 'BOE' not in wb.sheetnames or 'Data 14072021' not in wb.sheetnames:
        return
        
    ws_boe = wb['BOE']
    boe_rows = list(ws_boe.iter_rows(values_only=True))
    if not boe_rows or len(boe_rows) < 2:
        boe_rows = [ ['LICNO.', 'VALUE', 'BOENO', 'BOE DATE', 'Job No'] ]
        
    header_boe = [str(h).strip().upper() for h in boe_rows[0]]
    
    lic_col = -1
    val_col = -1
    boeno_col = -1
    jobno_col = -1
    
    for idx, col in enumerate(header_boe):
        col_str = str(col).strip().upper()
        if 'LIC' in col_str:
            lic_col = idx
        elif 'VAL' in col_str:
            val_col = idx
        elif 'BOENO' in col_str or 'BOE NO' in col_str:
            boeno_col = idx
        elif 'JOB' in col_str:
            jobno_col = idx
            
    if lic_col == -1: lic_col = 0
    if val_col == -1: val_col = 1
    if boeno_col == -1: boeno_col = 2
    if jobno_col == -1: jobno_col = 4
    
    from collections import defaultdict
    used_map = defaultdict(float)
    boe_sum_map = defaultdict(float)
    job_sum_map = defaultdict(float)
    
    for row in boe_rows[1:]:
        if len(row) > max(lic_col, val_col) and row[lic_col] is not None:
            lic_str = str(row[lic_col]).replace(".0", "").strip()
            
            # Value
            val = float(row[val_col]) if val_col < len(row) and row[val_col] is not None else 0.0
            used_map[lic_str] += val
            
            # BOE No
            if boeno_col < len(row) and row[boeno_col] is not None:
                try:
                    boe_sum_map[lic_str] += float(row[boeno_col])
                except ValueError:
                    pass
                    
            # Job No
            if jobno_col < len(row) and row[jobno_col] is not None:
                try:
                    job_sum_map[lic_str] += float(row[jobno_col])
                except ValueError:
                    pass
                    
    # 2. Recalculate Data 14072021 sheet
    ws_data = wb['Data 14072021']
    data_rows = list(ws_data.iter_rows(values_only=True))
    if not data_rows:
        return
        
    header_scrip = [str(h).strip().upper() for h in data_rows[0]]
    lic_idx = header_scrip.index('LICNO/')
    val_idx = header_scrip.index('VALUE')
    date_idx = header_scrip.index('LIC DATE')
    
    # Locate other formula columns
    exp_col_letter = 'G'
    noboe_col_letter = 'H'
    used_col_letter = 'I'
    bal_col_letter = 'J'
    job_col_letter = 'K'
    
    for col_idx, h in enumerate(header_scrip, start=1):
        h_str = str(h).strip().upper()
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        if 'EXPIRY' in h_str:
            exp_col_letter = col_letter
        elif 'NO OF BOE' in h_str:
            noboe_col_letter = col_letter
        elif 'USED VALUE' in h_str or 'USED_VALUE' in h_str:
            used_col_letter = col_letter
        elif 'BALANCE' in h_str:
            bal_col_letter = col_letter
        elif 'JOB NO' in h_str or 'JOB_NO' in h_str:
            job_col_letter = col_letter
            
    global CACHED_VALUES
    CACHED_VALUES.clear()
    
    for row_idx, row in enumerate(data_rows[1:], start=2):
        if len(row) > lic_idx and row[lic_idx] is not None:
            lic_str = str(row[lic_idx]).replace(".0", "").strip()
            
            # Base value
            base_val = float(row[val_idx]) if val_idx < len(row) and row[val_idx] is not None else 0.0
            
            # Recalculated values
            used_val = round(used_map[lic_str], 2)
            no_boe = round(boe_sum_map[lic_str], 2)
            job_no_sum = round(job_sum_map[lic_str], 2)
            bal_val = round(base_val - used_val, 2)
            
            # Expiry date (LIC DATE + 720 days)
            exp_val = None
            lic_date_val = row[date_idx]
            if lic_date_val is not None:
                if isinstance(lic_date_val, datetime):
                    exp_val = lic_date_val + timedelta(days=720)
                elif isinstance(lic_date_val, str):
                    try:
                        dt = datetime.strptime(lic_date_val.split()[0], "%Y-%m-%d")
                        exp_val = dt + timedelta(days=720)
                    except Exception:
                        pass
                        
            if exp_val is not None:
                exp_val_str = exp_val.strftime("%Y-%m-%d %H:%M:%S")
            else:
                exp_val_str = None
                
            # Add to CACHED_VALUES registry
            sheet_title = ws_data.title
            CACHED_VALUES[(sheet_title, f"{exp_col_letter}{row_idx}")] = exp_val_str
            CACHED_VALUES[(sheet_title, f"{noboe_col_letter}{row_idx}")] = no_boe
            CACHED_VALUES[(sheet_title, f"{used_col_letter}{row_idx}")] = used_val
            CACHED_VALUES[(sheet_title, f"{bal_col_letter}{row_idx}")] = bal_val
            CACHED_VALUES[(sheet_title, f"{job_col_letter}{row_idx}")] = job_no_sum


if __name__ == "__main__":
    root = tk.Tk()
    app = LicenseAutomationApp(root)
    root.mainloop()
