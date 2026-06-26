# Nagarkot Skoda License Automation Tool - User Guide

This tool automates the allocation of active customs licenses (such as RoDTEP) to import checklists. It calculates the basic duty required for each item, allocates it from the earliest-expiring active licenses, updates the checklist's `LICENSE` sheet for re-importing into Logisys, and logs the debits in the Scrip Master database.

---

## 1. Prerequisites and Setup
1. Place the compiled **`Skoda_License_Automation.exe`** (located in the `dist` folder) in a convenient folder on your PC.
2. Ensure you have the company logo image (`logo.png`) in the same folder as the `.exe` so the application is properly branded (it will fallback to text if missing).

---

## 2. Input Files Required
To run the automation, you will need three files:
1. **JobData Excel file:** The checklist Excel file exported from Logisys for a specific Job.
   - *Example:* `JobData_IR_53719_blank lic sheet.xlsx`
   - *Status:* The `LICENSE` sheet in this file must be empty (only containing the headers).
2. **Import Item Report:** The Excel report downloaded from Logisys listing all items, quantities, assessable values, and basic duty rates for that Job.
   - *Example:* `Import Item Report 25-JUN-2026_04_44_PM.xlsx`
3. **Scrip Master Excel:** The global registry spreadsheet maintained for active licenses and balances.
   - *Example:* `SAVWIPL Scrip Master 12-06-2026.xlsx`
   - *Status:* Must contain the `Data 14072021` sheet (active licenses) and `BOE` sheet (debits logs).

---

## 3. Step-by-Step Walkthrough

### Step 1: Select the Input Files
1. Double-click `Skoda_License_Automation.exe` to launch the application.
2. Click **Browse** next to each field to select:
   - **JobData Excel**
   - **Item Report Excel**
   - **Scrip Master Excel**

### Step 2: Load and Analyze
1. Click the **LOAD AND ANALYZE FILES** button.
2. The application will scan the files and display:
   - **Job Details Summary:** Job Number, Bill of Entry (BE) Number, BE Date, Import Port, Scheme type, and number of checklist items.
   - **Financials:** Total duty required for the Job, the cumulative balance of selected licenses, and coverage status (Fully Covered vs. Shortfall).
   - **Active Licenses List:** All active licenses of the matching scheme type (e.g. RoDTEP) are loaded from the Scrip Master, sorted by earliest expiry date, and displayed in a table.
   
### Step 3: Review and Refine License Selection
- By default, the tool will **automatically pre-select** (checkmark `[ X ]`) the earliest expiring licenses that have enough combined balance to cover the Job's total duty.
- If you know a license is not registered in ICEGATE for this specific Job, or is otherwise unavailable:
  - Click on the checkbox `[ X ]` in the first column of the table to **uncheck** it.
  - The tool will automatically re-calculate the selected balance and update the status indicator instantly.

### Step 4: Run the Allocation
1. Click **GENERATE LOGISYS EXCEL & DEBIT LICENSES**.
2. The tool will:
   - Greedy-allocate the items to the selected licenses in order of expiry date.
   - Handle 0% duty items by mapping them to the current license with a `0.00` debit value (as required for declaration).
   - Automatically **split items** if a license's balance runs out mid-item (allocating the remaining balance to the current license and starting the next license for the remainder of the item's duty, split proportionally by CIF value and quantity).
   - Generate a **new processed JobData file** in the same folder as the raw file, named `JobData_<JobNo>_Processed_<Timestamp>.xlsx`.
   - Update the global **Scrip Master database** in-place by appending a new debit entry for each license used to the `BOE` sheet.
3. Review the **Execution Logs** text pane at the bottom for a detailed line-by-line report of the allocation.
4. When complete, a success dialog box will appear.

---

## 4. Re-Importing into Logisys
1. Open Logisys and navigate to your Job.
2. Import the newly generated processed JobData file (e.g., `JobData_53719_Processed_...xlsx`).
3. The Logisys system will read the populated `LICENSE` sheet and set the effective payable basic duty to **zero**.

---

## 5. Troubleshooting & FAQ

* **Shortfall Status:** If the selected licenses don't have enough balance to cover the duty, check additional licenses in the table, or add new credit entries to your Scrip Master.
* **Open File Errors:** Ensure that none of the input files are open in Microsoft Excel while running the automation. Excel locks the files, which prevents the tool from saving modifications.
* **Formulas Recalculation:** The Scrip Master sheet `Data 14072021` uses Excel formulas to calculate `Used value` and `Balance`. The static cell values will update automatically once you open the Scrip Master in Microsoft Excel.
