# Nagarkot Skoda License Automation Tool - User Guide

This tool automates the allocation of active customs licenses (such as RoDTEP) to import checklists. It calculates the basic duty required for each item, allocates it from the earliest-expiring active licenses, updates the checklist's `LICENSE` sheet for re-importing into Logisys, and logs the debits in a centralized **Google Sheets** database.

---

## 1. Central Google Sheets Database Integration
We have replaced the local Excel-based Scrip Master with a **centralized Google Sheets database**. This allows:
1. **Multi-User Access:** Multiple users can run the tool simultaneously on different computers, and balances are updated and shared in real-time.
2. **Simplified Inputs:** You no longer need to browse or manage local Scrip Master Excel files during checklists generation.
3. **Internal Passbook Logging:** Debits are recorded item-by-item in a detailed passbook layout, including invoice number, product description, and duty details.

To configure the Google Sheet and its Apps Script Web App, please refer to the [GOOGLE_SHEET_SETUP.md](file:///c:/Users/Admin/Documents/NAGARKOT/Documentation/Skoda%201702/Skoda%20License%20management%20system/GOOGLE_SHEET_SETUP.md) guide.

---

## 2. Input Files Required
To run the automation, you only need **two local files** now:
1. **JobData Excel file:** The checklist Excel file exported from Logisys for a specific Job.
   - *Example:* `JobData_IR_53719_blank lic sheet.xlsx`
   - *Status:* The `LICENSE` sheet in this file must be empty (only containing the headers).
2. **Import Item Report:** The Excel report downloaded from Logisys listing all items, quantities, assessable values, product descriptions, and basic duty rates for that Job.
   - *Example:* `Import Item Report 25-JUN-2026_04_44_PM.xlsx`

---

## 3. Step-by-Step Walkthrough

### Step 1: Configure the Cloud Database (One-time Setup)
1. Double-click `Skoda_License_Automation.exe` to launch the application.
2. Navigate to the **Database Master** tab.
3. Paste your published Google Apps Script **Web App URL** into the *Google Web App URL* field.
4. Click **Save Configuration**.
5. Click **Refresh Database View** to verify the connection and see your active licenses loaded from the cloud.

### Step 2: Select the Local Files and Load
1. Navigate back to the **License Automation** tab.
2. Click **Browse** next to each field to select:
   - **JobData Excel**
   - **Item Report Excel**
3. Select the License Scheme (defaults to Auto-Detect).
4. Click **LOAD AND ANALYZE FILES**.
5. The application will fetch active licenses from Google Sheets and display:
   - **Job Details Summary:** Job Number, Bill of Entry (BE) Number, BE Date, Import Port, Scheme type, and number of checklist items.
   - **Financials:** Total duty required for the Job, the cumulative balance of selected licenses, and coverage status (Fully Covered vs. Shortfall).
   - **Active Licenses List:** All active licenses of the matching scheme type (e.g. RoDTEP) loaded from the Google Sheets database, sorted by earliest expiry date.

### Step 3: Review License Selection
- By default, the tool will **automatically pre-select** (checkmark `[ X ]`) the earliest expiring licenses that have a balance above **3.00 INR** and are needed to cover the Job's total duty.
- To check or uncheck any license, simply click on its row in the table.
- The tool will instantly re-calculate the selected capacity and update the coverage indicators.

### Step 4: Run the Allocation
1. Click **GENERATE LOGISYS EXCEL & DEBIT LICENSES**.
2. The tool will:
   - Greedy-allocate the items to the selected licenses in order of expiry date.
   - Handle 0% duty items by mapping them to the current license with a `0.00` debit value (as required for declaration).
   - Automatically **split items** if a license's balance runs out mid-item (allocating the remaining balance to the current license and starting the next license for the remainder of the item's duty, split proportionally by CIF value and quantity).
   - Generate a **new processed JobData file** in the same folder as the raw file, named `JobData_<JobNo>_Processed_<Timestamp>.xlsx`.
   - **Push debits directly to Google Sheets** cloud. This automatically appends debit logs to the `BOE` passbook sheet and recalculates license balances in the `Data` sheet.
3. Review the **Execution Logs** text pane at the bottom for a detailed line-by-line report of the allocation.
4. When complete, a success dialog box will appear and the active license list will refresh automatically.

---

## 4. Managing Licenses (Database Master Tab)
Instead of opening Google Sheets manually in your browser, you can add new licenses directly through the tool:

### Method A: Paste License Rows
1. Open the spreadsheet where you have new licenses.
2. Select and copy rows. The expected columns (separated by tabs or commas) are:
   `LICNO/` | `LIC Date` | `Port Of Registration` | `Licence Type` | `Value` | `[Expiry date]`
3. Paste the rows into the text area in the **Add New Licenses** section.
4. Click **Push Pasted Licenses to Cloud**.

### Method B: Upload Excel File
1. Click **Browse** under *Method B* to select an Excel file containing new licenses.
2. Ensure the Excel file contains columns matching: `LICNO/`, `LIC Date`, `Port Of Registration`, `Licence Type`, `Value`, `Expiry Date`.
3. Click **Import and Push Excel to Cloud**.

---

## 5. Re-Importing into Logisys
1. Open Logisys and navigate to your Job.
2. Import the newly generated processed JobData file (e.g., `JobData_53719_Processed_...xlsx`).
3. The Logisys system will read the populated `LICENSE` sheet and set the effective payable basic duty to **zero**.
