# Google Sheets Integration Setup Guide

Follow these simple steps to configure your central license master database in Google Sheets.

---

## Step 1: Create the Google Sheet
1. Open [Google Sheets](https://sheets.google.com) and create a **Blank Spreadsheet**.
2. Rename the spreadsheet to something descriptive, e.g. `Skoda License Master`.
3. Rename the first sheet tab (at the bottom) to exactly:
   **`Data`**
4. Create a second sheet tab (at the bottom) and name it exactly:
   **`Debit Ledger`**
5. Create a third sheet tab (at the bottom) and name it exactly:
   **`Used License`**

---

## Step 2: Paste the Setup Headers
To ensure correct formatting, copy-paste these headers as the first row of each sheet.

### For `Data` (Sheet 1):
Paste these into Row 1:
```
LICNO/ | LIC DATE | PORT OF REGISTRATION | LICENCE TYPE | VALUE | EXPIRY DATE | COUNT OF JOB | USED VALUE | BALANCE | JOB NO
```

### For `Debit Ledger` (Sheet 2):
Paste these into Row 1:
```
LICNO. | Used value from Lic for this item | LIC used in the Job | Invoice No | Product Desc | Assessable Value of the item | Basic Duty Rate | Basic Duty Value of item | Timestamp
```

---

## Step 3: Add the Google Apps Script
1. In your Google Sheet, click **Extensions** in the top menu, then select **Apps Script**.
2. Delete any code inside the editor, and paste the code below:

```javascript
// Security token to prevent unauthorized access. Match this in config.json
const SECURITY_TOKEN = "NagarkotSkoda2026";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.token !== SECURITY_TOKEN) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Unauthorized" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const action = data.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "fetch") {
      const dataSheet = ss.getSheetByName("Data");
      const boeSheet = ss.getSheetByName("Debit Ledger");
      
      let dataRows = dataSheet ? dataSheet.getDataRange().getValues() : [];
      let boeRows = boeSheet ? boeSheet.getDataRange().getValues() : [];
      
      dataRows = formatRows(ss, dataRows);
      boeRows = formatRows(ss, boeRows);
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: dataRows,
        boe: boeRows
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "debit") {
      const dataSheet = ss.getSheetByName("Data");
      const boeSheet = ss.getSheetByName("Debit Ledger");
      
      if (!dataSheet || !boeSheet) {
        throw new Error("Required sheets not found");
      }
      
      const newDebits = data.debits;
      
      // 1. Bulk append to BOE sheet (Passbook log)
      if (newDebits.length > 0) {
        const tz = ss.getSpreadsheetTimeZone();
        const timestampStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
        
        const boeRows = newDebits.map(function(d) {
          return [
            d.lic_no,     // LICNO.
            d.val,        // Used value from Lic for this item
            d.job_no,     // LIC used in the Job
            d.inv_no,     // Invoice No
            d.desc,       // Product Desc
            d.av,         // Assessable Value of the item
            d.rate,       // Basic Duty Rate
            d.duty_val,   // Basic Duty Value of item
            timestampStr  // Timestamp
          ];
        });
        const lastRow = boeSheet.getLastRow();
        boeSheet.getRange(lastRow + 1, 1, boeRows.length, boeRows[0].length).setValues(boeRows);
        
        // 2. Incremental update of Data sheet columns
        const dataRange = dataSheet.getDataRange();
        const dataValues = dataRange.getValues();
        if (dataValues.length > 1) {
          const header = dataValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
          const licIdx = header.indexOf('LICNO/');
          const valIdx = header.indexOf('VALUE');
          const noboeIdx = header.indexOf('COUNT OF JOB') !== -1 ? header.indexOf('COUNT OF JOB') : 6;
          const usedIdx = header.indexOf('USED VALUE') !== -1 ? header.indexOf('USED VALUE') : 7;
          const balIdx = header.indexOf('BALANCE') !== -1 ? header.indexOf('BALANCE') : 8;
          const jobIdx = header.indexOf('JOB NO') !== -1 ? header.indexOf('JOB NO') : 9;
          
          if (licIdx !== -1 && usedIdx !== -1 && balIdx !== -1) {
            // Build index map of licenses in Data sheet
            const licRowMap = {};
            for (let i = 1; i < dataValues.length; i++) {
              const licNo = String(dataValues[i][licIdx]).replace(".0", "").trim();
              if (licNo) {
                licRowMap[licNo] = i;
              }
            }
            
            // Loop through new debits to update in-memory array
            newDebits.forEach(function(d) {
              const licNo = String(d.lic_no).replace(".0", "").trim();
              const rIdx = licRowMap[licNo];
              if (rIdx !== undefined) {
                const debitVal = parseFloat(d.val) || 0;
                const jobNo = String(d.job_no || "").trim();
                
                // Add to USED VALUE
                const prevUsed = parseFloat(dataValues[rIdx][usedIdx]) || 0;
                dataValues[rIdx][usedIdx] = parseFloat((prevUsed + debitVal).toFixed(2));
                
                // Subtract from BALANCE
                const baseVal = parseFloat(dataValues[rIdx][valIdx]) || 0;
                dataValues[rIdx][balIdx] = parseFloat((baseVal - dataValues[rIdx][usedIdx]).toFixed(2));
                
                // Update JOB NO list and count if present
                if (jobNo && jobIdx !== -1 && noboeIdx !== -1) {
                  let existingJobs = String(dataValues[rIdx][jobIdx] || "").trim();
                  let jobsList = existingJobs ? existingJobs.split(",").map(function(j) { return j.trim(); }) : [];
                  if (jobsList.indexOf(jobNo) === -1) {
                    jobsList.push(jobNo);
                    dataValues[rIdx][jobIdx] = jobsList.join(", ");
                    dataValues[rIdx][noboeIdx] = jobsList.length;
                  }
                }
              }
            });
            
            // Re-partition rows: keep active ones, move completed ones to "Used License"
            const keepRows = [dataValues[0]];
            const moveRows = [];
            for (let i = 1; i < dataValues.length; i++) {
              const balance = parseFloat(dataValues[i][balIdx]) || 0;
              if (balance <= 1.0) {
                moveRows.push(dataValues[i]);
              } else {
                keepRows.push(dataValues[i]);
              }
            }
            
            // Clear Data sheet and write back active rows
            dataSheet.clearContents();
            dataSheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
            
            // Append moveRows to Used License sheet
            if (moveRows.length > 0) {
              let usedLicSheet = ss.getSheetByName("Used License");
              if (!usedLicSheet) {
                usedLicSheet = ss.insertSheet("Used License");
                usedLicSheet.getRange(1, 1, 1, dataValues[0].length).setValues([dataValues[0]]);
              }
              const lastRow = usedLicSheet.getLastRow();
              usedLicSheet.getRange(lastRow + 1, 1, moveRows.length, moveRows[0].length).setValues(moveRows);
            }
          }
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "add_licenses") {
      const dataSheet = ss.getSheetByName("Data");
      if (!dataSheet) {
        throw new Error("Data sheet not found");
      }
      
      const newLicenses = data.licenses;
      
      if (newLicenses.length > 0) {
        const keepRows = [];
        const moveRows = [];
        
        newLicenses.forEach(function(l) {
          const balance = l.balance !== undefined ? l.balance : l.val;
          const row = [
            l.lic_no,
            l.date,
            l.port,
            l.type,
            l.val,
            l.expiry,
            l.noboe !== undefined ? l.noboe : 0,
            l.used_val !== undefined ? l.used_val : 0,
            l.balance !== undefined ? l.balance : l.val,
            l.job_no !== undefined ? l.job_no : ""
          ];
          
          if (balance <= 1.0) {
            moveRows.push(row);
          } else {
            keepRows.push(row);
          }
        });
        
        // Append active ones to Data sheet
        if (keepRows.length > 0) {
          const lastRow = dataSheet.getLastRow();
          dataSheet.getRange(lastRow + 1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
        }
        
        // Append depleted ones to Used License sheet
        if (moveRows.length > 0) {
          let usedLicSheet = ss.getSheetByName("Used License");
          if (!usedLicSheet) {
            usedLicSheet = ss.insertSheet("Used License");
            const header = dataSheet.getRange(1, 1, 1, 10).getValues()[0];
            usedLicSheet.getRange(1, 1, 1, header.length).setValues([header]);
          }
          const lastRow = usedLicSheet.getLastRow();
          usedLicSheet.getRange(lastRow + 1, 1, moveRows.length, moveRows[0].length).setValues(moveRows);
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Invalid Action" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function formatRows(ss, rows) {
  if (!rows || rows.length === 0) return rows;
  const tz = ss.getSpreadsheetTimeZone();
  return rows.map(function(row) {
    return row.map(function(cell) {
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, tz, "yyyy-MM-dd");
      }
      return cell;
    });
  });
}

// One-time utility: Run this from the Apps Script editor to move all existing used licenses from 'Data' to 'Used License'
function moveExistingUsedLicenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data");
  if (!dataSheet) return;
  
  const dataRange = dataSheet.getDataRange();
  const dataValues = dataRange.getValues();
  if (dataValues.length <= 1) return;
  
  const header = dataValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
  const balIdx = header.indexOf('BALANCE');
  if (balIdx === -1) return;
  
  const keepRows = [dataValues[0]];
  const moveRows = [];
  
  for (let i = 1; i < dataValues.length; i++) {
    const balance = parseFloat(dataValues[i][balIdx]) || 0;
    if (balance <= 1.0) {
      moveRows.push(dataValues[i]);
    } else {
      keepRows.push(dataValues[i]);
    }
  }
  
  if (moveRows.length > 0) {
    dataSheet.clearContents();
    dataSheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
    
    let usedLicSheet = ss.getSheetByName("Used License");
    if (!usedLicSheet) {
      usedLicSheet = ss.insertSheet("Used License");
      usedLicSheet.getRange(1, 1, 1, dataValues[0].length).setValues([dataValues[0]]);
    }
    const lastRow = usedLicSheet.getLastRow();
    usedLicSheet.getRange(lastRow + 1, 1, moveRows.length, moveRows[0].length).setValues(moveRows);
    
    Logger.log("Successfully moved " + moveRows.length + " existing used licenses to 'Used License' sheet.");
  } else {
    Logger.log("No used licenses found to move.");
  }
}
```

3. Click the **Save** icon (floppy disk) to save the script.

---

## Step 4: Deploy as a Web App
1. In the Apps Script window, click the **Deploy** button at the top right, then select **New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Configure the settings:
   - **Description:** `Skoda License Web App Proxy`
   - **Execute as:** `Me (your-email@gmail.com)`
   - **Who has access:** `Anyone` *(Note: This is required so the Python app can access it. Security is maintained via the SECURITY_TOKEN parameter inside the code.)*
4. Click **Deploy**.
5. Google will ask you to authorize access. Click **Authorize access**, choose your account, click **Advanced** (at the bottom), click **Go to Untitled project (unsafe)**, and click **Allow**.
6. Once deployed, copy the **Web app URL** provided. (It will look like `https://script.google.com/macros/s/AKfycb.../exec`).

---

## Step 5: Save in Application Configuration
1. Open the Skoda License Automation Tool.
2. Go to the **Database Master** tab.
3. Paste your Web App URL into the **Google Web App URL** text field.
4. Click **Save Configuration**.
5. Click **Refresh Master** to load the spreadsheet data.
