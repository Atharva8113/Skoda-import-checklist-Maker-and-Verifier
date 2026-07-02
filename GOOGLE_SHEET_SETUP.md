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

*(Note: The `Blacklisted Licenses` tab will be created automatically by the script when you first refresh or use the tool!)*

---

## Step 2: Paste the Setup Headers
To ensure correct formatting, copy-paste these headers as the first row of each sheet.

### For `Data` (Sheet 1):
Paste these into Row 1:
```
LICNO/ | LIC DATE | PORT OF REGISTRATION | LICENCE TYPE | VALUE | EXPIRY DATE | COUNT OF JOB | USED VALUE | BALANCE | JOB NO | ADDED TIME
```

### For `Debit Ledger` (Sheet 2):
Paste these into Row 1:
```
LICNO. | Used value from Lic for this item | LIC used in the Job | Invoice No | Inv_SrNo | ItemSrNo | Product Desc | Assessable Value of the item | Basic Duty Rate | Basic Duty Value of item | Timestamp | Replaced With
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
    
    // Ensure Blacklisted Licenses sheet exists
    let blacklistSheet = ss.getSheetByName("Blacklisted Licenses");
    if (!blacklistSheet) {
      blacklistSheet = ss.insertSheet("Blacklisted Licenses");
      blacklistSheet.getRange(1, 1, 1, 3).setValues([["LICNO/", "DATE BLACKLISTED", "REASON"]]);
    }
    
    if (action === "fetch") {
      const dataSheet = ss.getSheetByName("Data");
      const boeSheet = ss.getSheetByName("Debit Ledger");
      const usedSheet = ss.getSheetByName("Used License");
      
      let dataRows = dataSheet ? dataSheet.getDataRange().getValues() : [];
      let usedRows = usedSheet ? usedSheet.getDataRange().getValues() : [];
      let boeRows = [["LIC USED IN THE JOB"]];
      let blacklistRows = blacklistSheet.getDataRange().getValues();
      
      // OPTIMIZATION: Pull ONLY the 3rd column (Job No) of Debit Ledger to save network bandwidth
      if (boeSheet && boeSheet.getLastRow() > 1) {
        const jobNos = boeSheet.getRange(2, 3, boeSheet.getLastRow() - 1, 1).getValues();
        boeRows = boeRows.concat(jobNos);
      }
      
      dataRows = formatRows(ss, dataRows);
      usedRows = formatRows(ss, usedRows);
      blacklistRows = formatRows(ss, blacklistRows);
      
      // Parse blacklisted licenses into a set for fast filtering
      const blacklistedSet = new Set();
      if (blacklistRows.length > 1) {
        for (let i = 1; i < blacklistRows.length; i++) {
          if (blacklistRows[i][0]) {
            blacklistedSet.add(String(blacklistRows[i][0]).replace(".0", "").trim());
          }
        }
      }
      
      // Enrich blacklist rows with full details from Data and Used License sheets
      const blacklistDetails = [["LICNO/", "DATE BLACKLISTED", "REASON", "LIC DATE", "PORT OF REGISTRATION", "LICENCE TYPE", "VALUE", "EXPIRY DATE"]];
      if (blacklistRows.length > 1) {
        let allLicRows = [];
        if (dataSheet && dataSheet.getLastRow() > 1) {
          allLicRows = allLicRows.concat(formatRows(ss, dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 10).getValues()));
        }
        if (usedSheet && usedSheet.getLastRow() > 1) {
          allLicRows = allLicRows.concat(formatRows(ss, usedSheet.getRange(2, 1, usedSheet.getLastRow() - 1, 10).getValues()));
        }
        
        const licMap = {};
        allLicRows.forEach(function(row) {
          const lNo = String(row[0]).replace(".0", "").trim();
          if (lNo) {
            licMap[lNo] = row;
          }
        });
        
        for (let i = 1; i < blacklistRows.length; i++) {
          const licNo = String(blacklistRows[i][0]).replace(".0", "").trim();
          const detail = licMap[licNo] || [];
          
          blacklistDetails.push([
            licNo,
            blacklistRows[i][1], // DATE BLACKLISTED
            blacklistRows[i][2], // REASON
            detail[1] !== undefined ? detail[1] : "", // LIC DATE
            detail[2] !== undefined ? detail[2] : "", // PORT OF REGISTRATION
            detail[3] !== undefined ? detail[3] : "", // LICENCE TYPE
            detail[4] !== undefined ? detail[4] : "", // VALUE
            detail[5] !== undefined ? detail[5] : ""  // EXPIRY DATE
          ]);
        }
      }
      
      // Filter out blacklisted licenses from active dataRows (except the header)
      let filteredDataRows = [];
      if (dataRows.length > 0) {
        filteredDataRows.push(dataRows[0]);
        const header = dataRows[0].map(function(h) { return String(h).trim().toUpperCase(); });
        const licIdx = header.indexOf('LICNO/');
        
        if (licIdx !== -1) {
          for (let i = 1; i < dataRows.length; i++) {
            const licNo = String(dataRows[i][licIdx]).replace(".0", "").trim();
            if (!blacklistedSet.has(licNo)) {
              filteredDataRows.push(dataRows[i]);
            }
          }
        } else {
          filteredDataRows = dataRows;
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: filteredDataRows,
        used: usedRows,
        boe: boeRows,
        blacklist: blacklistDetails
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
        
        const boeRange = boeSheet.getDataRange();
        const boeValues = boeRange.getValues();
        const boeHeader = boeValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
        
        const boeRows = newDebits.map(function(d) {
          const row = [];
          for (let col = 0; col < boeHeader.length; col++) {
            const h = boeHeader[col];
            if (h.indexOf('LICNO') !== -1) row.push(d.lic_no);
            else if (h.indexOf('USED VALUE FROM LIC') !== -1 || h.indexOf('VALUE') !== -1) row.push(d.val);
            else if (h.indexOf('LIC USED IN THE JOB') !== -1 || h.indexOf('JOB') !== -1) row.push(d.job_no);
            else if (h.indexOf('INVOICE NO') !== -1) row.push(d.inv_no);
            else if (h.indexOf('INV_SRNO') !== -1 || h.indexOf('INVSRNO') !== -1) row.push(d.inv_sr_no || "");
            else if (h.indexOf('ITEM_SRNO') !== -1 || h.indexOf('ITEMSRNO') !== -1) row.push(d.item_sr_no || "");
            else if (h.indexOf('PRODUCT DESC') !== -1) row.push(d.desc);
            else if (h.indexOf('ASSESSABLE VALUE') !== -1) row.push(d.av);
            else if (h.indexOf('BASIC DUTY RATE') !== -1 || h.indexOf('RATE') !== -1) row.push(d.rate);
            else if (h.indexOf('BASIC DUTY VALUE') !== -1 || h.indexOf('DUTY') !== -1) row.push(d.duty_val);
            else if (h.indexOf('TIMESTAMP') !== -1 || h.indexOf('TIME') !== -1) row.push(timestampStr);
            else if (h.indexOf('REPLACED WITH') !== -1) row.push("");
            else row.push("");
          }
          return row;
        });
        const lastRow = boeSheet.getLastRow();
        boeSheet.getRange(lastRow + 1, 1, boeRows.length, boeRows[0].length).setValues(boeRows);
        
        // 2. Incremental update of Data sheet columns
        updateLicenseBalances(ss, dataSheet, newDebits, true);
        
        // 3. Move fully used licenses
        moveExistingUsedLicenses();
        
        // 4. Rebuild job lists
        rebuildLicenseJobLists(ss, dataSheet);
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
        // Collect all existing license numbers from Data, Used License and Blacklisted sheets to prevent duplicates
        const existingLics = new Set();
        const sheetsToScan = [ss.getSheetByName("Data"), ss.getSheetByName("Used License"), ss.getSheetByName("Blacklisted Licenses")];
        sheetsToScan.forEach(function(sh) {
          if (sh && sh.getLastRow() > 1) {
            const licNos = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
            licNos.forEach(function(row) {
              const lNo = String(row[0]).replace(".0", "").trim();
              if (lNo) {
                existingLics.add(lNo);
              }
            });
          }
        });
        
        const keepRows = [];
        const moveRows = [];
        
        const tz = ss.getSpreadsheetTimeZone();
        const timestampStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
        
        newLicenses.forEach(function(l) {
          const licNoStr = String(l.lic_no).replace(".0", "").trim();
          if (existingLics.has(licNoStr)) {
            return; // Skip duplicate license numbers!
          }
          
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
            l.job_no !== undefined ? l.job_no : "",
            timestampStr // Column K: ADDED TIME
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
            const header = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues()[0];
            usedLicSheet.getRange(1, 1, 1, header.length).setValues([header]);
          }
          const lastRow = usedLicSheet.getLastRow();
          usedLicSheet.getRange(lastRow + 1, 1, moveRows.length, moveRows[0].length).setValues(moveRows);
        }
        
        // Move fully used licenses
        moveExistingUsedLicenses();

        var addedCount = keepRows.length + moveRows.length;
        var skippedCount = newLicenses.length - addedCount;
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          added: addedCount,
          skipped: skippedCount
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, added: 0, skipped: 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "blacklist_license") {
      const dataSheet = ss.getSheetByName("Data");
      const usedSheet = ss.getSheetByName("Used License");
      const boeSheet = ss.getSheetByName("Debit Ledger");
      
      // Uses global cleanVal
      
      const cleanDesc = function(desc) {
        return String(desc || "").toLowerCase().replace(/[^a-z0-9]/g, "");
      };
      
      // Normalize data.lic_no to an array of clean string values
      let licNoList = [];
      if (Array.isArray(data.lic_no)) {
        licNoList = data.lic_no.map(cleanVal);
      } else {
        licNoList = [cleanVal(data.lic_no)];
      }
      
      const jobNo = cleanVal(data.job_no);
      const reason = String(data.reason || "").trim();
      const reallocations = data.reallocations || [];
      
      // 1. Add all to Blacklisted Licenses sheet if not already there
      licNoList.forEach(function(licNo) {
        let isAlreadyBlacklisted = false;
        const lastBlRow = blacklistSheet.getLastRow();
        if (lastBlRow > 1) {
          const blValues = blacklistSheet.getRange(2, 1, lastBlRow - 1, 1).getValues();
          for (let i = 0; i < blValues.length; i++) {
            if (cleanVal(blValues[i][0]) === licNo) {
              isAlreadyBlacklisted = true;
              break;
            }
          }
        }
        if (!isAlreadyBlacklisted) {
          const tz = ss.getSpreadsheetTimeZone();
          const timestampStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
          blacklistSheet.appendRow([licNo, timestampStr, reason]);
        }
      });
      
      // 2. Perform Re-allocations in Balances
      if (reallocations.length > 0) {
        // Adjust old license balance (restoring the debited value)
        const oldDebitsRestore = reallocations.map(function(r) {
          return { lic_no: r.old_lic, val: -parseFloat(r.val), job_no: r.job_no };
        });
        updateLicenseBalances(ss, dataSheet, oldDebitsRestore, false);
        
        // Adjust new license balance (deducting the debited value)
        const newDebitsApply = reallocations.map(function(r) {
          return { lic_no: r.new_lic, val: parseFloat(r.val), job_no: r.job_no };
        });
        updateLicenseBalances(ss, dataSheet, newDebitsApply, true);
        
        // 3. Highlight old wrong rows and append new re-allocated rows in Debit Ledger
        if (boeSheet && boeSheet.getLastRow() > 1) {
          const boeRange = boeSheet.getDataRange();
          const boeValues = boeRange.getValues();
          const boeHeader = boeValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
          const boeLicIdx = boeHeader.indexOf('LICNO.');
          const boeJobIdx = boeHeader.indexOf('LIC USED IN THE JOB');
          const boeInvIdx = boeHeader.indexOf('INVOICE NO');
          const boeDescIdx = boeHeader.indexOf('PRODUCT DESC');
          
          let boeInvSrIdx = -1;
          for (let col = 0; col < boeHeader.length; col++) {
            if (boeHeader[col].indexOf('INV_SRNO') !== -1 || boeHeader[col].indexOf('INVSRNO') !== -1) {
              boeInvSrIdx = col;
              break;
            }
          }
          let boeItemSrIdx = -1;
          for (let col = 0; col < boeHeader.length; col++) {
            if (boeHeader[col].indexOf('ITEM_SRNO') !== -1 || boeHeader[col].indexOf('ITEMSRNO') !== -1) {
              boeItemSrIdx = col;
              break;
            }
          }
          
          if (boeLicIdx !== -1 && boeJobIdx !== -1 && boeInvIdx !== -1) {
            reallocations.forEach(function(r) {
              const oldLic = cleanVal(r.old_lic);
              const newLic = cleanVal(r.new_lic);
              const invNo = cleanVal(r.inv_no);
              const rInvSr = String(r.inv_sr_no || "").trim();
              const rItemSr = String(r.item_sr_no || "").trim();
              const rDesc = cleanDesc(r.desc);
              
              let originalRowVal = null;
              let bestMatchIdx = -1;
              
              // 1. Precise match using Inv_SrNo and ItemSrNo if columns exist
              if (boeInvSrIdx !== -1 && boeItemSrIdx !== -1) {
                for (let i = 1; i < boeValues.length; i++) {
                  const bLic = cleanVal(boeValues[i][boeLicIdx]);
                  const bJob = cleanVal(boeValues[i][boeJobIdx]);
                  const bInv = cleanVal(boeValues[i][boeInvIdx]);
                  const bInvSr = String(boeValues[i][boeInvSrIdx]).trim();
                  const bItemSr = String(boeValues[i][boeItemSrIdx]).trim();
                  
                  if (bLic === oldLic && bJob === jobNo && bInv === invNo && bInvSr === rInvSr && bItemSr === rItemSr) {
                    bestMatchIdx = i;
                    break;
                  }
                }
              }
              
              // 2. Fallback to product description match if columns don't exist or exact match failed
              if (bestMatchIdx === -1) {
                for (let i = 1; i < boeValues.length; i++) {
                  const bLic = cleanVal(boeValues[i][boeLicIdx]);
                  const bJob = cleanVal(boeValues[i][boeJobIdx]);
                  const bInv = cleanVal(boeValues[i][boeInvIdx]);
                  const bDesc = boeDescIdx !== -1 ? cleanDesc(boeValues[i][boeDescIdx]) : "";
                  
                  if (bLic === oldLic && bJob === jobNo && bInv === invNo) {
                    if (bDesc === rDesc || bDesc.indexOf(rDesc) !== -1 || rDesc.indexOf(bDesc) !== -1) {
                      bestMatchIdx = i;
                      break;
                    }
                  }
                }
              }
              
              // 3. Fallback to first matching invoice/license if all else fails
              if (bestMatchIdx === -1) {
                for (let i = 1; i < boeValues.length; i++) {
                  const bLic = cleanVal(boeValues[i][boeLicIdx]);
                  const bJob = cleanVal(boeValues[i][boeJobIdx]);
                  const bInv = cleanVal(boeValues[i][boeInvIdx]);
                  if (bLic === oldLic && bJob === jobNo && bInv === invNo) {
                    bestMatchIdx = i;
                    break;
                  }
                }
              }
              
              if (bestMatchIdx !== -1) {
                // Highlight the old incorrect row in light red
                boeSheet.getRange(bestMatchIdx + 1, 1, 1, boeHeader.length).setBackground("#ffebee");
                originalRowVal = boeValues[bestMatchIdx];
                
                const repIdx = boeHeader.indexOf('REPLACED WITH');
                if (repIdx !== -1) {
                  boeSheet.getRange(bestMatchIdx + 1, repIdx + 1).setValue(newLic || "REMOVED/UNALLOCATED");
                }
                
                // Mark this row as matched so subsequent items of same invoice won't match it again
                boeValues[bestMatchIdx][boeLicIdx] = "MATCHED_OLD_DEBIT";
              }
              
              if (originalRowVal) {
                const tz = ss.getSpreadsheetTimeZone();
                const timestampStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
                
                // Append a new row mapped dynamically to the headers
                const row = [];
                for (let col = 0; col < boeHeader.length; col++) {
                  const h = boeHeader[col];
                  if (h.indexOf('LICNO') !== -1) row.push(newLic);
                  else if (h.indexOf('USED VALUE FROM LIC') !== -1 || h.indexOf('VALUE') !== -1) row.push(parseFloat(r.val));
                  else if (h.indexOf('LIC USED IN THE JOB') !== -1 || h.indexOf('JOB') !== -1) row.push(jobNo);
                  else if (h.indexOf('INVOICE NO') !== -1) row.push(r.inv_no);
                  else if (h.indexOf('INV_SRNO') !== -1 || h.indexOf('INVSRNO') !== -1) row.push(r.inv_sr_no || "");
                  else if (h.indexOf('ITEM_SRNO') !== -1 || h.indexOf('ITEMSRNO') !== -1) row.push(r.item_sr_no || "");
                  else if (h.indexOf('PRODUCT DESC') !== -1) {
                    const idx = boeHeader.indexOf('PRODUCT DESC');
                    row.push(r.desc || (idx !== -1 && originalRowVal[idx]) || "");
                  }
                  else if (h.indexOf('ASSESSABLE VALUE') !== -1) {
                    const idx = boeHeader.indexOf('ASSESSABLE VALUE OF THE ITEM');
                    row.push(idx !== -1 && originalRowVal[idx] !== undefined ? originalRowVal[idx] : 0);
                  }
                  else if (h.indexOf('BASIC DUTY RATE') !== -1 || h.indexOf('RATE') !== -1) {
                    const idx = boeHeader.indexOf('BASIC DUTY RATE');
                    row.push(idx !== -1 && originalRowVal[idx] !== undefined ? originalRowVal[idx] : 0);
                  }
                  else if (h.indexOf('BASIC DUTY VALUE') !== -1 || h.indexOf('DUTY') !== -1) {
                    const idx = boeHeader.indexOf('BASIC DUTY VALUE OF ITEM');
                    row.push(idx !== -1 && originalRowVal[idx] !== undefined ? originalRowVal[idx] : 0);
                  }
                  else if (h.indexOf('TIMESTAMP') !== -1 || h.indexOf('TIME') !== -1) row.push(timestampStr);
                  else if (h.indexOf('REPLACED WITH') !== -1) row.push("");
                  else row.push("");
                }
                
                if (newLic) {
                  boeSheet.appendRow(row);
                  const lastRowNow = boeSheet.getLastRow();
                  boeSheet.getRange(lastRowNow, 1, 1, row.length).setBackground("#e8f5e9");
                }
              }
            });
          }
        }
        
        // Move fully consumed or restored licenses between sheets
        moveExistingUsedLicenses();
        
        // Highlight all blacklisted license rows in Data and Used License sheets
        applyBlacklistHighlighting(ss);
        
        // Rebuild job lists
        rebuildLicenseJobLists(ss, dataSheet);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "rectify_license") {
      const dataSheet = ss.getSheetByName("Data");
      const usedSheet = ss.getSheetByName("Used License");
      
      const licNo = String(data.lic_no).replace(".0", "").trim();
      const edited = data.edited_data || {};
      
      // 1. Remove from Blacklisted Licenses sheet
      if (blacklistSheet && blacklistSheet.getLastRow() > 1) {
        const blValues = blacklistSheet.getRange(1, 1, blacklistSheet.getLastRow(), 1).getValues();
        for (let i = 1; i < blValues.length; i++) {
          if (String(blValues[i][0]).replace(".0", "").trim() === licNo) {
            blacklistSheet.deleteRow(i + 1);
            break;
          }
        }
      }
      
      // 2. Find and update details in Data or Used License sheets
      const sheets = [dataSheet, usedSheet];
      let updated = false;
      
      sheets.forEach(function(sh) {
        if (!sh || updated) return;
        const range = sh.getDataRange();
        const values = range.getValues();
        if (values.length <= 1) return;
        
        const header = values[0].map(function(h) { return String(h).trim().toUpperCase(); });
        const licIdx = header.indexOf('LICNO/');
        const valIdx = header.indexOf('VALUE');
        const usedIdx = header.indexOf('USED VALUE') !== -1 ? header.indexOf('USED VALUE') : 7;
        const balIdx = header.indexOf('BALANCE') !== -1 ? header.indexOf('BALANCE') : 8;
        const dateIdx = header.indexOf('LIC DATE') !== -1 ? header.indexOf('LIC DATE') : 1;
        const portIdx = header.indexOf('PORT OF REGISTRATION') !== -1 ? header.indexOf('PORT OF REGISTRATION') : 2;
        const typeIdx = header.indexOf('LICENCE TYPE') !== -1 ? header.indexOf('LICENCE TYPE') : 3;
        const expIdx = header.indexOf('EXPIRY DATE') !== -1 ? header.indexOf('EXPIRY DATE') : 5;
        
        for (let i = 1; i < values.length; i++) {
          const lNo = String(values[i][licIdx]).replace(".0", "").trim();
          if (lNo === licNo) {
            if (edited.date && dateIdx !== -1) sh.getRange(i + 1, dateIdx + 1).setValue(edited.date);
            if (edited.port && portIdx !== -1) sh.getRange(i + 1, portIdx + 1).setValue(edited.port);
            if (edited.type && typeIdx !== -1) sh.getRange(i + 1, typeIdx + 1).setValue(edited.type);
            if (edited.val !== undefined && valIdx !== -1) sh.getRange(i + 1, valIdx + 1).setValue(parseFloat(edited.val));
            if (edited.expiry && expIdx !== -1) sh.getRange(i + 1, expIdx + 1).setValue(edited.expiry);
            
            // Re-calculate balance
            const newVal = edited.val !== undefined ? parseFloat(edited.val) : parseFloat(values[i][valIdx]);
            const prevUsed = parseFloat(values[i][usedIdx]) || 0;
            const newBal = parseFloat((newVal - prevUsed).toFixed(2));
            if (balIdx !== -1) {
              sh.getRange(i + 1, balIdx + 1).setValue(newBal);
            }
            updated = true;
            break;
          }
        }
      });
      
      // Recalculate sheets to move licenses based on new balances
      moveExistingUsedLicenses();
      
      // Re-apply highlights after move
      applyBlacklistHighlighting(ss);
      
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

function updateLicenseBalances(ss, dataSheet, debits, isDebit) {
  const usedSheet = ss.getSheetByName("Used License");
  const sheets = [dataSheet, usedSheet];
  
  sheets.forEach(function(sh) {
    if (!sh) return;
    const dataRange = sh.getDataRange();
    const dataValues = dataRange.getValues();
    if (dataValues.length <= 1) return;
    
    const header = dataValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
    
    let licIdx = -1;
    for (let i = 0; i < header.length; i++) {
      if (header[i].indexOf('LICNO') !== -1) {
        licIdx = i;
        break;
      }
    }
    
    const valIdx = header.indexOf('VALUE');
    const noboeIdx = header.indexOf('COUNT OF JOB') !== -1 ? header.indexOf('COUNT OF JOB') : 6;
    const usedIdx = header.indexOf('USED VALUE') !== -1 ? header.indexOf('USED VALUE') : 7;
    const balIdx = header.indexOf('BALANCE') !== -1 ? header.indexOf('BALANCE') : 8;
    const jobIdx = header.indexOf('JOB NO') !== -1 ? header.indexOf('JOB NO') : 9;
    
    if (licIdx !== -1 && usedIdx !== -1 && balIdx !== -1) {
      const licRowMap = {};
      for (let i = 1; i < dataValues.length; i++) {
        const licNo = cleanVal(dataValues[i][licIdx]);
        if (licNo) {
          licRowMap[licNo] = i;
        }
      }
      
      let modified = false;
      debits.forEach(function(d) {
        const licNo = cleanVal(d.lic_no);
        const rIdx = licRowMap[licNo];
        if (rIdx !== undefined) {
          const debitVal = parseFloat(d.val) || 0;
          const jobNo = String(d.job_no || "").trim();
          
          const prevUsed = parseFloat(dataValues[rIdx][usedIdx]) || 0;
          dataValues[rIdx][usedIdx] = parseFloat((prevUsed + debitVal).toFixed(2));
          
          const baseVal = parseFloat(dataValues[rIdx][valIdx]) || 0;
          dataValues[rIdx][balIdx] = parseFloat((baseVal - dataValues[rIdx][usedIdx]).toFixed(2));
          
          if (isDebit && jobNo && jobIdx !== -1 && noboeIdx !== -1) {
            let existingJobs = String(dataValues[rIdx][jobIdx] || "").trim();
            let jobsList = existingJobs ? existingJobs.split(",").map(function(j) { return j.trim(); }) : [];
            if (jobsList.indexOf(jobNo) === -1) {
              jobsList.push(jobNo);
              dataValues[rIdx][jobIdx] = jobsList.join(", ");
              dataValues[rIdx][noboeIdx] = jobsList.length;
            }
          } else if (!isDebit && jobNo && jobIdx !== -1 && noboeIdx !== -1) {
            let existingJobs = String(dataValues[rIdx][jobIdx] || "").trim();
            let jobsList = existingJobs ? existingJobs.split(",").map(function(j) { return j.trim(); }) : [];
            let pos = jobsList.indexOf(jobNo);
            if (pos !== -1) {
              jobsList.splice(pos, 1);
              dataValues[rIdx][jobIdx] = jobsList.join(", ");
              dataValues[rIdx][noboeIdx] = jobsList.length;
            }
          }
          modified = true;
        }
      });
      
      if (modified) {
        sh.clearContents();
        sh.getRange(1, 1, dataValues.length, dataValues[0].length).setValues(dataValues);
      }
    }
  });
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

function cleanVal(v) {
  if (v === null || v === undefined) return "";
  let s = String(v).trim();
  // Remove currency symbols, commas and spaces
  s = s.replace(/[₹$,]/g, "");
  // Handle scientific notation
  if (s.indexOf("E") !== -1 || s.indexOf("e") !== -1) {
    let num = Number(s);
    if (!isNaN(num)) {
      s = num.toFixed(0);
    }
  }
  // Remove decimal part if it ends with .0 or is just an integer representation
  if (s.indexOf(".") !== -1) {
    s = s.split(".")[0];
  }
  return s;
}

function moveExistingUsedLicenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data");
  const usedLicSheet = ss.getSheetByName("Used License");
  
  if (!dataSheet || !usedLicSheet) return;
  
  // Combine all rows from Data and Used License sheets (excluding headers)
  const dataRange = dataSheet.getDataRange();
  const dataValues = dataRange.getValues();
  const usedRange = usedLicSheet.getDataRange();
  const usedValues = usedRange.getValues();
  
  const header = dataValues[0];
  const headerNorm = header.map(function(h) { return String(h).trim().toUpperCase(); });
  const balIdx = headerNorm.indexOf('BALANCE');
  if (balIdx === -1) return;
  
  const allRows = [];
  const licSet = new Set();
  
  // Collect from Data sheet
  for (let i = 1; i < dataValues.length; i++) {
    const licNo = String(dataValues[i][0]).replace(".0", "").trim();
    if (licNo && !licSet.has(licNo)) {
      allRows.push(dataValues[i]);
      licSet.add(licNo);
    }
  }
  // Collect from Used License sheet
  for (let i = 1; i < usedValues.length; i++) {
    const licNo = String(usedValues[i][0]).replace(".0", "").trim();
    if (licNo && !licSet.has(licNo)) {
      allRows.push(usedValues[i]);
      licSet.add(licNo);
    }
  }
  
  const activeRows = [header];
  const completedRows = [];
  
  allRows.forEach(function(row) {
    const balance = parseFloat(row[balIdx]) || 0;
    if (balance <= 1.0) {
      completedRows.push(row);
    } else {
      activeRows.push(row);
    }
  });
  
  // Write back to Data sheet
  dataSheet.clearContents();
  dataSheet.getRange(1, 1, activeRows.length, activeRows[0].length).setValues(activeRows);
  
  // Write back to Used License sheet
  usedLicSheet.clearContents();
  const usedHeader = [header];
  usedLicSheet.getRange(1, 1, 1, header.length).setValues(usedHeader);
  if (completedRows.length > 0) {
    usedLicSheet.getRange(2, 1, completedRows.length, completedRows[0].length).setValues(completedRows);
  }
  
  // Re-apply blacklist highlights after moving/shuffling rows
  applyBlacklistHighlighting(ss);
}

function rebuildLicenseJobLists(ss, dataSheet) {
  const usedSheet = ss.getSheetByName("Used License");
  const boeSheet = ss.getSheetByName("Debit Ledger");
  if (!boeSheet) return;
  
  const boeRange = boeSheet.getDataRange();
  const boeValues = boeRange.getValues();
  const boeBackgrounds = boeRange.getBackgrounds();
  if (boeValues.length <= 1) return;
  
  const boeHeader = boeValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
  
  let boeLicIdx = -1;
  for (let i = 0; i < boeHeader.length; i++) {
    if (boeHeader[i].indexOf('LICNO') !== -1) {
      boeLicIdx = i;
      break;
    }
  }
  
  let boeJobIdx = -1;
  for (let i = 0; i < boeHeader.length; i++) {
    if (boeHeader[i].indexOf('JOB') !== -1) {
      boeJobIdx = i;
      break;
    }
  }
  
  if (boeLicIdx === -1 || boeJobIdx === -1) return;
  
  // Build a map of licNo -> Set of jobNos
  const licJobsMap = {};
  
  for (let i = 1; i < boeValues.length; i++) {
    const licNo = cleanVal(boeValues[i][boeLicIdx]);
    const jobNo = String(boeValues[i][boeJobIdx]).trim();
    const bg = boeBackgrounds[i][0];
    
    // If row is highlighted in light red (deleted/blacklisted old debit), skip it
    if (bg === "#ffebee" || licNo === "MATCHED_OLD_DEBIT" || !licNo || !jobNo) {
      continue;
    }
    
    if (!licJobsMap[licNo]) {
      licJobsMap[licNo] = {};
    }
    licJobsMap[licNo][jobNo] = true;
  }
  
  // Now update the Data and Used License sheets
  const sheets = [dataSheet, usedSheet];
  sheets.forEach(function(sh) {
    if (!sh) return;
    const dataRange = sh.getDataRange();
    const dataValues = dataRange.getValues();
    if (dataValues.length <= 1) return;
    
    const header = dataValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
    
    let licIdx = -1;
    for (let i = 0; i < header.length; i++) {
      if (header[i].indexOf('LICNO') !== -1) {
        licIdx = i;
        break;
      }
    }
    const noboeIdx = header.indexOf('COUNT OF JOB') !== -1 ? header.indexOf('COUNT OF JOB') : 6;
    const jobIdx = header.indexOf('JOB NO') !== -1 ? header.indexOf('JOB NO') : 9;
    
    if (licIdx !== -1 && jobIdx !== -1 && noboeIdx !== -1) {
      let modified = false;
      for (let i = 1; i < dataValues.length; i++) {
        const licNo = String(dataValues[i][licIdx]).replace(".0", "").trim();
        const activeJobsMap = licJobsMap[licNo] || {};
        const jobsList = Object.keys(activeJobsMap);
        
        const newJobNoVal = jobsList.join(", ");
        const newCountVal = jobsList.length;
        
        if (String(dataValues[i][jobIdx] || "").trim() !== newJobNoVal || 
            parseInt(dataValues[i][noboeIdx]) !== newCountVal) {
          dataValues[i][jobIdx] = newJobNoVal;
          dataValues[i][noboeIdx] = newCountVal;
          modified = true;
        }
      }
      if (modified) {
        sh.clearContents();
        sh.getRange(1, 1, dataValues.length, dataValues[0].length).setValues(dataValues);
      }
    }
  });
}

function applyBlacklistHighlighting(ss) {
  const dataSheet = ss.getSheetByName("Data");
  const usedSheet = ss.getSheetByName("Used License");
  const blacklistSheet = ss.getSheetByName("Blacklisted Licenses");
  if (!dataSheet || !usedSheet || !blacklistSheet) return;
  
  // Get list of all blacklisted license numbers
  const allBlacklistedLics = [];
  if (blacklistSheet.getLastRow() > 1) {
    const blValues = blacklistSheet.getRange(2, 1, blacklistSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < blValues.length; i++) {
      allBlacklistedLics.push(cleanVal(blValues[i][0]));
    }
  }
  
  [dataSheet, usedSheet].forEach(function(sh) {
    if (!sh || sh.getLastRow() <= 1) return;
    const shRange = sh.getDataRange();
    const shValues = shRange.getValues();
    const shHeader = shValues[0].map(function(h) { return String(h).trim().toUpperCase(); });
    const shLicIdx = shHeader.indexOf('LICNO/');
    
    if (shLicIdx !== -1) {
      // Clear background colors for all rows in the sheet
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).setBackground(null);
      
      // Apply red highlight only to currently blacklisted licenses
      for (let i = 1; i < shValues.length; i++) {
        const currentLic = cleanVal(shValues[i][shLicIdx]);
        if (allBlacklistedLics.indexOf(currentLic) !== -1) {
          sh.getRange(i + 1, 1, 1, sh.getLastColumn()).setBackground("#ffebee");
        }
      }
    }
  });
}
```

3. Click the **Save** icon (floppy disk) to save the script.

---

## Step 4: Deploy as a Web App
1. In the Apps Script window, click the **Deploy** button at the top right, then select **New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Configure the settings:
   - **Description:** `Skoda License Web App Proxy v3`
   - **Execute as:** `Me (your-email@gmail.com)`
   - **Who has access:** `Anyone`
4. Click **Deploy**.
5. Once deployed, copy the new **Web app URL** and save it in your Database Master configuration tab.
