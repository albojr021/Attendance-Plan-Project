// --- CONFIGURATION ---
const SPREADSHEET_ID = '1rQnJGqcWcEBjoyAccjYYMOQj7EkIu1ykXTMLGFzzn2I';
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; 
const CONTRACTS_SHEET_NAME = 'MASTER';

const MASTER_HEADER_ROW = 5;
const PLAN_HEADER_ROW = 6;

// --- UPDATED CONFIGURATION FOR PRINT LOG (10 Columns) ---
const LOG_SHEET_NAME = 'PrintLog';
const LOG_HEADERS = [
    'Reference #', 
    'SFC Ref#', 
    'Plan Sheet Name',      // NEW: For precise locking key
    'Plan Period Display',  // UPDATED: Original 'Plan Period'
    'Payor Company', 
    'Agency', 
    'Service Type', 
    'User Email', 
    'Timestamp',
    'Locked Personnel IDs'  // NEW: Comma-separated list of IDs
];
// ---------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Attendance Plan Monitor');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function sanitizeHeader(header) {
    if (!header) return '';
    return String(header).replace(/[^A-Za-z0-9]/g, '');
}

function getSheetData(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId); 
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  let startRow = 1;
  let numRows = sheet.getLastRow();
  let numColumns = sheet.getLastColumn();
  if (sheetName === CONTRACTS_SHEET_NAME) {
    startRow = MASTER_HEADER_ROW;
    if (numRows < startRow) {
      Logger.log(`[getSheetData] MASTER sheet has no data starting from Row ${startRow}.`);
      return [];
    }
    numRows = sheet.getLastRow() - startRow + 1;
  } 
  else if (sheetName.includes('Employees')) {
      startRow = 1;
  }
  else if (sheetName.includes('AttendancePlan')) {
      startRow = PLAN_HEADER_ROW;
    if (numRows < startRow) {
          numRows = 0;
    } else {
          numRows = sheet.getLastRow() - startRow + 1;
    }
  }

  if (numRows <= 0 || numColumns === 0) return [];
  const range = sheet.getRange(startRow, 1, numRows, numColumns);
  // Gumagamit ng getDisplayValues() para sa data consistency
  const values = range.getDisplayValues();
  const headers = values[0];
  const cleanHeaders = headers.map(header => (header || '').toString().trim());

  const data = [];
  for (let i = 1; i < values.length; i++) { 
    const row = values[i];
    if (sheetName === CONTRACTS_SHEET_NAME || sheetName.includes('Employees')) {
      if (!row.some(cell => String(cell).trim() !== '')) {
        continue;
      }
    }
    
    const item = {};
    cleanHeaders.forEach((headerKey, index) => {
      if (headerKey) {
          item[headerKey] = row[index];
      }
    });
    data.push(item);
  }
  
  Logger.log(`[getSheetData] Total data rows processed (excluding header): ${data.length}`);
  return data;
}

function getDynamicSheetName(sfcRef, type, year, month, shift) {
    const safeRef = (sfcRef || '').replace(/[\\/?*[]/g, '_');
    if (type === 'employees') {
        return `${safeRef} - Employees`;
    }
    
    if (type === 'plan' && year !== undefined && month !== undefined && shift) {
        const tempDate = new Date(year, month, 1);
        const monthName = tempDate.toLocaleString('en-US', { month: 'short' });
        return `${safeRef} - ${monthName} ${year} - ${shift} AttendancePlan`;
    }
    
    return `${safeRef} - AttendancePlan`;
}

function checkContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef || year === undefined || month === undefined || !shift) return false;
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
        const empSheetName = getDynamicSheetName(sfcRef, 'employees');
        const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
        return !!ss.getSheetByName(empSheetName) && !!ss.getSheetByName(planSheetName);
    } catch (e) {
         Logger.log(`[checkContractSheets] ERROR: Failed to open Spreadsheet ID ${TARGET_SPREADSHEET_ID}. Check ID and permissions. Error: ${e.message}`);
         return false;
    }
}

function appendExistingEmployeeRowsToPlan(sfcRef, planSheet, shiftToAppend) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empSheet = ss.getSheetByName(empSheetName);

    if (!empSheet) return;
    const empData = getSheetData(TARGET_SPREADSHEET_ID, empSheetName);
    const existingIds = empData.map(e => cleanPersonnelId(e['Personnel ID'])).filter(id => id);

    if (existingIds.length === 0) return;
    Logger.log(`[appendExistingEmployeeRowsToPlan] Found ${existingIds.length} existing employees. Populating plan sheet.`);

    const planHeadersCount = 33;
    const rowsToAppend = [];
    existingIds.forEach(id => {
        if (shiftToAppend === '1stHalf') {
            const planRow1 = Array(planHeadersCount).fill('');
            planRow1[0] = id; 
            planRow1[1] = '1stHalf'; 
            rowsToAppend.push(planRow1);
        }
        if (shiftToAppend === '2ndHalf') {
         
            const planRow2 = Array(planHeadersCount).fill('');
            planRow2[0] = id; 
            planRow2[1] = '2ndHalf'; 
            rowsToAppend.push(planRow2);
        }
    });
    if (rowsToAppend.length > 0) {
        planSheet.getRange(planSheet.getLastRow() + 1, 1, rowsToAppend.length, planHeadersCount).setValues(rowsToAppend);
        Logger.log(`[appendExistingEmployeeRowsToPlan] Successfully pre-populated ${rowsToAppend.length} plan rows for ${shiftToAppend}.`);
    }
}

function createContractSheets(sfcRef, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    // --- EMPLOYEES SHEET ---
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    let empSheet = ss.getSheetByName(empSheetName);
    if (!empSheet) {
        empSheet = ss.insertSheet(empSheetName);
        empSheet.clear();
        const empHeaders = ['Personnel ID', 'Personnel Name', 'Position', 'Area Posting'];
        empSheet.getRange(1, 1, 1, empHeaders.length).setValues([empHeaders]);
        empSheet.setFrozenRows(1);
        Logger.log(`[createContractSheets] Created Employee sheet for ${sfcRef}`);
    } 
    
    // --- ATTENDANCE PLAN SHEET ---
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    let planSheet = ss.getSheetByName(planSheetName);

    const getHorizontalPlanHeaders = (sheetYear, sheetMonth) => {
        const base = ['Personnel ID', 'Shift'];
        for (let d = 1; d <= 31; d++) {
            const currentDate = new Date(sheetYear, sheetMonth, d);
            if (currentDate.getMonth() === sheetMonth) {
                const monthShortRaw = currentDate.toLocaleString('en-US', { month: 'short' });
                const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '');
                base.push(`${monthShort}${d}`); 
            } else {
                base.push(`Day${d}`);
            }
        }
        return base;
    };
    if (!planSheet) {
        planSheet = ss.insertSheet(planSheetName);
        planSheet.clear();
        const planHeaders = getHorizontalPlanHeaders(year, month);
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, planHeaders.length).setValues([planHeaders]);
        planSheet.setFrozenRows(PLAN_HEADER_ROW); 
        
        Logger.log(`[createContractSheets] Created Horizontal Attendance Plan sheet for ${planSheetName} with headers at Row ${PLAN_HEADER_ROW}.`);
        appendExistingEmployeeRowsToPlan(sfcRef, planSheet, shift);
        
        if (shift === '1stHalf') {
            const START_COL_TO_HIDE = 18;
            const NUM_COLS_TO_HIDE = 16; 
            planSheet.hideColumns(START_COL_TO_HIDE, NUM_COLS_TO_HIDE);
            Logger.log(`[createContractSheets] Hiding Day 16-31 columns for 1stHalf sheet.`);
        } else if (shift === '2ndHalf') {
            const START_COL_TO_HIDE = 3;
            const NUM_COLS_TO_HIDE = 15; 
            planSheet.hideColumns(START_COL_TO_HIDE, NUM_COLS_TO_HIDE);
            Logger.log(`[createContractSheets] Hiding Day 1-15 columns for 2ndHalf sheet.`);
        }
    } 
}

function ensureContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required to ensure sheets.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    if (!ss.getSheetByName(empSheetName)) {
        createContractSheets(sfcRef, year, month, shift);
        Logger.log(`[ensureContractSheets] Created Employee sheet for ${sfcRef}.`);
    }
    
    if (!ss.getSheetByName(planSheetName)) {
        createContractSheets(sfcRef, year, month, shift);
        Logger.log(`[ensureContractSheets] Created new Plan sheet for ${planSheetName}.`);
    } 
}

function getContracts() {
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE' || !SPREADSHEET_ID) {
    throw new Error("CONFIGURATION ERROR: Pakipalitan ang 'YOUR_SPREADSHEET_ID_HERE' sa Code.gs ng tamang Spreadsheet ID.");
  }
    
  const allContracts = getSheetData(SPREADSHEET_ID, CONTRACTS_SHEET_NAME);
  const findKey = (c, search) => {
      const keys = Object.keys(c);
      return keys.find(key => (key || '').trim().toLowerCase() === search.toLowerCase());
  };
    
  const filteredContracts = allContracts.filter((c) => {
    const statusKey = findKey(c, 'Status of SFC');
    const contractIdKey = findKey(c, 'CONTRACT GRP ID');
    
    if (!statusKey || !contractIdKey) return false;

    const contractIdValue = (c[contractIdKey] || '').toString().trim();
    if (!contractIdValue) return false; 
    
    const status = (c[statusKey] || '').toString().trim().toLowerCase();
    const isLive = status === 'live' || status === 'on process - live';

    return isLive;
  
  });

  return filteredContracts.map(c => {
    
    const contractIdKey = findKey(c, 'CONTRACT GRP ID');
    const statusKey = findKey(c, 'Status of SFC');
    const payorKey = findKey(c, 'PAYOR COMPANY NAME');
    const agencyKey = findKey(c, 'PAYEE/ SUPPLIER/ SERVICE PROVIDER COMPANY/ AGENCY NAME');
    const serviceTypeKey = findKey(c, 'SERVICE TYPE');
    const headCountKey = findKey(c, 'TOTAL HEAD COUNT');
    const sfcRefKey = findKey(c, 'SFC Ref#');
      
    return {
      id: 
      contractIdKey ? 
      (c[contractIdKey] || '').toString() : '',     
      status: statusKey ? (c[statusKey] || '').toString() : '',   
      payorCompany: payorKey ? (c[payorKey] || '').toString() : '', 
      agency: agencyKey ? (c[agencyKey] || '').toString() : '',       
      serviceType: serviceTypeKey ? (c[serviceTypeKey] || '').toString() : '',   
      headCount: parseInt(headCountKey ? c[headCountKey] : 0) || 0, 
   
      sfcRef: sfcRefKey ? (c[sfcRefKey] || 
      '').toString() : '', 
    };
  });
}

function cleanPersonnelId(rawId) {
    let idString = String(rawId || '').trim();
    // Tiyakin na numbers lang at tanggalin ang space/comma
    return idString.replace(/\D/g, '');
}

/**
 * UPDATED FUNCTION: Kumuha ng map ng locked IDs at ang Reference # na nag-lock sa kanila.
 * Returns: { 'Personnel ID': 'Reference #' }
 */
function getLockedPersonnelIds(ss, planSheetName) {
    const logSheet = getOrCreateLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return {}; 

    // Define column indices (based on the 10-column LOG_HEADERS)
    const REF_NUM_COL = 1;              // Col A
    const PLAN_SHEET_NAME_COL = 3;      // Col C
    const LOCKED_IDS_COL = LOG_HEADERS.length; // Col J (10)

    // Basahin ang lahat ng data mula Col A (1) hanggang Col J (10)
    // Gamit ang getDisplayValues() para makuha ang string na value (FIX)
    const values = logSheet.getRange(2, 1, lastRow - 1, LOCKED_IDS_COL).getDisplayValues();

    const lockedIdRefMap = {}; 

    values.forEach(row => {
        const refNum = String(row[REF_NUM_COL - 1] || '').trim(); // Col A (index 0)
        const planSheetNameInLog = String(row[PLAN_SHEET_NAME_COL - 1] || '').trim(); // Col C (index 2)
        const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim(); // Col J (index 9)
        
        if (planSheetNameInLog === planSheetName && lockedIdsString) {
            
            // FIX for corrupted reading: Strip non-numeric/non-comma characters for safety, then split.
            const cleanIdsString = lockedIdsString.replace(/[^0-9,]/g, '');
            
            // Split and check each ID
            cleanIdsString.split(',').forEach(id => {
                const cleanId = cleanPersonnelId(id);
                // CRITICAL CHECK: Tiyakin na ang ID ay valid (minimum 7 digits for safety).
                if (cleanId.length >= 7) { 
                     // Store the ID and its corresponding Reference Number
                     // We use the first Ref # found for this ID as the source of truth
                     if (!lockedIdRefMap[cleanId]) { 
                        lockedIdRefMap[cleanId] = refNum;
                     }
                }
            });
        }
    });
        
    return lockedIdRefMap; // Return map { 'ID': 'RefNum' }
}

function getAttendancePlan(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empData = getSheetData(TARGET_SPREADSHEET_ID, empSheetName);
    
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const planSheet = ss.getSheetByName(planSheetName);
    
    // NEW: Kunin ang map ng locked IDs at Reference #
    const lockedIdRefMap = getLockedPersonnelIds(ss, planSheetName);
    const lockedIds = Object.keys(lockedIdRefMap); 
    
    if (!planSheet) return { employees: empData, planMap: {}, lockedIds: lockedIds, lockedIdRefMap: lockedIdRefMap }; 
    
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();
    if (numRowsToRead <= 0 || numColumns < 33) { 
        return { employees: empData, planMap: {}, lockedIds: lockedIds, lockedIdRefMap: lockedIdRefMap };
    }

    // CRITICAL FIX: Use getDisplayValues() to ensure time formats (08:00-17:00) are read as strings.
    const planValues = planSheet.getRange(HEADER_ROW, 1, numRowsToRead + 1, numColumns).getDisplayValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const shiftIndex = headers.indexOf('Shift');
    
    const planMap = {};
    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        const cleanedHeader = sanitizeHeader(header); 
        sanitizedHeadersMap[cleanedHeader] = index;
    });
    dataRows.forEach((row, rowIndex) => {
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        const currentShift = String(row[shiftIndex] || '').trim();
        
        if (currentShift === shift) {
            
            for (let d = 1; d <= 31; d++) {
         
                const dayKey = `${year}-${month + 1}-${d}`; 
                const date = new Date(year, month, d);
                
                let lookupHeader = '';
                if (date.getMonth() === month) {
 
                    const monthShortRaw = date.toLocaleString('en-US', { month: 'short' });
                    const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '').replace(/\s/g, '');
                    lookupHeader = `${monthShort}${d}`;
               
                } else {
                
                    lookupHeader = `Day${d}`;
                }
                
                const dayIndex = sanitizedHeadersMap[sanitizeHeader(lookupHeader)];
                if (dayIndex !== undefined && id && currentShift) {
                    // CRITICAL FIX: Direct access from the values array (already using getDisplayValues)
                    const status = String(row[dayIndex] || '').trim();
                    const key = `${id}_${dayKey}_${currentShift}`;
                    if (status) {
                         planMap[key] = status;
                    }
                }
            }
        }
    });
    const employees = empData.map((e, index) => {
        const id = cleanPersonnelId(e['Personnel ID']);
        return {
           no: index + 1, 
            id: id, 
            name: String(e['Personnel Name'] || '').trim(),
            position: String(e['Position'] || '').trim(),
            area: 
            String(e['Area Posting'] || '').trim(),
        }
    }).filter(e => e.id);
    return { employees, planMap, lockedIds: lockedIds, lockedIdRefMap: lockedIdRefMap }; 
}

function updatePlanKeysOnIdChange(sfcRef, employeeChanges) {
    Logger.log("[updatePlanKeysOnIdChange] Skipped Plan Sheet ID update for ID change due to new dynamic naming.");
}

function saveAllData(sfcRef, contractInfo, employeeChanges, attendanceChanges, year, month, shift) {
    Logger.log(`[saveAllData] Starting save for SFC Ref#: ${sfcRef}, Month/Shift: ${month}/${shift}`);
    if (!sfcRef) {
      throw new Error("SFC Ref# is required.");
    }
    ensureContractSheets(sfcRef, year, month, shift);
    
    // NEW: Check for locked IDs before saving
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const lockedIds = Object.keys(getLockedPersonnelIds(ss, planSheetName)); // Use keys from the map

    // Filter out changes for locked IDs
    const finalEmployeeChanges = employeeChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.id || change.oldPersonnelId);
        if (lockedIds.includes(idToCheck) && !change.isDeleted) {
            Logger.log(`[saveAllData] Skipping employee info update for locked ID: ${idToCheck}`);
            return false;
        }
        return true;
    });

    const finalAttendanceChanges = attendanceChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.personnelId);
        if (lockedIds.includes(idToCheck)) {
            Logger.log(`[saveAllData] Skipping attendance plan update for locked ID: ${idToCheck}`);
            return false;
        }
        return true;
    });

    // Proceed to save with filtered changes
    saveContractInfo(sfcRef, contractInfo, year, month, shift);
    if (finalEmployeeChanges && finalEmployeeChanges.length > 0) {
        saveEmployeeInfoBulk(sfcRef, finalEmployeeChanges, year, month, shift);
    }
    
    if (finalAttendanceChanges && finalAttendanceChanges.length > 0) {
        saveAttendancePlanBulk(sfcRef, finalAttendanceChanges, year, month, shift);
    }
    
    Logger.log(`[saveAllData] Save completed.`);
}

function saveContractInfo(sfcRef, info, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const planSheet = ss.getSheetByName(planSheetName);
    if (!planSheet) throw new Error(`Plan Sheet for ${planSheetName} not found.`);
    const date = new Date(year, month, 1);
    const monthName = date.toLocaleString('en-US', { month: 'long' });
    const yearNum = date.getFullYear();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    let dateRange = '';
    if (shift === '1stHalf') {
        dateRange = `${monthName} 1-15, ${yearNum} (${shift})`;
    } else {
        dateRange = `${monthName} 16-${daysInMonth}, ${yearNum} (${shift})`;
    }
    
    const data = [
        ['PAYOR COMPANY', info.payor],          
        ['AGENCY', info.agency],                
        ['SERVICE TYPE', info.serviceType],     
        ['TOTAL HEAD COUNT', info.headCount],    
        ['PLAN PERIOD', dateRange]    
    ];
    planSheet.getRange('A1:B5').clearContent();
    planSheet.getRange('A1:B5').setValues(data);
    planSheet.setFrozenRows(PLAN_HEADER_ROW);
    Logger.log(`[saveContractInfo] Saved metadata and date range for ${planSheetName}.`);
}

function saveAttendancePlanBulk(sfcRef, changes, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const planSheet = ss.getSheetByName(planSheetName);
    if (!planSheet) throw new Error(`AttendancePlan Sheet for ${planSheetName} not found.`);
    const HEADER_ROW = PLAN_HEADER_ROW;
    
    planSheet.setFrozenRows(0);
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();
    
    let values = [];
    let headers = [];
    // CRITICAL FIX: Use getDisplayValues() to get string headers (e.g., 'Nov1')
    if (numRowsToRead >= 0 && numColumns > 0) {
         values = planSheet.getRange(HEADER_ROW, 1, numRowsToRead + 1, numColumns).getDisplayValues();
        headers = values[0]; 
    } else {
        headers = ['Personnel ID', 'Shift'];
        for (let i = 1; i <= 31; i++) {
            headers.push(`Day${i}`);
        }
        values.push(headers);
    }
    
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const shiftIndex = headers.indexOf('Shift');
    
    if (personnelIdIndex === -1 || shiftIndex === -1) {
        throw new Error("Missing critical column in AttendancePlan sheet (Personnel ID or Shift).");
    }
    
    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        const cleanedHeader = sanitizeHeader(header); 
        sanitizedHeadersMap[cleanedHeader] = index;
    });
    const rowLookupMap = {};
    for (let i = 1; i < values.length; i++) { 
        const rowId = String(values[i][personnelIdIndex] || '').trim();
        const rowShift = String(values[i][shiftIndex] || '').trim();
        if (rowId) {
            rowLookupMap[`${rowId}_${rowShift}`] = i;
        }
    }
    
    const updatesMap = {};
    changes.forEach(data => {
        const { personnelId, dayKey, shift: dataShift, status } = data;
        
        if (dataShift !== shift) return; 
        
        const rowKey = `${personnelId}_${dataShift}`;
        
        const dayNumber = parseInt(dayKey.split('-')[2], 10);
        const date = new Date(year, month, dayNumber);
       
        let targetLookupHeader = '';
        if (date.getMonth() === month) {
            const monthShortRaw = date.toLocaleString('en-US', { month: 'short' });
            const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '').replace(/\s/g, '');
            targetLookupHeader = `${monthShort}${dayNumber}`; 
        } else {
         
            targetLookupHeader 
            = `Day${dayNumber}`;
        }
        
        const dayColIndex = sanitizedHeadersMap[sanitizeHeader(targetLookupHeader)];
        if (dayColIndex === undefined) {
            Logger.log(`[savePlanBulk] FATAL MISS: Header Lookup '${targetLookupHeader}' failed.
            Available Sanitized Keys: ${Object.keys(sanitizedHeadersMap).join(' | ')}`);
            return; 
        }

        const rowIndexInValues = rowLookupMap[rowKey];
        if (rowIndexInValues !== undefined) {
            const sheetRowNumber = rowIndexInValues + HEADER_ROW;
            if (!updatesMap[sheetRowNumber]) {
                updatesMap[sheetRowNumber] = {};
            }
            updatesMap[sheetRowNumber][dayColIndex + 1] = status;
        } else {
            Logger.log(`[savePlanBulk] WARNING: ID/Shift combination not found for row: ${rowKey}. Skipping update for this cell.`);
        }
    });

    Object.keys(updatesMap).forEach(sheetRowNumber => {
        const colUpdates = updatesMap[sheetRowNumber];
        
        Object.keys(colUpdates).forEach(colNum => {
            const status = colUpdates[colNum];
            planSheet.getRange(parseInt(sheetRowNumber), parseInt(colNum)).setValue(status);
        });
    });
    planSheet.setFrozenRows(HEADER_ROW); 
    Logger.log(`[saveAttendancePlanBulk] Completed Horizontal Plan update for ${planSheetName}.`);
}

function saveEmployeeInfoBulk(sfcRef, changes, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empSheet = ss.getSheetByName(empSheetName);
    
    // Kunin ang parehong plan sheet para sa 1st at 2nd half, kailangan ito para sa deletion
    const planSheetName1st = getDynamicSheetName(sfcRef, 'plan', year, month, '1stHalf');
    const planSheet1st = ss.getSheetByName(planSheetName1st);
    const planSheetName2nd = getDynamicSheetName(sfcRef, 'plan', year, month, '2ndHalf'); 
    const planSheet2nd = ss.getSheetByName(planSheetName2nd);
    if (!empSheet) throw new Error(`Employee Sheet for SFC Ref# ${sfcRef} not found.`);
    empSheet.setFrozenRows(0);
    
    const numRows = empSheet.getLastRow() > 0 ?
    empSheet.getLastRow() : 1;
    const numColumns = empSheet.getLastColumn() > 0 ? empSheet.getLastColumn() : 4;
    const values = empSheet.getRange(1, 1, numRows, numColumns).getValues();
    const headers = values[0]; 
    empSheet.setFrozenRows(1); 
    
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('Position');
    const areaIndex = headers.indexOf('Area Posting');

    const rowsToUpdate = {};
    const rowsToAppend = [];
    const rowsToDelete = [];
    // Bagong array para sa mga row na ide-delete
    
    const planRowsToAppend1st = [];
    const planRowsToAppend2nd = [];
    const personnelIdMap = {}; // Map: Personnel ID -> Sheet Row Number (1-based)
    for (let i = 1; i < values.length; i++) { 
        personnelIdMap[String(values[i][personnelIdIndex] || '').trim()] = i + 1;
    }
    
    changes.forEach((data) => {
        const oldId = String(data.oldPersonnelId || '').trim();
        const newId = String(data.id || '').trim();
        
        if (data.isDeleted && oldId && personnelIdMap[oldId]) {
            // MARK FOR DELETION
            rowsToDelete.push({ 
              
                rowNum: personnelIdMap[oldId], 
                id: oldId 
            });
            delete personnelIdMap[oldId]; // I-alis sa map
            return;
        }

        if (!data.isNew && personnelIdMap[oldId]) { 
  
           const sheetRowNumber = 
             personnelIdMap[oldId];
          
            if(oldId !== newId && personnelIdMap[newId] && personnelIdMap[newId] !== sheetRowNumber) return; 
            
            rowsToUpdate[sheetRowNumber] = [newId, data.name, data.position, data.area];
            
            if (oldId !== newId) {
                delete personnelIdMap[oldId];
      
                personnelIdMap[newId] = sheetRowNumber;
            }

        } 
        else if (data.isNew) {
       
             if (personnelIdMap[newId]) return;
            const newRow = [];
  
            newRow[personnelIdIndex] = data.id;
            newRow[nameIndex] = data.name;
            newRow[positionIndex] = data.position;
            newRow[areaIndex] = data.area;
            
            const finalRow = [];
            for(let i = 0; i < headers.length; i++) {
                finalRow.push(newRow[i] !== undefined ? newRow[i] : '');
            }
            
            rowsToAppend.push(finalRow);
            const planHeadersCount = 33;
            
            const planRow1 = Array(planHeadersCount).fill('');
            planRow1[0] = newId; 
            planRow1[1] = '1stHalf';
            planRowsToAppend1st.push(planRow1);
            
            const planRow2 = Array(planHeadersCount).fill('');
            planRow2[0] = newId; 
            planRow2[1] = '2ndHalf';
            planRowsToAppend2nd.push(planRow2);
            
            personnelIdMap[newId] = -1;
            // Temporary marker
        }
    });
    // 1. I-delete ang mga row sa Employee Sheet (mula sa huli para hindi magulo ang row number)
    rowsToDelete.sort((a, b) => b.rowNum - a.rowNum);
    rowsToDelete.forEach(item => {
        // Tiyakin na ang row number ay tama (laging 1-based, at ang 1st row ay header)
        if (item.rowNum > 1) { 
            empSheet.deleteRow(item.rowNum);
            Logger.log(`[saveEmployeeInfoBulk] Deleted Employee row ${item.rowNum} for ID ${item.id}.`);
            
            // I-delete ang Plan rows sa magkabilang 
            const deletePlanRow = (planSheet) => {
                if (planSheet) {
                    const planValues = planSheet.getRange(PLAN_HEADER_ROW + 1, 1, planSheet.getLastRow() - PLAN_HEADER_ROW, 2).getValues();
                    for(let i = planValues.length - 1; i >= 0; i--) {
  
   
                        if (String(planValues[i][0] || '').trim() === item.id) {
                            planSheet.deleteRow(PLAN_HEADER_ROW + i + 1);
                            // Dahil isang ID/Shift lang ang mayroon, puwede tayong mag-break.
   
 
                           // Huwag mag-break, dahil may 1stHalf at 2ndHalf na shift na may parehong ID.
                        }
                    }
                }
            };
            deletePlanRow(planSheet1st);
            deletePlanRow(planSheet2nd);
            Logger.log(`[saveEmployeeInfoBulk] Deleted Plan rows for ID ${item.id} in both shifts.`);
        }
    });
    // I-reload ang map ng Employee Sheet after deletion
    const currentEmpData = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, numColumns).getValues();
    const updatedPersonnelIdMap = {};
    currentEmpData.forEach((row, index) => {
        updatedPersonnelIdMap[String(row[personnelIdIndex] || '').trim()] = index + 2;
    });
    // 2. Update existing rows (Employee Sheet)
    Object.keys(rowsToUpdate).forEach(sheetRowNumber => {
        const rowData = rowsToUpdate[sheetRowNumber];
        // Gamitin ang updated map para mahanap ang ROW number kung nag-iba ang ID (mas kumplikado, kaya UPDATE lang natin ang content)
        // Ito ang original logic na tama sa content update:
        empSheet.getRange(parseInt(sheetRowNumber), personnelIdIndex + 1, 1, 4).setValues([
            [rowData[0], rowData[1], rowData[2], rowData[3]]
   
         ]);
        
        // NOTE: Kung nagbago ang ID, hindi na natin kailangang i-update ang plan sheet,
        // dahil ang `getAttendancePlan` ay gumagamit ng Personnel ID, kaya kapag ni-load ulit,
        // ang lumang plan row ay hindi na lalabas at ang bagong ID ay wala pa ring plan (blanko).
        // PERO! Kung ang ID ay NA-UPDATE, dapat 
        // nating i-clear ang lumang 
        // Dahil sa pagiging kumplikado at risk sa data integrity, pansamantala ay mananatili tayong
        // HINDI nag-u-update ng ID sa Employee Sheet para sa saved rows.
        // *Ang update ng ID ay i-a-allow lang para sa NEW rows na hindi pa na-save.*
        
    });
    // 3. Append new rows (Employee Sheet)
    if (rowsToAppend.length > 0) {
      rowsToAppend.forEach(row => {
          empSheet.appendRow(row); 
      });
    }
    
    // 4. Append new rows (Attendance Plan Sheet - 1st Half)
    if (planRowsToAppend1st.length > 0 && planSheet1st) {
        planSheet1st.getRange(planSheet1st.getLastRow() + 1, 1, planRowsToAppend1st.length, planRowsToAppend1st[0].length).setValues(planRowsToAppend1st);
    }
    
    // 5. Append new rows (Attendance Plan Sheet - 2nd Half)
    if (planRowsToAppend2nd.length > 0 && planSheet2nd) {
        planSheet2nd.getRange(planSheet2nd.getLastRow() + 1, 1, planRowsToAppend2nd.length, planRowsToAppend2nd[0].length).setValues(planRowsToAppend2nd);
    }
}


// --- NEW FUNCTION: LOGGING & LOCKING ACTION ---

function getOrCreateLogSheet(ss) {
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (sheet) {
        return sheet;
    }
    
    try {
        // Subukan i-insert.
        sheet = ss.insertSheet(LOG_SHEET_NAME);
        // Set Headers at Row 1
        sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, LOG_HEADERS.length, 120); // Set column width for readability
        Logger.log(`[getOrCreateLogSheet] Created Log sheet: ${LOG_SHEET_NAME}`);
        return sheet;
        
    } catch (e) {
        // Fallback para sa transient error (tulad ng na-experience mo)
        if (e.message.includes(`sheet with the name "${LOG_SHEET_NAME}" already exists`)) {
             Logger.log(`[getOrCreateLogSheet] WARN: Transient sheet creation failure, retrieving existing sheet.`);
             return ss.getSheetByName(LOG_SHEET_NAME);
        }
        throw e; // I-re-throw ang iba pang unexpected errors
    }
}

function getNextReferenceNumber(logSheet) {
    const lastRow = logSheet.getLastRow();
    // Start at 1 if only header row exists
    if (lastRow < 2) return 1;
    // Read the last value in Column A (Reference #)
    const lastRef = logSheet.getRange(lastRow, 1).getValue();
    // Attempt to parse and increment, defaulting to 1 if it's not a number
    return (parseInt(lastRef) || 0) + 1;
}

/**
 * Logs the print action to the PrintLog sheet and generates an incrementing reference number.
 * @param {string} sfcRef 
 * @param {object} contractInfo 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @returns {number} The generated Reference Number.
 */
function logPrintAction(sfcRef, contractInfo, year, month, shift) {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);
        
        const nextRefNum = getNextReferenceNumber(logSheet);
        
        Logger.log(`[logPrintAction] Generated Print Reference Number: ${nextRefNum} for ${sfcRef}.`);
        return nextRefNum;
} catch (e) {
        Logger.log(`[logPrintAction] FATAL ERROR: ${e.message}`);
        throw new Error(`Failed to generate print reference number. Error: ${e.message}`);
    }
}


/**
 * UPDATED FUNCTION: Records the actual print log entry using the pre-generated Ref # and locks printed IDs.
 * Ginamitan ng setNumberFormat('@') FIX.
 * * @param {number} refNum 
 * @param {string} sfcRef 
 * @param {object} contractInfo 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @param {string[]} printedPersonnelIds The list of IDs successfully printed.
 */
function recordPrintLogEntry(refNum, sfcRef, contractInfo, year, month, shift, printedPersonnelIds) {
    if (!refNum) {
        Logger.log(`[recordPrintLogEntry] ERROR: No Reference Number provided.`);
        return;
    }
    
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);

        // Get Plan Sheet Name (used for locking)
        const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);

        // Get Plan Period Display string (used for display)
        const date = new Date(year, month, 1);
        const monthName = date.toLocaleString('en-US', { month: 'long' });
        const yearNum = date.getFullYear();
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        
        let dateRange = '';
        if (shift === '1stHalf') {
            dateRange = `${monthName} 1-15, ${yearNum} (${shift})`;
        } else {
            dateRange = `${monthName} 16-${daysInMonth}, ${yearNum} (${shift})`;
        }
        
        // 1. Prepare the Log Entry (10 Columns)
        const logEntry = [
            refNum, 
            sfcRef,
            planSheetName, 
            dateRange,     
            contractInfo.payor,
            contractInfo.agency,
            contractInfo.serviceType,
            Session.getActiveUser().getEmail(), 
            new Date(),
            printedPersonnelIds.join(',') // Col J: IDs joined by comma
        ];
        
        const lastLoggedRow = logSheet.getLastRow();
        const newRow = lastLoggedRow + 1;
        const LOCKED_IDS_COL = LOG_HEADERS.length; // 10

        const logEntryRange = logSheet.getRange(newRow, 1, 1, LOG_HEADERS.length);

        // 2. *** CRITICAL FIX: I-set ang format ng Locked Personnel IDs cell (Col J) sa Plain Text ('@') ***
        logEntryRange.getCell(1, LOCKED_IDS_COL).setNumberFormat('@'); 

        // 3. Write the whole row of data
        logEntryRange.setValues([logEntry]);
        
        // 4. Final Touches
        logSheet.getRange(newRow, 1, 1, LOG_HEADERS.length).setHorizontalAlignment('left');
        
        Logger.log(`[recordPrintLogEntry] Logged and Locked ${printedPersonnelIds.length} IDs for Ref# ${refNum} in ${planSheetName}.`);
        
    } catch (e) {
        Logger.log(`[recordPrintLogEntry] FATAL ERROR: Failed to log print action #${refNum}. Error: ${e.message}`);
    }
}
