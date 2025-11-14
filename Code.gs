// --- CONFIGURATION ---
const SPREADSHEET_ID = '1rQnJGqcWcEBjoyAccjYYMOQj7EkIu1ykXTMLGFzzn2I';
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; 
const CONTRACTS_SHEET_NAME = 'MASTER';

const MASTER_HEADER_ROW = 5;
const PLAN_HEADER_ROW = 6;

// ADMIN USER CONFIGURATION FOR UNLOCK FEATURE
const ADMIN_EMAILS = ['mcdmarketingstorage@megaworld-lifestyle.com'];
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
// --- NEW CONFIGURATION FOR SCHEDULE AUDIT LOG ---
const AUDIT_LOG_SHEET_NAME = 'ScheduleAuditLog';
const AUDIT_LOG_HEADERS = [
    'Timestamp', 
    'User Email', 
    'SFC Ref#', 
    'Personnel ID', 
    'Plan Sheet Name', 
    'Date (YYYY-M-D)', 
    'Shift', 
    'Old Status', 
    'New Status'
];
// ---------------------------------------

function doGet(e) {
  // Check if the request is for the direct unlock action from the email link
  if (e.parameter.action) {
    return processAdminUnlockFromUrl(e.parameter);
  }
  
  // Default action: load the main UI
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
        // Note: appendExistingEmployeeRowsToPlan adds to BOTH 1stHalf and 2ndHalf sheets if they exist.
        // This is acceptable because the client-side rendering only shows the employees that exist 
        // in the current shift's plan sheet.
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
      id: contractIdKey ? (c[contractIdKey] || '').toString() : '',     
      status: statusKey ? (c[statusKey] || '').toString() : '',   
      payorCompany: payorKey ? (c[payorKey] || '').toString() : '', 
      agency: agencyKey ? (c[agencyKey] || '').toString() : '',       
      serviceType: serviceTypeKey ? (c[serviceTypeKey] || '').toString() : '',   
      headCount: parseInt(headCountKey ? c[headCountKey] : 0) || 0, 
      sfcRef: sfcRefKey ? (c[sfcRefKey] || '').toString() : '', 
    };
  });
}

function cleanPersonnelId(rawId) {
    let idString = String(rawId || '').trim();
    // Tiyakin na numbers lang at tanggalin ang space/comma
    return idString.replace(/\D/g, '');
}

// --- NEW HELPER: Fetches the clean employee master data for auto-filling/datalist. ---
function getEmployeeMasterData(sfcRef) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    // Use the existing getSheetData to read the employee master sheet
    const masterData = getSheetData(TARGET_SPREADSHEET_ID, getDynamicSheetName(sfcRef, 'employees'));
    // Map and clean the data for client-side use
    return masterData.map((e, index) => ({
        // We use the raw ID for the datalist option value/text, but the clean ID for lookup
        id: cleanPersonnelId(e['Personnel ID']),
        name: String(e['Personnel Name'] || '').trim(),
        position: String(e['Position'] || '').trim(),
        area: String(e['Area Posting'] || '').trim(),
    })).filter(e => e.id); // Only return rows with a valid ID
}
// --- END NEW FUNCTION ---

/**
 * UPDATED FUNCTION: Kumuha ng map ng locked IDs at ang Reference # na nag-lock sa kanila.
 * Returns: { 'Personnel ID': 'Reference #' }
 */
function getLockedPersonnelIds(ss, planSheetName) {
    const logSheet = getOrCreateLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return {};
    // Define column indices (based on the 10-column LOG_HEADERS)
    const REF_NUM_COL = 1;
    // Col A
    const PLAN_SHEET_NAME_COL = 3;      // Col C
    const LOCKED_IDS_COL = LOG_HEADERS.length;
    // Col J (10)

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
                // CRITICAL CHECK: Tiyakin na ang ID ay valid (minimum 3 digits for safety).
                if (cleanId.length >= 3) { 
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
    // NOTE: empData is the FULL master employee list (used by client for datalist/lookup)
    const empData = getSheetData(TARGET_SPREADSHEET_ID, empSheetName);
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const planSheet = ss.getSheetByName(planSheetName);
    // NEW: Kunin ang map ng locked IDs at Reference #
    const lockedIdRefMap = getLockedPersonnelIds(ss, planSheetName);
    const lockedIds = Object.keys(lockedIdRefMap); 
    
    // Employee Map for quick lookup
    const empDetailMap = {};
    empData.forEach(e => {
        const id = cleanPersonnelId(e['Personnel ID']);
        if (id) {
            empDetailMap[id] = {
                 id: id, 
                 name: String(e['Personnel Name'] || '').trim(),
                 position: String(e['Position'] || '').trim(),
                 area: String(e['Area Posting'] || '').trim(),
            };
        }
    });
    if (!planSheet) return { employees: [], planMap: {}, lockedIds: lockedIds, lockedIdRefMap: lockedIdRefMap };
    // UPDATED: Return empty employees array
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();
    if (numRowsToRead <= 0 || numColumns < 33) { 
        return { employees: [], planMap: {}, lockedIds: lockedIds, lockedIdRefMap: lockedIdRefMap };
    // UPDATED: Return empty employees array
    }

    // CRITICAL FIX: Use getDisplayValues() to ensure time formats (08:00-17:00) are read as strings.
    const planValues = planSheet.getRange(HEADER_ROW, 1, numRowsToRead + 1, numColumns).getDisplayValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const shiftIndex = headers.indexOf('Shift');
    
    const planMap = {};
    const sanitizedHeadersMap = {};
    const employeesInPlan = new Set(); // NEW: Set of IDs that actually have a plan entry for this shift
    headers.forEach((header, index) => {
        const cleanedHeader = sanitizeHeader(header); 
        sanitizedHeadersMap[cleanedHeader] = index;
    });
    dataRows.forEach((row, rowIndex) => {
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        const currentShift = String(row[shiftIndex] || '').trim();
        
        if (currentShift === shift) {
            employeesInPlan.add(id); // ADD ID to the set
            
            
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
    // NEW LOGIC: Only return employee details for those found in the current plan sheet + master data
    let renderedEmployees = Array.from(employeesInPlan).map(id => {
         const emp = empDetailMap[id] || { id: id, name: '', position: '', area: '' };
         // The number (no) is no longer guaranteed to be correct here, but we will recalculate on the client side
         return {
            no: 0, // Temporarily set to 0
            id: emp.id, 
            name: emp.name,
            position: emp.position,
            area: emp.area,
         }
    }).filter(e => e.id);
    const employees = renderedEmployees.map((e, index) => { // RE-MAP AND ADD SEQUENCE NUMBER
        return {
           no: index + 1, 
            id: e.id, 
            name: e.name,
            position: e.position,
            area: e.area,
        }
    }).filter(e => e.id);
    // Filter out any empty IDs resulting from a broken map
    
    return { employees, planMap, lockedIds: lockedIds, lockedIdRefMap: lockedIdRefMap };
    // The 'employees' list is now filtered to the current plan content
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
    const lockedIds = Object.keys(getLockedPersonnelIds(ss, planSheetName));
    // Use keys from the map

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

function getOrCreateAuditLogSheet(ss) {
    let sheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
    if (sheet) {
        return sheet;
    }
    
    try {
        sheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME);
        // Set Headers at Row 1
        sheet.getRange(1, 1, 1, AUDIT_LOG_HEADERS.length).setValues([AUDIT_LOG_HEADERS]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, AUDIT_LOG_HEADERS.length, 120); // Set column width for readability
        Logger.log(`[getOrCreateAuditLogSheet] Created Audit Log sheet: ${AUDIT_LOG_SHEET_NAME}`);
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${AUDIT_LOG_SHEET_NAME}" already exists`)) {
             Logger.log(`[getOrCreateAuditLogSheet] WARN: Transient sheet creation failure, retrieving existing sheet.`);
            return ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
        }
        throw e;
    }
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
    const auditLogSheet = getOrCreateAuditLogSheet(ss);
    const userEmail = Session.getActiveUser().getEmail();

    changes.forEach(data => {
        const { personnelId, dayKey, shift: dataShift, status: newStatus } = data; // Renamed status to newStatus
        
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
      
           
     
            targetLookupHeader = `Day${dayNumber}`;
        }
        
        const dayColIndex = sanitizedHeadersMap[sanitizeHeader(targetLookupHeader)];
        const rowIndexInValues = rowLookupMap[rowKey];
        if (rowIndexInValues !== undefined && dayColIndex !== undefined) {
            const sheetRowNumber = rowIndexInValues + HEADER_ROW;
            // CRITICAL STEP 1: Get OLD Status from the original 'values' array
            // NOTE: values[0] is header row, so we use rowIndexInValues (1-based index to data row in 'values' array)
            const oldStatus = String(values[rowIndexInValues][dayColIndex] || '').trim();
            if (oldStatus !== newStatus) { // Log only if status actually changes
                
                // CRITICAL STEP 2: Log the change to the audit sheet
                const logEntry = [
                    new Date(), 
                    userEmail, 
                    sfcRef, 
                    personnelId, 
                    planSheetName, 
                    dayKey, 
                    dataShift, 
                    oldStatus, 
                    newStatus
                ];
                auditLogSheet.appendRow(logEntry);
                Logger.log(`[AuditLog] Change logged for ID ${personnelId}, Day ${dayKey}: ${oldStatus} -> ${newStatus}`);
                // CRITICAL STEP 3: Proceed with updating the updatesMap
                if (!updatesMap[sheetRowNumber]) {
                    updatesMap[sheetRowNumber] = {};
                }
                updatesMap[sheetRowNumber][dayColIndex + 1] = newStatus;
            } 
            
        } else {
            Logger.log(`[savePlanBulk] WARNING: ID/Shift combination not found or column missing for row: ${rowKey}. Skipping update for this cell.`);
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

// ********** NEW HELPER FUNCTION FOR BUG FIX **********
/**
 * Reads the Attendance Plan Sheet (ID and Shift columns only) and creates a map of existing Personnel ID_Shift keys.
 * Used to prevent duplicate rows when appending new employees.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} planSheet 
 * @returns {object} Map where key is 'PersonnelID_Shift' and value is 'true'.
 */
function getPlanKeyMap(planSheet) {
    if (!planSheet) return {};
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    // Only read data rows (below the header)
    if (lastRow <= HEADER_ROW) return {};

    // Read only the first two columns (Personnel ID and Shift)
    const values = planSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, 2).getValues();
    const planKeyMap = {};
    
    values.forEach((row) => {
        const id = String(row[0] || '').trim();
        const shift = String(row[1] || '').trim();
        if (id && shift) {
            // Store the ID_Shift key. The value being 'true' is enough to check existence.
            planKeyMap[`${id}_${shift}`] = true; 
        }
    });
    Logger.log(`[getPlanKeyMap] Found ${Object.keys(planKeyMap).length} existing ID_Shift entries in sheet: ${planSheet.getName()}`);
    return planKeyMap;
}
// *******************************************************


// ********** NEW HELPER FUNCTION FOR AUDIT LOGGING DELETED ROWS **********
/**
 * Logs all existing schedule statuses for a given employee as 'DELETED_ROW' before the row is physically deleted.
 */
function logScheduleDeletion(sfcRef, planSheet, targetShift, personnelId, userEmail, year, month) {
    if (!planSheet) return;
    const planSheetName = planSheet.getName();
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    
    if (lastRow <= HEADER_ROW) return;

    // Read headers and all data rows
    const numColumns = planSheet.getLastColumn();
    const planValues = planSheet.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, numColumns).getDisplayValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);

    const personnelIdIndex = headers.indexOf('Personnel ID');
    const shiftIndex = headers.indexOf('Shift');
    
    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        sanitizedHeadersMap[sanitizeHeader(header)] = index;
    });
    
    // Find the row to be deleted
    const targetRowIndex = dataRows.findIndex(row => 
        cleanPersonnelId(row[personnelIdIndex]) === personnelId && 
        String(row[shiftIndex] || '').trim() === targetShift
    );
    
    if (targetRowIndex === -1) return;

    const rowToDelete = dataRows[targetRowIndex];
    const auditLogSheet = getOrCreateAuditLogSheet(planSheet.getParent());
    const logEntries = [];
    
    // Iterate over days (columns)
    for (let d = 1; d <= 31; d++) {
        const dayKey = `${year}-${month + 1}-${d}`; 
        const date = new Date(year, month, d);
        let lookupHeader = '';
        
        // Determine the column header (e.g., Nov1, Dec31)
        if (date.getMonth() === month) {
            const monthShortRaw = date.toLocaleString('en-US', { month: 'short' });
            const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '').replace(/\s/g, '');
            lookupHeader = `${monthShort}${d}`;
        } else {
            lookupHeader = `Day${d}`;
        }
        
        const dayColIndex = sanitizedHeadersMap[sanitizeHeader(lookupHeader)];
        if (dayColIndex !== undefined) {
            const oldStatus = String(rowToDelete[dayColIndex] || '').trim();
            
            // Log only statuses that are NOT blank or NA (to focus on actual schedules being removed)
            if (oldStatus && oldStatus !== 'NA') {
                 const logEntry = [
                    new Date(), 
                    userEmail, 
                    sfcRef, 
                    personnelId, 
                    planSheetName, 
                    dayKey, 
                    targetShift, 
                    oldStatus, 
                    'DELETED_ROW' // New status flag for audit trail
                ];
                logEntries.push(logEntry);
            }
        }
    }
    
    if (logEntries.length > 0) {
        // Bulk append the new log entries
        auditLogSheet.getRange(auditLogSheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
        Logger.log(`[logScheduleDeletion] Logged ${logEntries.length} schedule deletions for ID ${personnelId}.`);
    }
}
// **************************************************************************


function saveEmployeeInfoBulk(sfcRef, changes, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empSheet = ss.getSheetByName(empSheetName);
    
    // Kunin ang parehong plan sheet para sa 1st at 2nd half, kailangan ito para sa deletion
    const planSheetName1st = getDynamicSheetName(sfcRef, 'plan', year, month, '1stHalf');
    const planSheet1st = ss.getSheetByName(planSheetName1st);
    const planSheetName2nd = getDynamicSheetName(sfcRef, 'plan', year, month, '2ndHalf'); 
    const planSheet2nd = ss.getSheetByName(planSheetName2nd);
    
    // ********** NEW BUG FIX LOGIC (Step 1: Get existing keys) **********
    const planKeyMap1st = getPlanKeyMap(planSheet1st);
    const planKeyMap2nd = getPlanKeyMap(planSheet2nd);
    // *******************************************************************
    
    const userEmail = Session.getActiveUser().getEmail(); // CRITICAL: Kunin ang user email dito

    if (!empSheet) throw new Error(`Employee Sheet for SFC Ref# ${sfcRef} not found.`);
    empSheet.setFrozenRows(0);
    
    const numRows = empSheet.getLastRow() > 0 ? empSheet.getLastRow() : 1;
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
    
    // UPDATED: I-declare ang mga ito, ngunit lalagyan lang ng laman ang kasalukuyang shift
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
            // CRITICAL FIX: Huwag tanggalin sa Employee Master Sheet kung may ID
            // I-delete lang ang Plan row sa 
            if (!data.isNew) {
                // For EXISTING employees, mark for deletion from the PLAN sheet only for the current shift
                rowsToDelete.push({ 
                    rowNum: -1, // Dummy row num for master sheet
                    id: oldId,
                    isMasterDelete: false // Flag to prevent master sheet deletion
                });
            } else {
                // For NEW employees (not 
                // yet saved), delete from master sheet (which is just appending logic removal)
                rowsToDelete.push({ 
                    rowNum: personnelIdMap[oldId], 
                    id: oldId,
                    isMasterDelete: true // Flag to delete from master sheet
                });
                delete personnelIdMap[oldId];
                // I-alis sa map
            }
            
            return;
        }

        // UPDATED LOGIC: Only proceed if it's a new entry (isNew)
        if (data.isNew) {
            
            // 1. Employee Master Sheet Logic (Only append if truly new to master, not just new to plan)
            if (!data.isExistingEmployeeAdded) { // If it's a truly new employee (not in master)
               
                if (personnelIdMap[newId]) return;
                // Skip if ID already exists in master list
                
                // Prepare row for Employee Master Sheet
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
                // Add to master sheet append list
                personnelIdMap[newId] = -1;
                // Mark as added to master
            }


            // 2. Attendance Plan Sheet Logic (Always append if new to plan, but check for duplicates)
            // --- FIXED LOGIC: Only append if the ID_Shift key is NOT already present ---
            const planHeadersCount = 33;
            
            if (shift === '1stHalf' && !planKeyMap1st[`${newId}_1stHalf`]) {
                const planRow1 = Array(planHeadersCount).fill('');
                planRow1[0] = newId; 
                planRow1[1] = '1stHalf';
                planRowsToAppend1st.push(planRow1);
            } else if (shift === '2ndHalf' && !planKeyMap2nd[`${newId}_2ndHalf`]) {
                const planRow2 = Array(planHeadersCount).fill('');
                planRow2[0] = newId; 
                planRow2[1] = '2ndHalf';
                planRowsToAppend2nd.push(planRow2);
            }
            // --- END FIXED LOGIC ---
        }
    });

    // 1. I-delete ang mga row sa Plan Sheet para sa kasalukuyang shift (DELETION ACTION)
    rowsToDelete.forEach(item => {
        // Function para i-delete ang row sa Plan Sheet
        const deletePlanRowForCurrentShift = (planSheet, targetShift) => {
            if (planSheet) {
                // Basahin lang ang ID at Shift columns
                const planValues = planSheet.getRange(PLAN_HEADER_ROW + 1, 1, planSheet.getLastRow() - PLAN_HEADER_ROW, 2).getValues();
                
                for(let i = planValues.length - 1; i >= 0; i--) {
                    if (String(planValues[i][0] || '').trim() === item.id && 
                        String(planValues[i][1] || '').trim() === targetShift) {
             
                        // Row index sa sheet (1-based)
                        planSheet.deleteRow(PLAN_HEADER_ROW + i + 1);
               
                        Logger.log(`[saveEmployeeInfoBulk] Deleted Plan row for ID ${item.id} in ${targetShift} shift.`);
                        return;
                    // Exit after finding and deleting the specific row
                    }
                }
            }
        };
        
        // ********** AUDIT LOG DELETION HERE **********
        // Tiyakin na mayroon tayong Year at Month para sa log
        const date = new Date(year, month, 1);
        const logYear = date.getFullYear();
        const logMonth = date.getMonth(); 
        
        // I-delete ang plan row sa kasalukuyang shift
        if (shift === '1stHalf' && planSheet1st) {
             // CRITICAL: Log deletion before physically deleting the row
             logScheduleDeletion(sfcRef, planSheet1st, '1stHalf', item.id, userEmail, logYear, logMonth);
             deletePlanRowForCurrentShift(planSheet1st, '1stHalf');
        } else if (shift === '2ndHalf' && planSheet2nd) {
             // CRITICAL: Log deletion before physically deleting the row
             logScheduleDeletion(sfcRef, planSheet2nd, '2ndHalf', item.id, userEmail, logYear, logMonth);
             deletePlanRowForCurrentShift(planSheet2nd, '2ndHalf');
        }
        // ***********************************************
        
        // Kung ito ay isang BAGONG ROW na DELETED (isMasterDelete: true), tanggalin din sa Employee Sheet
        if (item.isMasterDelete && item.rowNum > 1) {
             empSheet.deleteRow(item.rowNum);
             Logger.log(`[saveEmployeeInfoBulk] Deleted NEW Employee row ${item.rowNum} for ID ${item.id} from Master Sheet.`);
        }
    });
    // 2. Append new rows (Employee Sheet)
    if (rowsToAppend.length > 0) {
      rowsToAppend.forEach(row => {
          empSheet.appendRow(row); 
      });
    }
    
    // 3. Append new rows (Attendance Plan Sheet - 1st Half)
    if (planRowsToAppend1st.length > 0 && planSheet1st) {
        planSheet1st.getRange(planSheet1st.getLastRow() + 1, 1, planRowsToAppend1st.length, planRowsToAppend1st[0].length).setValues(planRowsToAppend1st);
    }
    
    // 4. Append new rows (Attendance Plan Sheet - 2nd Half)
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
        if (e.message.includes(`sheet with the name "${LOG_SHEET_NAME}" already exists`)) {
             Logger.log(`[getOrCreateLogSheet] WARN: Transient sheet creation failure, retrieving existing sheet.`);
            return ss.getSheetByName(LOG_SHEET_NAME);
        }
        throw e;
    }
}

/**
 * Finds the next sequential reference number by scanning all existing entries.
 * This ensures the number never resets, even if the last row was deleted.
 */
function getNextReferenceNumber(logSheet) {
    const lastRow = logSheet.getLastRow();
    // Start at 1 if only header row exists
    if (lastRow < 2) return 1;
    // Read ALL values from Column A (Reference #) starting from Row 2 (skipping header)
    const range = logSheet.getRange(2, 1, lastRow - 1, 1);
    const refNumbers = range.getValues();
    
    let maxRef = 0;
    
    refNumbers.forEach(row => {
        // Parse the value, defaulting to 0 if invalid
        const currentRef = parseInt(row[0]) || 0;
        if (currentRef > maxRef) {
            maxRef = currentRef;
        }
    });
    // Return the maximum reference number found, incremented by 1
    return maxRef + 1;
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
        const LOCKED_IDS_COL = LOG_HEADERS.length;
        // 10

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


/**
 * Sends a notification email to the original requester about the status of their unlock request.
 * @param {string} status - 'APPROVED' or 'REJECTED'
 * @param {string} personnelId
 * @param {string} requesterEmail
 */
function sendRequesterNotification(status, personnelId, requesterEmail) {
  if (requesterEmail === 'UNKNOWN_REQUESTER' || !requesterEmail) return;
  const subject = `Unlock Request Status: ${status} for Personnel ID ${personnelId}`;
  
  let body = '';
  if (status === 'APPROVED') {
    body = `
      Good news!
      Your request to unlock Personnel ID ${personnelId} has been **APPROVED** by the Admin.
      You may now return to the Attendance Plan Monitor app and refresh your browser to edit the schedule.
      ---
      This notification confirms the lock is removed.
    `;
  } else if (status === 'REJECTED') {
    body = `
      Your request to unlock Personnel ID ${personnelId} has been **REJECTED** by the Admin.
      The print lock remains active, and the schedule cannot be edited at this time. Please contact your Admin for details.
      ---
      This is an automated notification.
    `;
  } else {
      return; // Do nothing for other statuses
  }
  
  try {
    MailApp.sendEmail({
      to: requesterEmail,
      subject: subject,
      body: body,
      name: 'Attendance Plan Monitor (Status Update)'
    });
    Logger.log(`[sendRequesterNotification] Status ${status} email sent to requester: ${requesterEmail}`);
  } catch (e) {
    Logger.log(`[sendRequesterNotification] Failed to send status email to ${requesterEmail}: ${e.message}`);
  }
}


/**
 * Admin Unlock for Printed Personnel IDs.
 * Checks for Admin authority and removes the IDs from the PrintLog entry.
 * @param {string} sfcRef 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @param {string[]} personnelIdsToUnlock - Array of clean Personnel IDs to unlock.
 */
function unlockPersonnelIds(sfcRef, year, month, shift, personnelIdsToUnlock) {
    const userEmail = Session.getActiveUser().getEmail();
    if (!ADMIN_EMAILS.includes(userEmail)) {
      // Throw error if user is not in the Admin list
      throw new Error("AUTHORIZATION ERROR: Only admin users can unlock printed schedules. Contact administrator.");
    }

    if (!personnelIdsToUnlock || personnelIdsToUnlock.length === 0) return;

    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateLogSheet(ss);
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;
    const PLAN_SHEET_NAME_COL_INDEX = 2; // Col C (0-based)
    const LOCKED_IDS_COL_INDEX = LOG_HEADERS.length - 1;
    // Col J (0-based)
    
    const range = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length);
    // Use getValues() to get raw array values
    const values = range.getValues(); 
    const rangeToUpdate = [];
    values.forEach((row, rowIndex) => {
      const rowNumInSheet = rowIndex + 2; // 1-based row number in the sheet
      const planSheetNameInLog = String(row[PLAN_SHEET_NAME_COL_INDEX] || '').trim();
      const lockedIdsString = String(row[LOCKED_IDS_COL_INDEX] || '').trim();
      
      // Only check log entries for the CURRENT plan sheet
      if (planSheetNameInLog === planSheetName && lockedIdsString) {
        // Map to clean IDs and filter out empty/invalid ones
        const currentLockedIds = lockedIdsString.split(',')
                                                .map(id => cleanPersonnelId(id))
                                       
        .filter(id => id.length >= 3);
        
        let updatedLockedIds = [...currentLockedIds];
        let changed = false;

        // Filter out the requested IDs to unlock
        personnelIdsToUnlock.forEach(unlockId => {
          const index = updatedLockedIds.indexOf(unlockId);
          if (index > -1) {
            
            updatedLockedIds.splice(index, 1);
            changed = true;
          }
        });
        if (changed) {
          const newLockedIdsString = updatedLockedIds.join(',');
          // Mark the cell in column J for update
          rangeToUpdate.push({
              row: rowNumInSheet,
              col: LOCKED_IDS_COL_INDEX + 1, // 1-based
              value: newLockedIdsString
          });
        }
      }
    });
    // --- AUDIT FIX: Preserve the PrintLog entry for audit.
    // Only update Col J. ---
    rangeToUpdate.forEach(update => {
        const targetRange = logSheet.getRange(update.row, update.col);
        const newValue = update.value;
        
        // CRITICAL: We always write the newValue (a list of remaining IDs or an empty string).
        // setNumberFormat('@') is crucial to preserve IDs as strings.
        targetRange.setNumberFormat('@').setValue(newValue);
    });
    // NOTE: The previous logic to delete the log row when Col J becomes empty has been
    // permanently removed to preserve the audit trail (Reference #).
    Logger.log(`[unlockPersonnelIds] Successfully unlocked ${personnelIdsToUnlock.length} IDs for ${planSheetName}. (Log entry preserved for audit.)`);
}


/**
 * UPDATED FUNCTION: Sends an HTML email notification to Admin(s) for an unlock request,
 * including actionable Approve/Reject buttons via Web App URL.
 */
function requestUnlockEmailNotification(sfcRef, year, month, shift, personnelId, lockedRefNum) {
  const requestingUserEmail = Session.getActiveUser().getEmail(); 
  const adminEmails = ADMIN_EMAILS.join(', ');
  const date = new Date(year, month, 1);
  const planPeriod = date.toLocaleString('en-US', { month: 'long', year: 'numeric' });
  const shiftDisplay = (shift === '1stHalf' ? '1st to 15th' : '16th to End');
  // **NEW: Encode Requester Email in the URL**
  const requesterEmailEncoded = encodeURIComponent(requestingUserEmail);
  // Build the Web App URL
  const webAppUrl = ScriptApp.getService().getUrl();
  // CRITICAL: We include req_email in the URL
  const unlockUrl = `${webAppUrl}?action=unlock&sfc=${sfcRef}&yr=${year}&mon=${month + 1}&shift=${shift}&id=${personnelId}&ref=${lockedRefNum}&req_email=${requesterEmailEncoded}`;
  const rejectUrl = `${webAppUrl}?action=reject_info&sfc=${sfcRef}&id=${personnelId}&req_email=${requesterEmailEncoded}`;
  const htmlBody = `
    <p style="font-size: 14px;">Mayroong Attendance Plan Unlock Request na isinumite.</p>

    <hr style="margin: 10px 0;">
    
    <p style="font-size: 14px;"><b>Requested By:</b> ${requestingUserEmail}</p>
    <p style="font-size: 14px;"><b>Personnel ID:</b> ${personnelId}</p>
    <p style="font-size: 14px;"><b>SFC Ref #:</b> ${sfcRef}</p>
    <p style="font-size: 14px;"><b>Plan Period:</b> ${planPeriod} (${shiftDisplay} Half)</p>
    <p style="font-size: 14px;"><b>Locked Ref #:</b> ${lockedRefNum}</p>
    
    <hr style="margin: 10px 0;">

    <h3 style="color: #1e40af;">Admin Action Required:</h3>
    
    <div style="margin-top: 15px;">
        <a href="${unlockUrl}" target="_blank" 
           style="background-color: #10b981; color: white; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; font-weight: bold; margin-right: 10px;">
           ✅ APPROVE & UNLOCK
        </a>
        
        <a href="${rejectUrl}" target="_blank" 
           style="background-color: #f59e0b; color: white; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; font-weight: bold;">
           ❌ REJECT (Log Only)
        </a>
    </div>

    <p style="margin-top: 20px; font-size: 12px; color: #6b7280;">Ang pag-Approve ay magre-remove ng print lock. Kailangan naka-login ka bilang Admin user upang gumana ang link.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: adminEmails,
      subject: `ATTN: Admin Unlock Request - ID ${personnelId} for ${sfcRef}`,
      htmlBody: htmlBody, // Gamitin ang htmlBody para sa buttons
      name: 'Attendance Plan Monitor (Automated Request)'
    });
    Logger.log(`[requestUnlockEmailNotification] Sent request email for ID ${personnelId} to ${adminEmails}`);
    return { success: true, message: `Unlock request sent to Admin(s): ${adminEmails}` };
  } catch (e) {
    Logger.log(`[requestUnlockEmailNotification] Failed to send email: ${e.message}`);
    return { success: false, message: `WARNING: Failed to send request email. Error: ${e.message}` };
  }
}

/**
 * Handles the unlock action triggered by the URL/email link.
 */
function processAdminUnlockFromUrl(params) {
  const sfcRef = params.sfc;
  const personnelId = params.id;
  const lockedRefNum = params.ref;
  const requesterEmail = params.req_email ? decodeURIComponent(params.req_email) : 'UNKNOWN_REQUESTER';
  // NEW: Requester Email
  
  // 1. Check Admin authorization (Crucial for security)
  const userEmail = Session.getActiveUser().getEmail();
  if (!ADMIN_EMAILS.includes(userEmail)) {
    return HtmlService.createHtmlOutput('<h1 style="color: red;">AUTHORIZATION FAILED</h1><p>You are not authorized to perform this action. Your email: ' + userEmail + '</p>');
  }

  // Handle Reject (Informational) action
  if (params.action === 'reject_info') {
      // 1. Send notification to the original requester (REJECTED)
      sendRequesterNotification('REJECTED', personnelId, requesterEmail);
      const template = HtmlService.createTemplateFromFile('UnlockStatus');
      template.status = 'INFO';
      template.message = `Admin (${userEmail}) acknowledged the REJECT click for ID ${personnelId}.
      Notification sent to ${requesterEmail}. No data was changed. The lock remains active.`;
      return template.evaluate().setTitle('Reject Status');
  }
  
  // Handle Approve (Unlock) action
  if (params.action === 'unlock' && params.yr && params.mon && params.shift) {
      try {
        const year = parseInt(params.yr, 10);
        // Month in URL is 1-based, GAS is 0-based.
        const month = parseInt(params.mon, 10) - 1; 
        const shift = params.shift;
        // Perform the unlock using the existing logic
        unlockPersonnelIds(sfcRef, year, month, shift, [personnelId]);
        // 1. Send notification to the original requester (APPROVED)
        sendRequesterNotification('APPROVED', personnelId, requesterEmail);
        // Return Success HTML
        const template = HtmlService.createTemplateFromFile('UnlockStatus');
        template.status = 'SUCCESS';
        template.message = `Successfully unlocked Personnel ID ${personnelId}. The Print Lock (Ref #${lockedRefNum}) has been removed by Admin (${userEmail}).
        Notification sent to ${requesterEmail}.`;
        return template.evaluate().setTitle('Unlock Status');

      } catch (e) {
        // Return Failure HTML
        const template = HtmlService.createTemplateFromFile('UnlockStatus');
        template.status = 'ERROR';
        template.message = `Failed to unlock ID ${personnelId}. Error: ${e.message}`;
        return template.evaluate().setTitle('Unlock Status');
      }
  }
  
  return HtmlService.createHtmlOutput('<h1>Invalid Action</h1><p>The URL provided is incomplete or incorrect.</p>');
}
