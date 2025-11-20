const SPREADSHEET_ID = '1rQnJGqcWcEBjoyAccjYYMOQj7EkIu1ykXTMLGFzzn2I';
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; 
const CONTRACTS_SHEET_NAME = 'MASTER';

const MASTER_HEADER_ROW = 5;
const PLAN_HEADER_ROW = 6;
const PLAN_FIXED_COLUMNS = 4; // Personnel ID, Personnel Name, Version, Shift
// NEW: Signatory Master Sheet for Checked/Approved By repository
const SIGNATORY_MASTER_SHEET = 'SignatoryMaster';
// ADMIN USER CONFIGURATION FOR UNLOCK FEATURE
const ADMIN_EMAILS = ['mcdmarketingstorage@megaworld-lifestyle.com'];
// --- UPDATED CONFIGURATION FOR PRINT LOG (11 Columns) ---
const LOG_SHEET_NAME = 'PrintLog';
const LOG_HEADERS = [
    'Reference #', 
    'SFC Ref#', 
    'Plan Sheet Name',      
    'Plan Period Display', 
    'Payor Company', 
    'Agency',
    'Sub Property',         // <--- NEW COLUMN
    'Service Type',
    'User Email', 
    'Timestamp',
    'Locked Personnel IDs'  
];
// --- UPDATED CONFIGURATION FOR SCHEDULE AUDIT LOG ---
const AUDIT_LOG_SHEET_NAME = 'ScheduleAuditLog';
const AUDIT_LOG_HEADERS = [
    'Timestamp', 
    'User Email', 
    'SFC Ref#', 
    'Personnel ID', 
    'Personnel Name', // <--- NEW COLUMN: PERSONNEL NAME
    'Plan Sheet Name', 
    'Date (YYYY-M-D)', 
    'Shift', 
    'Reference #', // <--- ADDED: Reference # for edited locked schedule
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

    const planHeadersCount = PLAN_FIXED_COLUMNS + 31; // 35 Columns
    const rowsToAppend = [];
    existingIds.forEach(id => {
        // Find the employee's name for the new column
        const employee = empData.find(e => cleanPersonnelId(e['Personnel ID']) === id);
        const name = employee ? employee['Personnel Name'] : 'N/A';
        
        if (shiftToAppend === '1stHalf') {
            const planRow1 = Array(planHeadersCount).fill('');
            planRow1[0] = id; // Personnel ID
            planRow1[1] = name; // Personnel Name
            planRow1[2] = (1).toFixed(1); // Version 1.0
            planRow1[3] = '1stHalf'; // Shift
            rowsToAppend.push(planRow1);
        }
        if (shiftToAppend === '2ndHalf') {
         
            const planRow2 = Array(planHeadersCount).fill('');
            planRow2[0] = id; // Personnel ID
            planRow2[1] = name; // Personnel Name
            planRow2[2] = (1).toFixed(1); // Version 1.0
            planRow2[3] = '2ndHalf'; // Shift
            rowsToAppend.push(planRow2);
        }
    });
    if (rowsToAppend.length > 0) {
        // Set the column number format for the first 3 columns (ID, Name, Version) to Plain Text
        const startRow = planSheet.getLastRow() + 1;
        const range = planSheet.getRange(startRow, 1, rowsToAppend.length, planHeadersCount);
        range.setValues(rowsToAppend);
        planSheet.getRange(startRow, 1, rowsToAppend.length, PLAN_FIXED_COLUMNS - 1).setNumberFormat('@'); // ID, Name, Version
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
        // UPDATED: Added Personnel Name and Version
        const base = ['Personnel ID', 'Personnel Name', 'Version', 'Shift'];
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
        
        // Set number format for Personnel ID, Name, Version columns ('@' for plain text)
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, PLAN_FIXED_COLUMNS - 1).setNumberFormat('@');
        
        Logger.log(`[createContractSheets] Created Horizontal Attendance Plan sheet for ${planSheetName} with headers at Row ${PLAN_HEADER_ROW}.`);
        // Note: I-REMOVE ANG AUTOMATIC PRE-POPULATION DITO PARA HINDI LUMABAS ANG BLANK ROWS
        // REMOVED: appendExistingEmployeeRowsToPlan(sfcRef, planSheet, shift);
        // UPDATED: Column Hiding Indices (Day 1 is now Col 5, Day 16 is Col 20)
        if (shift === '1stHalf') {
            const START_COL_TO_HIDE = 20; // Day 16
            const NUM_COLS_TO_HIDE = 16; 
            planSheet.hideColumns(START_COL_TO_HIDE, NUM_COLS_TO_HIDE);
            Logger.log(`[createContractSheets] Hiding Day 16-31 columns for 1stHalf sheet.`);
        } else if (shift === '2ndHalf') {
            const START_COL_TO_HIDE = 5; // Day 1
            const NUM_COLS_TO_HIDE = 15; // Day 1 to Day 15
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
  
      contractIdKey ? (c[contractIdKey] || '').toString() : '',     
      status: statusKey ? (c[statusKey] || '').toString() : '',   
      payorCompany: payorKey ? (c[payorKey] || '').toString() : '', 
      agency: agencyKey ? (c[agencyKey] || '').toString() : '',       
      serviceType: serviceTypeKey ? (c[serviceTypeKey] || '').toString() : '',   
      headCount: parseInt(headCountKey ? c[headCountKey] 
     
       : 0) || 
      0, 
      sfcRef: sfcRefKey ? 
      (c[sfcRefKey] ||
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
 * Fetches the Signatory Master Data (list of names for Checked By/Approved By options).
 * @returns {string[]} An array of Signatory Names.
 */
function getSignatoryMasterData() {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = getOrCreateSignatoryMasterSheet(ss);

        if (sheet.getLastRow() < 2) return [];
        // Basahin ang lahat ng pangalan mula sa Row 2, Column 1
        const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
        const values = range.getDisplayValues();

        // Flatten the 2D array and filter out empty strings
        return values.map(row => String(row[0] || '').trim()).filter(name => name);
    } catch (e) {
        Logger.log(`[getSignatoryMasterData] ERROR: ${e.message}`);
        return [];
    }
}

// --- NEW HELPER: Fetches the clean employee master data for auto-filling/datalist.
// ---
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
    })).filter(e => e.id);
    // Only return rows with a valid ID
}
// --- END NEW FUNCTION ---

/**
 * NEW HELPER: Fetches the Personnel Name from the master list using the Personnel ID.
 * @param {string} sfcRef The SFC Ref #.
 * @param {string} personnelId The Personnel ID to look up.
 * @returns {string} The employee name or 'N/A' if not found.
 */
function getEmployeeNameFromMaster(sfcRef, personnelId) {
    if (!sfcRef || !personnelId) return 'N/A';
    const masterData = getEmployeeMasterData(sfcRef);
    // Reuses existing function
    const cleanId = cleanPersonnelId(personnelId);
    const employee = masterData.find(e => e.id === cleanId);
    return employee ? employee.name : 'N/A';
}


// --- NEW FUNCTION: Schedule Pattern Lookup (AutoFill Schedule - FIXED FOR RECENCY) ---

/**
 * Parses the sheet name to extract the date value for comparison.
 * @param {string} sheetName The sheet name (e.g., "SFC - Nov 2025 - 2ndHalf AttendancePlan").
 * @returns {Date} The starting Date object for the sheet's period.
 */
function parseSheetDate(sheetName) {
    try {
        const parts = sheetName.split(' - ');
        if (parts.length < 3) return null; 
        
        const datePart = parts[1];
        // e.g., "Nov 2025"
        const shiftPart = parts[2].split(' ')[0];
        // e.g., "2ndHalf"
        
        const monthYear = new Date(datePart);
        const day = (shiftPart === '2ndHalf') ? 16 : 1;
        // Return a date object that correctly represents the start of the shift
        return new Date(monthYear.getFullYear(), monthYear.getMonth(), day);
    } catch (e) {
        Logger.log(`[parseSheetDate] Error parsing date from sheet name ${sheetName}: ${e.message}`);
        return null;
    }
}


/**
 * Finds ALL Attendance Plan sheets for a given SFC Ref and sorts them by date (most recent first).
 * @param {string} sfcRef The SFC Ref #.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The target spreadsheet.
 * @returns {Array<object>} An array of {name: string, date: Date}, sorted descending.
 */
function getSortedPlanSheets(sfcRef, ss) {
    const sfcPrefix = (sfcRef || '').replace(/[\\/?*[]/g, '_');
    const allSheetNames = ss.getSheets().map(s => s.getName()).filter(name => 
        name.startsWith(sfcPrefix) && name.endsWith('AttendancePlan')
    );
    if (allSheetNames.length === 0) return [];

    const sheetInfo = allSheetNames.map(sheetName => {
        const sheetDate = parseSheetDate(sheetName);
        return { name: sheetName, date: sheetDate };
    }).filter(info => info.date);
    // Filter out sheets that couldn't be parsed

    // Sort by date (most recent first)
    sheetInfo.sort((a, b) => b.date.getTime() - a.date.getTime());
    Logger.log(`[getSortedPlanSheets] Found and sorted ${sheetInfo.length} plan sheets.`);
    return sheetInfo;
}


/**
 * Scans the MOST RECENT available Attendance Plan sheet for a given Personnel ID
 * to determine the schedule status for each day of the week.
 * @param {string} sfcRef The SFC Ref #.
 * @param {string} personnelId The Personnel ID to look up.
 * @returns {object} A map: { DayOfWeek (0-6): StatusString (e.g., '08:00-17:00' or 'RD') }
 */
function getEmployeeSchedulePattern(sfcRef, personnelId) {
    if (!sfcRef || !personnelId) return {};
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const cleanId = cleanPersonnelId(personnelId);
    
    // 1. Find ALL plan sheets, sorted by recency
    const sortedPlanSheets = getSortedPlanSheets(sfcRef, ss);
    // CRITICAL FIX
    
    let targetSheetInfo = null;
    let targetRows = [];
    let sheetName = null;
    const versionIndex = 2; // Version is the 3rd column (index 2)

    // --- NEW LOGIC: Iterate through sorted sheets until the Personnel ID is found ---
    for (const sheetInfo of sortedPlanSheets) {
        const sheet = ss.getSheetByName(sheetInfo.name);
        if (!sheet || sheet.getLastRow() < PLAN_HEADER_ROW) continue;

        try {
            // Read data from the plan sheet (using getDisplayValues for consistency)
            const lastRow = sheet.getLastRow();
            const numRowsToRead = lastRow - PLAN_HEADER_ROW;
            const numColumns = sheet.getLastColumn();
            
            if (numRowsToRead <= 0 || numColumns < (PLAN_FIXED_COLUMNS + 3)) continue; // Check minimum columns (4 fixed + 3 days)
            // Read all data including headers
            const planValues = sheet.getRange(PLAN_HEADER_ROW, 1, numRowsToRead + 1, numColumns).getDisplayValues();
            const headers = planValues[0];
            const dataRows = planValues.slice(1);
            const personnelIdIndex = headers.indexOf('Personnel ID');
            
            if (personnelIdIndex === -1) continue;
            // Find ALL rows for the target Personnel ID in this sheet
            const rowsForId = dataRows.filter(row => cleanPersonnelId(row[personnelIdIndex]) === cleanId);
            if (rowsForId.length > 0) {
                // NEW: Find the single latest version row among all rows for this ID in this sheet
                const latestRow = rowsForId.reduce((latest, current) => {
                     const currentVersion = parseFloat(current[versionIndex]) || 0; // Use parseFloat
                     if (currentVersion > (parseFloat(latest[versionIndex]) || 0)) {
                         return current;
                     }
                     return latest;
                }, rowsForId[0]);
                
                targetSheetInfo = sheetInfo;
                targetRows = [latestRow]; // Only use the latest version row for the pattern
                sheetName = sheetInfo.name;
                Logger.log(`[getEmployeeSchedulePattern] Found latest version (v${latestRow[versionIndex]}) of ID ${cleanId} in sheet: ${sheetInfo.name}. Stopping search.`);
                break;
            // Found the pattern, stop searching older sheets
            }
        } catch (e) {
            Logger.log(`[getEmployeeSchedulePattern] ERROR reading sheet ${sheetInfo.name}: ${e.message}`);
            continue;
        }
    }
    // --- END NEW LOGIC ---

    if (!targetSheetInfo || targetRows.length === 0) {
        Logger.log(`[getEmployeeSchedulePattern] ID ${cleanId} not found in any plan sheet.`);
        return {};
    }

    // 2. Extract Year and Month from the found sheet name
    const parts = sheetName.split(' - ');
    if (parts.length < 3) return {}; 
    const datePart = parts[1]; 
    
    const monthYear = new Date(datePart);
    const sheetYear = monthYear.getFullYear();
    const sheetMonth = monthYear.getMonth(); // 0-based

    const dayPatternCounter = {};
    // Key: DayOfWeek (0-6), Value: { Status: Count }
    const dayPatternMap = {};
    // Final result: { DayOfWeek: Status }

    try {
        // Prepare header map from the selected sheet's headers (re-read or pass the sheet object if possible)
        const sheet = ss.getSheetByName(sheetName);
        // CRITICAL FIX: Ensure to read the full range based on the latest row
        const planValues = sheet.getRange(PLAN_HEADER_ROW, 1, sheet.getLastRow() - PLAN_HEADER_ROW + 1, sheet.getLastColumn()).getDisplayValues();
        const headers = planValues[0];

        const sanitizedHeadersMap = {};
        headers.forEach((header, index) => {
            sanitizedHeadersMap[sanitizeHeader(header)] = index;
        });
        // 4. Aggregate pattern from the SINGLE latest row found
        targetRows.forEach(targetRow => {
             for (let d = 1; d <= 31; d++) {
                const currentDate = new Date(sheetYear, sheetMonth, d);
                // Check if the current day actually falls in the sheet's month/year
      
                if (currentDate.getMonth() !== sheetMonth) continue; 
             
                const dayOfWeek = currentDate.getDay(); 
                
                let lookupHeader = '';
           
                const 
                monthShortRaw = currentDate.toLocaleString('en-US', { month: 'short' });
                const monthShort = 
                (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '').replace(/\s/g, '');
                lookupHeader = `${monthShort}${d}`;
    
                const dayIndex = sanitizedHeadersMap[sanitizeHeader(lookupHeader)];
  
               
                if (dayIndex !== undefined) {
       
                    const status = String(targetRow[dayIndex] ||
                    '').trim();
                    if (status && status !== 'NA' && status !== 'RD' && status !== 'RH' && status !== 'SH') { 
                        // Only count actual schedules for pattern (time schedules)
           
                        const dayKey = dayOfWeek.toString();
                        if (!dayPatternCounter[dayKey]) {
                            dayPatternCounter[dayKey] = {};
                        }
                        dayPatternCounter[dayKey][status] = (dayPatternCounter[dayKey][status] || 0) + 1;
                    } else if (status === 'RD' || status === 'RH' || status === 'SH') {
                        // Special treatment for fixed statuses: if found, they are strong patterns
                        const dayKey = dayOfWeek.toString();
                        if (!dayPatternCounter[dayKey]) {
                            dayPatternCounter[dayKey] = {};
                        }
                        dayPatternCounter[dayKey][status] = (dayPatternCounter[dayKey][status] || 0) + 1;
                    }
                }
            }
        });
        // 5. Determine the final pattern (Most Frequent)
        Object.keys(dayPatternCounter).forEach(dayOfWeek => {
            const statusCounts = dayPatternCounter[dayOfWeek];
            let maxCount = 0;
            let mostFrequentStatus = '';
            
            Object.keys(statusCounts).forEach(status => {
            
                if (statusCounts[status] > maxCount) {
                    maxCount = statusCounts[status];
                    mostFrequentStatus = status;
                }
            });
        
            if (mostFrequentStatus) {
                dayPatternMap[dayOfWeek] = mostFrequentStatus;
            }
        });
    } catch (e) {
        Logger.log(`[getEmployeeSchedulePattern] ERROR processing sheet ${sheetName}: ${e.message}`);
    }

    Logger.log(`[getEmployeeSchedulePattern] Final Pattern for ID ${cleanId}: ${JSON.stringify(dayPatternMap)}`);
    return dayPatternMap;
}
// --- END NEW FUNCTION ---

/**
 * UPDATED FUNCTION: Kumuha ng map ng CURRENTLY locked IDs at ang Reference # na nag-lock sa kanila.
 * This is used for UI/Client-side locking/filtering (only CURRENTLY locked IDs).
 * It ignores IDs prefixed with 'UNLOCKED:'.
 * Returns: { 'Personnel ID': 'Reference #' }
 */
function getLockedPersonnelIds(ss, planSheetName) {
    const logSheet = getOrCreateLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return {};
    // Define column indices (based on the 11-column LOG_HEADERS)
    const REF_NUM_COL = 1;
    const PLAN_SHEET_NAME_COL = 3;
    const LOCKED_IDS_COL = LOG_HEADERS.length; 

    // Basahin ang lahat ng data mula Col A (1) hanggang Col K (11)
    const values = logSheet.getRange(2, 1, lastRow - 1, LOCKED_IDS_COL).getDisplayValues();
    const lockedIdRefMap = {}; 

    values.forEach(row => {
        const refNum = String(row[REF_NUM_COL - 1] || '').trim();
        const planSheetNameInLog = String(row[PLAN_SHEET_NAME_COL - 1] || '').trim();
        const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim(); 
        
        // This logic ensures only IDs *currently* in the LOCKED_IDS_COL are returned
        if (planSheetNameInLog === planSheetName && lockedIdsString) {
       
            const idsList = lockedIdsString.split(',').map(id => id.trim());
            
            idsList.forEach(idWithPrefix => {
                const cleanId = cleanPersonnelId(idWithPrefix);
                 
                // CRITICAL CHECK: Only include 
                // IDs that are NOT prefixed with 'UNLOCKED:'
                if (cleanId.length >= 3 && !idWithPrefix.startsWith('UNLOCKED:')) { 
                     // We use the first Ref # found for this ID as the source of truth
                  
                    if (!lockedIdRefMap[cleanId]) { 
       
                         lockedIdRefMap[cleanId] = refNum;
                    }
                }
            });
        }
    });
    // NOTE: This intentionally returns only IDs that are *currently* locked (not prefixed with UNLOCKED:).
    return lockedIdRefMap; 
}


/**
 * NEW HELPER: Kumuha ng map ng ALL Personnel IDs na NA-LOG sa PrintLog, at ang Reference # na nauugnay sa kanila.
 * Ito ay para sa audit trail (logging the history), HINDI para sa UI lock.
 * It reads ALL IDs (locked or unlocked/prefixed) from the PrintLog column K.
 * Returns: { 'Personnel ID': 'Reference #' }
 */
function getHistoricalReferenceMap(ss, planSheetName) {
    const logSheet = ss.getSheetByName('PrintLog');
    if (!logSheet || logSheet.getLastRow() < 2) return {};

    const LOG_HEADERS_COUNT = 11; 
    const REF_NUM_COL = 1;              
    const PLAN_SHEET_NAME_COL = 3;
    const LOCKED_IDS_COL = LOG_HEADERS.length; 

    const allValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, LOG_HEADERS_COUNT).getDisplayValues();
    const historicalRefMap = {};
    // Iterate backwards to prioritize the most recent log entry for an ID.
    for (let i = allValues.length - 1; i >= 0; i--) {
        const row = allValues[i];
        // CRITICAL FIX: I-re-pad ang Ref # dito
        const refNumRaw = String(row[REF_NUM_COL - 1] || '').trim();
        const refNum = refNumRaw.padStart(6, '0'); // Ensures 000001 format
        
        const planSheetNameInLog = String(row[PLAN_SHEET_NAME_COL - 1] || '').trim();
        const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim(); 
        
        if (planSheetNameInLog === planSheetName && refNum) {
            
            // CRITICAL FIX: Read ALL IDs in the string (including those prefixed with UNLOCKED:)
            const allIdsInString = lockedIdsString.split(',').map(s => s.trim());
            allIdsInString.forEach(idWithPrefix => {
                const cleanId = cleanPersonnelId(idWithPrefix);
                
                if (cleanId.length >= 3) { 
                     // Only log the Ref # for the most recent entry found
          
                     if (!historicalRefMap[cleanId]) { 
                        historicalRefMap[cleanId] = refNum;
                     }
                }
            });
        }
    }
    return historicalRefMap;
}

function getAttendancePlan(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    // NOTE: empData is the FULL master employee list (used by client for datalist/lookup)
    const empData = getSheetData(TARGET_SPREADSHEET_ID, empSheetName);
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const planSheet = ss.getSheetByName(planSheetName);
    // NEW: Kunin ang map ng CURRENTLY locked IDs at Reference #
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
    if (numRowsToRead <= 0 || numColumns < (PLAN_FIXED_COLUMNS + 3)) { // UPDATED MIN COLUMN CHECK
        return { employees: [], planMap: {}, lockedIds: lockedIds, lockedIdRefMap: lockedIdRefMap };
    // UPDATED: Return empty employees array
    }

    // CRITICAL FIX: Use getDisplayValues() to ensure time formats (08:00-17:00) are read as strings.
    const planValues = planSheet.getRange(HEADER_ROW, 1, numRowsToRead + 1, numColumns).getDisplayValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name'); // <-- NEW INDEX
    const versionIndex = headers.indexOf('Version');     // <-- NEW INDEX
    const shiftIndex = headers.indexOf('Shift');
    
    const planMap = {};
    const sanitizedHeadersMap = {};
    
    // --- NEW CRITICAL LOGIC: Filter for Latest Version Row ---
    const latestVersionMap = {}; // Key: Personnel ID_Shift, Value: Latest Row (Array)

    dataRows.forEach(row => {
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        const currentShift = String(row[shiftIndex] || '').trim();
        const version = parseFloat(row[versionIndex]) || 0; // Read Version as float
        
        if (currentShift === shift && id) {
            const mapKey = `${id}_${currentShift}`;
            const existingRow = latestVersionMap[mapKey];
            
            // Only keep the row with the highest version number
            if (!existingRow || version > (parseFloat(existingRow[versionIndex]) || 0)) {
                latestVersionMap[mapKey] = row;
            }
        }
    });

    const latestDataRows = Object.values(latestVersionMap);
    // --- END NEW CRITICAL LOGIC ---
    
    const employeesInPlan = new Set(); // NEW: Set of IDs that actually have a plan entry for this shift
    headers.forEach((header, index) => {
        const cleanedHeader = sanitizeHeader(header); 
        sanitizedHeadersMap[cleanedHeader] = index;
    });
    latestDataRows.forEach((row, rowIndex) => { // <-- USE LATEST FILTERED ROWS
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        const currentShift = String(row[shiftIndex] || '').trim();
        
        // This check is now redundant but kept for safety
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
            no: 
            0, // Temporarily set to 0
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
        // saveEmployeeInfoBulk now handles appending to both Employee Master and Plan sheet.
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
    
    // UPDATED: Removed Total Head Count, Added SFC Ref#
    const data = [
        ['PAYOR COMPANY', info.payor],          
        ['AGENCY', info.agency],                
        ['SERVICE TYPE', info.serviceType],     
        ['SFC REF#', sfcRef],                   // <-- NEW METADATA FIELD
        ['PLAN PERIOD', dateRange]    
// REMOVED: ['TOTAL HEAD COUNT', info.headCount],    
 
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
    // Tiyakin na makuha ang CURRENT lock status map (para sa filtering) - not needed here, already filtered in saveAllData
    // NEW STEP: Kunin ang HISTORICAL Reference map (para sa logging ng Ref# ng inedit na entry)
    const historicalRefMap = getHistoricalReferenceMap(ss, planSheetName);
    // --- End NEW STEP ---
    
    // 1. Read all data
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
        // If sheet is new/empty, stop here (should have been handled by createContractSheets)
        throw new Error("Plan sheet is empty or improperly formatted.");
    }
    
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name'); // <-- NEW INDEX
    const versionIndex = headers.indexOf('Version');     // <-- NEW INDEX
    const shiftIndex = headers.indexOf('Shift');
    
    if (personnelIdIndex === -1 || shiftIndex === -1 || versionIndex === -1) {
        throw new Error("Missing critical column in AttendancePlan sheet (Personnel ID, Version, or Shift).");
    }
    
    // 2. Find the Latest Version Row for each ID_Shift
    const latestVersionMap = {}; // Key: Personnel ID_Shift, Value: Latest Row (Array)
    const dataRows = values.slice(1);
    
    dataRows.forEach(row => {
        const id = cleanPersonnelId(row[personnelIdIndex]);
        const currentShift = String(row[shiftIndex] || '').trim();
        // Use parseFloat to accurately compare version numbers like 1.0, 2.0
        const version = parseFloat(row[versionIndex]) || 0; 
        
        if (currentShift === shift && id) {
            const mapKey = `${id}_${currentShift}`;
            const existingRow = latestVersionMap[mapKey];
            // Only keep the row with the highest version number
            if (!existingRow || version > (parseFloat(existingRow[versionIndex]) || 0)) {
                latestVersionMap[mapKey] = row;
            }
        }
    });
    
    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        sanitizedHeadersMap[sanitizeHeader(header)] = index;
    });
    
    // 3. Group changes by Personnel ID_Shift
    const changesByRow = changes.reduce((acc, change) => {
        const key = `${change.personnelId}_${change.shift}`;
        if (!acc[key]) acc[key] = [];
        acc[key].push(change);
        return acc;
    }, {});
    
    const rowsToAppend = [];
    const auditLogSheet = getOrCreateAuditLogSheet(ss);
    const userEmail = Session.getActiveUser().getEmail();
    
    // ** NEW: Fetch master data for name lookup **
    const masterEmployeeMap = getEmployeeMasterData(sfcRef).reduce((map, emp) => { 
        map[emp.id] = emp.name;
        return map;
    }, {});
    
    // ** NEW: Sort changes by Date (dayKey) before processing/logging **
    changes.sort((a, b) => {
        // dayKey format is YYYY-M-D (e.g., 2025-11-1)
        const dateA = new Date(a.dayKey);
        const dateB = new Date(b.dayKey);
        // Secondary sort by Personnel ID for consistent grouping (optional but good practice)
        if (dateA.getTime() !== dateB.getTime()) {
            return 
            dateA.getTime() - dateB.getTime();
        }
        return a.personnelId.localeCompare(b.personnelId);
    });
    // ** END NEW SORTING **


    Object.keys(changesByRow).forEach(rowKey => {
        const dailyChanges = changesByRow[rowKey];
        const [personnelId, dataShift] = rowKey.split('_');
        
        if (dataShift !== shift) return; // Skip if not the current shift

        const latestVersionRow = latestVersionMap[rowKey];
        let newRow;
        let currentVersion = 0;

        // --- CRITICAL FIX 1: V1 Creation Logic (First Schedule Save) ---
        if (!latestVersionRow) {
            // Initial Save (Version 1.0)
            const planHeadersCount = headers.length; 
            const name = masterEmployeeMap[personnelId] || 'N/A';
            
            newRow = Array(planHeadersCount).fill('');
            newRow[personnelIdIndex] = personnelId; // Personnel ID
            newRow[nameIndex] = name;             // Personnel Name
            newRow[shiftIndex] = dataShift;       // Shift
            currentVersion = 0; // Starting point for V1.0
            
        } else {
            // Standard Update Scenario (Version 2.0, 3.0, etc.)
            newRow = [...latestVersionRow]; // Copy the latest version
            currentVersion = parseFloat(latestVersionRow[versionIndex]) || 0; // Read current version as float
        }
        // --- END CRITICAL FIX 1 ---
        
        let isRowChanged = false;
        
        // Apply all daily changes to the new row and log audit trail
        dailyChanges.forEach(data => {
            const { dayKey, status: newStatus } = data;
            const dayNumber = parseInt(dayKey.split('-')[2], 10);
            
            // Re-create the header lookup (e.g., Nov1)
            let targetLookupHeader = '';
            const date = new Date(year, month, dayNumber);
            if (date.getMonth() === month) {
                const monthShortRaw = date.toLocaleString('en-US', { month: 'short' });
                const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '').replace(/\s/g, '');
                targetLookupHeader = `${monthShort}${dayNumber}`; 
            } else {
                targetLookupHeader = `Day${dayNumber}`;
            }
            
            const dayColIndex = sanitizedHeadersMap[sanitizeHeader(targetLookupHeader)];
            if (dayColIndex !== undefined) {
                // Get OLD Status from the row we are basing the update on (either latestVersionRow or the newly created V1 row)
                const oldStatus = String((latestVersionRow || newRow)[dayColIndex] || '').trim(); 
                
                if (oldStatus !== newStatus) {
                    isRowChanged = true;
                    
                    newRow[dayColIndex] = newStatus; // Apply change to the NEW row copy
                    
                    // Log the change to the audit sheet
                    const lockedRefNum = historicalRefMap[personnelId] || ''; 
                    const personnelName = masterEmployeeMap[personnelId] || 'N/A';

                    const logEntry = [
                        new Date(), userEmail, sfcRef, personnelId, personnelName, planSheetName, 
                        dayKey, dataShift, `'${lockedRefNum}`, oldStatus, newStatus
                    ];
                    auditLogSheet.appendRow(logEntry);
                    Logger.log(`[AuditLog] Change logged for ID ${personnelId}, Day ${dayKey}: ${oldStatus} -> ${newStatus} (Ref: ${lockedRefNum})`);
                }
            } else {
                Logger.log(`[savePlanBulk] WARNING: Day column index not found for header: ${targetLookupHeader}. Skipping update for this cell.`);
            }
        });
        
        // 5. If any change was detected, append the new row with incremented version
        if (isRowChanged) {
            // CRITICAL FIX 2: Increment and format the Version number (e.g., 0 -> 1.0, 1.0 -> 2.0)
            newRow[versionIndex] = String((currentVersion + 1).toFixed(1)); 
            rowsToAppend.push(newRow);
        }
    });

    // 6. Bulk append new version rows
    if (rowsToAppend.length > 0) {
        const newRowLength = rowsToAppend[0].length;
        const startRow = planSheet.getLastRow() + 1;
        
        planSheet.getRange(startRow, 1, rowsToAppend.length, newRowLength).setValues(rowsToAppend);
        
        // Set the format for Personnel ID, Name, Version columns (Col 1, 2, 3) to Plain Text
        planSheet.getRange(startRow, 1, rowsToAppend.length, PLAN_FIXED_COLUMNS - 1).setNumberFormat('@'); 
        
        Logger.log(`[saveAttendancePlanBulk] Appended ${rowsToAppend.length} new version rows for ${planSheetName}.`);
    }

    planSheet.setFrozenRows(HEADER_ROW); 
    Logger.log(`[saveAttendancePlanBulk] Completed Attendance Plan update for ${planSheetName}.`);
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
    // Read only the fixed columns (ID, Name, Version, Shift)
    const values = planSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, PLAN_FIXED_COLUMNS).getValues(); 
    const planKeyMap = {};
    
    // CRITICAL: We need to find the latest version for each ID_Shift
    const latestVersionMap = {}; // Key: ID_Shift, Value: Version
    
    values.forEach((row) => {
        const id = cleanPersonnelId(row[0]);
        const shift = String(row[3] || '').trim(); // Shift is now index 3
        const version = parseFloat(row[2]) || 0; // Version is index 2 (Read as float)
        
        if (id && shift) {
            const key = `${id}_${shift}`;
            // Only store the key if this is the latest version found so far
            if (version >= (latestVersionMap[key] || 0)) {
                latestVersionMap[key] = version;
                planKeyMap[key] = true; // Mark as existing (latest version)
            } else {
                 // Remove if an older version row is encountered later (shouldn't happen with correct versioning)
                 // NOTE: This logic relies on reading all and only keeping the MAX version. The simple keyMap is not enough.
                 // We will stick to the existing simple check for adding new rows, as we only need to check if ANY entry exists.
                 // If ANY entry exists, the employee is "on the plan". The versioning only applies to updates.
                 planKeyMap[key] = true; // Simpler logic for ensuring existence
            }
        }
    });

    Logger.log(`[getPlanKeyMap] Found ${Object.keys(planKeyMap).length} existing ID_Shift entries in sheet: ${planSheet.getName()}`);
    return planKeyMap;
}
// *******************************************************


// ********** NEW HELPER FUNCTION FOR AUDIT LOGGING DELETED ROWS **********
/**
 * Logs all existing schedule statuses for a given employee as 'DELETED_ROW' before the row is physically deleted.
 * @param {string} sfcRef 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} planSheet 
 * @param {string} targetShift 
 * @param {string} personnelId 
 * @param {string} userEmail 
 * @param {number} year 
 * @param {number} month 
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
    const versionIndex = headers.indexOf('Version'); // <-- NEW INDEX
    
    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        sanitizedHeadersMap[sanitizeHeader(header)] = index;
    });
    
    // 1. Find the LATEST version row to be logged/deleted
    let latestVersion = 0;
    let targetRowIndex = -1; // Index within dataRows (0-based)
    let rowToDelete = null;

    dataRows.forEach((row, index) => {
        const currentId = cleanPersonnelId(row[personnelIdIndex]);
        const currentShift = String(row[shiftIndex] || '').trim();
        const currentVersion = parseFloat(row[versionIndex]) || 0; // Use parseFloat

        if (currentId === personnelId && currentShift === targetShift && currentVersion >= latestVersion) {
            latestVersion = currentVersion;
            targetRowIndex = index;
            rowToDelete = row;
        }
    });
    
    if (targetRowIndex === -1) return; // No row found for deletion log

    const auditLogSheet = getOrCreateAuditLogSheet(planSheet.getParent());
    const logEntries = [];
    // CRITICAL FIX: Get the Historical Reference map for logging
    const ss = planSheet.getParent();
    const historicalRefMap = getHistoricalReferenceMap(ss, planSheetName); 
    const lockedRefNum = historicalRefMap[personnelId] || '';
    // --- End CRITICAL FIX ---
    
    // ** NEW: Get Personnel Name **
    const employeeName = getEmployeeNameFromMaster(sfcRef, personnelId);
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
            const oldStatus = String(rowToDelete[dayColIndex] || '').trim(); // <-- Use the latest version row
            // Log only statuses that are NOT blank or NA (to focus on actual schedules being removed)
            if (oldStatus && oldStatus !== 'NA') {
                 const logEntry = [
                    new Date(), 
                    userEmail, 
   
                    sfcRef, 
                    personnelId, 
                    employeeName, // <-- ADDED PERSONNEL NAME
                    planSheetName, 
            
                    dayKey, 
                   
                    targetShift, 
                    `'${lockedRefNum}`, // <--- FINAL FIX: Prefix with single quote (')
          
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
    
    const userEmail = Session.getActiveUser().getEmail();
    // CRITICAL: Kunin ang user email dito

    if (!empSheet) throw new Error(`Employee Sheet for SFC Ref# ${sfcRef} not found.`);
    empSheet.setFrozenRows(0);
    
    const numRows = empSheet.getLastRow() > 0 ? empSheet.getLastRow() : 1;
    const numColumns = empSheet.getLastColumn() > 0 ?
    empSheet.getLastColumn() : 4;
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
    
    // UPDATED: Walang plan rows na i-a-append dito, ipinasa sa saveAttendancePlanBulk
    // const planRowsToAppend1st = [];
    // const planRowsToAppend2nd = [];
    
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
                // For NEW employees 
 
                // yet saved), delete from master sheet (which is just appending logic removal)
                rowsToDelete.push({ 
                    rowNum: personnelIdMap[oldId], 
                    id: oldId,
 
           
                    isMasterDelete: true 
                // Flag to delete from master sheet
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


            // 2. Attendance Plan Sheet Logic: REMOVED. Initial row creation is now handled by saveAttendancePlanBulk.
            
            // // OLD LOGIC REMOVED:
            // const planHeadersCount = PLAN_FIXED_COLUMNS + 31; 
            // if (shift === '1stHalf' && !planKeyMap1st[`${newId}_1stHalf`]) {
            //     const planRow1 = Array(planHeadersCount).fill('');
            //     planRow1[0] = newId; 
            //     planRow1[1] = data.name; 
            //     planRow1[2] = (1).toFixed(1); 
            //     planRow1[3] = '1stHalf'; 
            //     planRowsToAppend1st.push(planRow1);
            // } else if (shift === '2ndHalf' && !planKeyMap2nd[`${newId}_2ndHalf`]) {
            //     const planRow2 = Array(planHeadersCount).fill('');
            //     planRow2[0] = newId; 
            //     planRow2[1] = data.name; 
            //     planRow2[2] = (1).toFixed(1); 
            //     planRow2[3] = '2ndHalf'; 
            //     planRowsToAppend2nd.push(planRow2);
            // }
            // --- END REMOVED LOGIC ---
        }
    });
    // 1. I-delete ang mga row sa Plan Sheet para sa kasalukuyang shift (DELETION ACTION)
    rowsToDelete.forEach(item => {
        // Function para i-delete ang row sa Plan Sheet
        const deletePlanRowForCurrentShift = (planSheet, targetShift) => {
            if (planSheet) {
                // Basahin ang lahat ng columns para makuha ang Version
             
                const lastRowInPlan = planSheet.getLastRow();
                if (lastRowInPlan <= PLAN_HEADER_ROW) return;
                
                const numColsInPlan = planSheet.getLastColumn();
                const planValues = planSheet.getRange(PLAN_HEADER_ROW + 1, 1, lastRowInPlan - PLAN_HEADER_ROW, numColsInPlan).getValues();
                const headers = planSheet.getRange(PLAN_HEADER_ROW, 1, 1, numColsInPlan).getValues()[0];
                const personnelIdIndexInPlan = headers.indexOf('Personnel ID');
                const shiftIndexInPlan = headers.indexOf('Shift');
                const versionIndexInPlan = headers.indexOf('Version');

                let latestVersion = 0;
                let latestRowIndex = -1; // Index within planValues (0-based)
                
                // 1. Find the latest version row to delete
                planValues.forEach((row, index) => {
                    const currentId = cleanPersonnelId(row[personnelIdIndexInPlan]);
                    const currentShift = String(row[shiftIndexInPlan] || '').trim();
                    const currentVersion = parseFloat(row[versionIndexInPlan]) || 0;

                    if (currentId === item.id && currentShift === targetShift && currentVersion >= latestVersion) {
                        latestVersion = currentVersion;
                        latestRowIndex = index;
                    }
                });
                
                if (latestRowIndex !== -1) {
                    // Row index sa sheet (1-based)
                    const sheetRowNumber = PLAN_HEADER_ROW + latestRowIndex + 1;
                    planSheet.deleteRow(sheetRowNumber);
                    Logger.log(`[saveEmployeeInfoBulk] Deleted Plan row (Version ${latestVersion}) for ID ${item.id} in ${targetShift} shift at row ${sheetRowNumber}.`);
                    return;
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
    
    // 3. Append new rows (Attendance Plan Sheet - 1st Half) - REMOVED, now in saveAttendancePlanBulk
    // if (planRowsToAppend1st.length > 0 && planSheet1st) { ... }
    
    // 4. Append new rows (Attendance Plan Sheet - 2nd Half) - REMOVED, now in saveAttendancePlanBulk
    // if (planRowsToAppend2nd.length > 0 && planSheet2nd) { ... }
}

/**
 * Ensures the Signatory Master Sheet exists and returns it.
 */
function getOrCreateSignatoryMasterSheet(ss) {
    let sheet = ss.getSheetByName(SIGNATORY_MASTER_SHEET);
    if (sheet) {
        return sheet;
    }

    try {
        sheet = ss.insertSheet(SIGNATORY_MASTER_SHEET);
        const headers = ['Signatory Name'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 200);
        Logger.log(`[getOrCreateSignatoryMasterSheet] Created Signatory Master sheet: ${SIGNATORY_MASTER_SHEET}`);
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${SIGNATORY_MASTER_SHEET}" already exists`)) {
             Logger.log(`[getOrCreateSignatoryMasterSheet] WARN: Transient sheet creation failure, retrieving existing sheet.`);
            return ss.getSheetByName(SIGNATORY_MASTER_SHEET);
        }
        throw e;
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
        // NOTE: We rely on parseInt to handle the padded string (e.g., '000001' -> 1)
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
 * @param {string} subProperty The manual Sub Property input from the user.
 * @param {string} sfcRef 
 * @param {object} contractInfo 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @returns {string} The generated 6-digit Padded Reference Number.
 */
function logPrintAction(subProperty, sfcRef, contractInfo, year, month, shift) {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);
        
        const nextRefNum = getNextReferenceNumber(logSheet);
        
        // --- CRITICAL FIX: Convert integer to 6-digit zero-padded string ---
        const paddedRefNum = String(nextRefNum).padStart(6, '0');
        Logger.log(`[logPrintAction] Generated Print Reference Number: ${paddedRefNum} for ${sfcRef}.`);
        return paddedRefNum;
// Return the padded string
    } catch (e) {
        Logger.log(`[logPrintAction] FATAL ERROR: ${e.message}`);
        throw new Error(`Failed to generate print reference number. Error: ${e.message}`);
    }
}


/**
 * UPDATED FUNCTION: Records the actual print log entry using the pre-generated Ref # and locks printed IDs.
 * @param {string} refNum (Now expected to be the 6-digit padded string)
 * @param {string} subProperty The manual Sub Property input from the user.
 * @param {string} sfcRef 
 * @param {object} contractInfo 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @param {string[]} printedPersonnelIds The list of IDs successfully printed.
 */
function recordPrintLogEntry(refNum, subProperty, sfcRef, contractInfo, year, month, shift, printedPersonnelIds) {
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
        
        // 1. Prepare the Log Entry (11 Columns)
        const logEntry = [
            refNum, 
            sfcRef,
            planSheetName, 
            dateRange,     
           
            contractInfo.payor,
            contractInfo.agency,
            subProperty,         // <--- NEW VALUE
            contractInfo.serviceType,
            Session.getActiveUser().getEmail(), 
            new Date(),
            printedPersonnelIds.join(',') // 
            // Col K: IDs joined by comma
      
        ];
        const lastLoggedRow = logSheet.getLastRow();
        const newRow = lastLoggedRow + 1;
        const LOCKED_IDS_COL = LOG_HEADERS.length;
        // 11

        const logEntryRange = logSheet.getRange(newRow, 1, 1, LOG_HEADERS.length);
        // 2. *** CRITICAL FIX: I-set ang format ng Locked Personnel IDs cell (Col K) sa Plain Text ('@') ***
        logEntryRange.getCell(1, LOCKED_IDS_COL).setNumberFormat('@');
        // Tiyakin na ang Ref # column ay naka-set sa Plain Text
        logEntryRange.getCell(1, 1).setNumberFormat('@');
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
 * @param {string[]} personnelIds - Array of Personnel IDs
 * @param {string[]} lockedRefNums - Array of Reference Numbers
 * @param {string[]} personnelNames - Array of Personnel Names <--- NEW
 * @param {string} requesterEmail
 */
function sendRequesterNotification(status, personnelIds, lockedRefNums, personnelNames, requesterEmail) {
  if (requesterEmail === 'UNKNOWN_REQUESTER' || !requesterEmail) return;
  const totalCount = personnelIds.length;
  
  // CRITICAL FIX 1: Get unique reference numbers and SORT them for the subject line
  const uniqueRefNums = [...new Set(lockedRefNums)].sort();
  const subject = `Unlock Request Status: ${status} for ${totalCount} Personnel Schedules (Ref# ${uniqueRefNums.join(', ')})`;
  // CRITICAL FIX 2: Combine and Sort the data for the body list
  const combinedRequests = personnelIds.map((id, index) => ({
    id: id,
    ref: lockedRefNums[index],
    name: personnelNames[index] // <--- NEW: Include Name
  }));
  combinedRequests.sort((a, b) => a.ref.localeCompare(b.ref)); // Sort by Ref#

  // Create the sorted list for email body
  const idList = combinedRequests.map(item => 
    `<li><b>${item.name}</b> (ID ${item.id}) (Ref #: ${item.ref})</li>` // <--- NEW: Display Name first
  ).join('');
  let body = '';
  if (status === 'APPROVED') {
    body = `
      Good news!
      Your request to unlock the following ${totalCount} schedules has been **APPROVED** by the Admin.
      <ul style="list-style-type: none; padding-left: 0; font-weight: bold;">${idList}</ul>
      
      You may now return to the Attendance Plan Monitor app and refresh your browser to edit the schedules.
      ---
      This notification confirms the lock is removed.
    `;
  } else if (status === 'REJECTED') {
    body = `
      Your request to unlock the following ${totalCount} schedules has been **REJECTED** by the Admin.
      <ul style="list-style-type: none; padding-left: 0; font-weight: bold;">${idList}</ul>
      
      The print locks remain active, and the schedules cannot be edited at this time.
      Please contact your Admin for details.
      ---
      This is an automated notification.
    `;
  } else {
      return; 
  }
  
  try {
    MailApp.sendEmail({
      to: requesterEmail,
      subject: subject,
      htmlBody: body, // Changed to htmlBody for list formatting
      name: 'Attendance Plan Monitor (Status Update)'
    });
    Logger.log(`[sendRequesterNotification] Status ${status} email sent to requester: ${requesterEmail} for ${totalCount} IDs.`);
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
    const range = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length);
    const values = range.getValues(); 
    const rangeToUpdate = [];
    values.forEach((row, rowIndex) => {
      const rowNumInSheet = rowIndex + 2; 
      const planSheetNameInLog = String(row[PLAN_SHEET_NAME_COL_INDEX] || '').trim();
      const lockedIdsString = String(row[LOCKED_IDS_COL_INDEX] || '').trim();
      
      if (planSheetNameInLog === planSheetName && lockedIdsString) {
        
        // Get list of IDs currently in the column (which may include UNLOCKED: prefixes)
        const currentIdsWithPrefix = lockedIdsString.split(',').map(id => id.trim());
   
        let updatedLockedIds = [...currentIdsWithPrefix];
        let changed = false;

        personnelIdsToUnlock.forEach(unlockId => {
          
          // 1. Find the index of the locked ID (without prefix)
          const lockedIndex = updatedLockedIds.indexOf(unlockId);
          
          if (lockedIndex > -1) {
 
            // 2. Remove the locked ID from the list
             updatedLockedIds.splice(lockedIndex, 1); 
             
             // 3. Add the UNLOCKED prefix to the original ID (to preserve history)
             //    NOTE: We must ensure the prefix is 
              const unlockedPrefixId = `UNLOCKED:${unlockId}`;
              if (!updatedLockedIds.includes(unlockedPrefixId)) {
                updatedLockedIds.push(unlockedPrefixId);
              }
             
             changed = true;
          }
        });
        
        if (changed) {
          // Filter out any accidental empty strings and join
          const newLockedIdsString = updatedLockedIds.filter(id => id.length > 0).join(',');
          rangeToUpdate.push({
              row: rowNumInSheet,
              col: LOCKED_IDS_COL_INDEX + 1, // 1-based
              value: newLockedIdsString
          });
        }
      }
    });
    rangeToUpdate.forEach(update => {
        const targetRange = logSheet.getRange(update.row, update.col);
        const newValue = update.value;
        targetRange.setNumberFormat('@').setValue(newValue);
    });
    Logger.log(`[unlockPersonnelIds] Successfully unlocked ${personnelIdsToUnlock.length} IDs for ${planSheetName}. (History preserved with UNLOCKED: prefix.)`);
}


/**
 * UPDATED FUNCTION: Sends an HTML email notification to Admin(s) for an unlock request,
 * including actionable Approve/Reject buttons via Web App URL.
 * @param {string} sfcRef 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @param {string[]} personnelIds - Array of clean Personnel IDs.
 * @param {string[]} lockedRefNums - Array of Reference Numbers (padded strings).
 * @param {string[]} personnelNames - Array of Personnel Names <--- NEW
 */
function requestUnlockEmailNotification(sfcRef, year, month, shift, personnelIds, lockedRefNums, personnelNames) { // <--- UPDATED SIGNATURE
  // CRITICAL: personnelIds and lockedRefNums are now arrays
  const requestingUserEmail = Session.getActiveUser().getEmail();
  const adminEmails = ADMIN_EMAILS.join(', ');
  const date = new Date(year, month, 1);
  const planPeriod = date.toLocaleString('en-US', { month: 'long', year: 'numeric' });
  const shiftDisplay = (shift === '1stHalf' ? '1st to 15th' : '16th to End');
  // --- NEW: Combine, Sort, and Prepare lists for email content ---
  const combinedRequests = personnelIds.map((id, index) => ({
    id: id,
    ref: lockedRefNums[index],
    name: personnelNames[index] // <--- NEW: Include Name
  }));
  // Sort the requests sequentially by Reference Number (ascending)
  combinedRequests.sort((a, b) => a.ref.localeCompare(b.ref));
  const requestDetails = combinedRequests.map(item => {
    return `<li style="font-size: 14px;"><b>${item.name}</b> (ID ${item.id}) (Ref #: ${item.ref})</li>`; // <--- NEW: Display Name first
  }).join('');
  // CRITICAL FIX: Get unique reference numbers and SORT them for the subject line
  const uniqueRefNums = [...new Set(lockedRefNums)].sort();
  const subjectRefNums = uniqueRefNums.join(', ');
  const subject = `ATTN: Admin Unlock Request - Ref# ${subjectRefNums} for ${sfcRef}`;
  // **NEW: Encode ALL IDs and Refs in the URL as comma-separated strings**
  // Note: We encode the UNSORTED original arrays for URL consistency, but the list display is sorted.
  const idsEncoded = encodeURIComponent(personnelIds.join(','));
  const refsEncoded = encodeURIComponent(lockedRefNums.join(','));
  const requesterEmailEncoded = encodeURIComponent(requestingUserEmail);
  // Build the Web App URL
  const webAppUrl = ScriptApp.getService().getUrl();
  // CRITICAL: We include ALL IDs/Refs/req_email in the URL
  const unlockUrl = `${webAppUrl}?action=unlock&sfc=${sfcRef}&yr=${year}&mon=${month + 1}&shift=${shift}&id=${idsEncoded}&ref=${refsEncoded}&req_email=${requesterEmailEncoded}`;
  const rejectUrl = `${webAppUrl}?action=reject_info&sfc=${sfcRef}&id=${idsEncoded}&ref=${refsEncoded}&req_email=${requesterEmailEncoded}`;
  const htmlBody = `
    <p style="font-size: 14px;">An Attendance Plan Unlock Request has been submitted.</p> 

    <hr style="margin: 10px 0;">
    
    <p style="font-size: 14px;"><b>Requested By:</b> ${requestingUserEmail}</p>
    <p style="font-size: 14px;"><b>SFC Ref #:</b> ${sfcRef}</p>
    <p style="font-size: 14px;"><b>Plan Period:</b> ${planPeriod} (${shiftDisplay} Half)</p>
    <h4 style="color: #1e40af; font-size: 16px; margin-top: 15px;">Personnel IDs Requested for Unlock (${personnelIds.length}):</h4>
    <ul style="list-style-type: none; padding-left: 0;">
        ${requestDetails} </ul>
    
    <hr style="margin: 10px 0;">

  
    <h3 style="color: #1e40af;">Admin Action Required:</h3>
    
    <div style="margin-top: 15px;">
        <a href="${unlockUrl}" target="_blank" 
           style="background-color: #10b981; color: white; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; font-weight: bold; margin-right: 10px;">
           ✅ APPROVE & UNLOCK ALL (${personnelIds.length})
        </a>
        
        <a href="${rejectUrl}" target="_blank" 
 
          style="background-color: #f59e0b;
           color: white; padding: 10px 20px; text-align: center;
           text-decoration: none; display: inline-block; border-radius: 5px;
           font-weight: bold;">
           ❌ REJECT (Log Only)
        </a>
    </div>

    <p style="margin-top: 20px;
    font-size: 12px; color: #6b7280;">Ang pag-Approve ay magre-remove ng print lock. Kailangan naka-login ka bilang Admin user upang gumana ang link.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: adminEmails,
      subject: subject, // Use the updated subject
      htmlBody: htmlBody, 
      name: 'Attendance Plan Monitor (Automated Request)'
    });
    Logger.log(`[requestUnlockEmailNotification] Sent request email for ${personnelIds.length} IDs to ${adminEmails}`);
    return { success: true, message: `Unlock request sent to Admin(s): ${adminEmails} for 
    ${personnelIds.length} IDs.` };
  } catch (e) {
    Logger.log(`[requestUnlockEmailNotification] Failed to send email: ${e.message}`);
    return { success: false, message: `WARNING: Failed to send request email. Error: ${e.message}` };
  }
}


/**
 * Handles the unlock action triggered by the URL/email link.
 */
function processAdminUnlockFromUrl(params) {
  // CRITICAL: IDs and Refs are now comma-separated strings
  const idsString = params.id ?
  decodeURIComponent(params.id) : '';
  const refsString = params.ref ? decodeURIComponent(params.ref) : '';
  // Split the strings into arrays
  const personnelIds = idsString.split(',').map(s => s.trim()).filter(s => s);
  // Array of IDs
  const lockedRefNums = refsString.split(',').map(s => s.trim()).filter(s => s);
  // Array of Ref#s
  
  const sfcRef = params.sfc;
  const requesterEmail = params.req_email ? decodeURIComponent(params.req_email) : 'UNKNOWN_REQUESTER';
  // NEW: Fetch names corresponding to the IDs
  const personnelNames = personnelIds.map(id => getEmployeeNameFromMaster(sfcRef, id));
  if (personnelIds.length === 0 || lockedRefNums.length === 0 || personnelIds.length !== lockedRefNums.length) {
     return HtmlService.createHtmlOutput('<h1 style="color: red;">INVALID REQUEST</h1><p>The Unlock URL is incomplete or the number of Personnel IDs does not match the number of Reference Numbers.</p>');
  }

  // 1. Check Admin authorization (Crucial for security)
  const userEmail = Session.getActiveUser().getEmail();
  if (!ADMIN_EMAILS.includes(userEmail)) {
    return HtmlService.createHtmlOutput('<h1 style="color: red;">AUTHORIZATION FAILED</h1><p>You are not authorized to perform this action. Your email: ' + userEmail + '</p>');
  }
  
  // Prepare string for status message
  const summary = `${personnelIds.length} schedules (Ref# ${lockedRefNums.join(', ')})`;
  // Handle Reject (Informational) action
  if (params.action === 'reject_info') {
      // 1. Send notification to the original requester (REJECTED)
      sendRequesterNotification('REJECTED', personnelIds, lockedRefNums, personnelNames, requesterEmail);
      // <--- PASSING NAMES
      const template = HtmlService.createTemplateFromFile('UnlockStatus');
      template.status = 'INFO';
      template.message = `Admin (${userEmail}) acknowledged the REJECT click for ${summary}. Notification sent to ${requesterEmail}. No data was changed.
      The locks remain active.`;
      return template.evaluate().setTitle('Reject Status');
  }
  
  // Handle Approve (Unlock) action
  if (params.action === 'unlock' && params.yr && params.mon && params.shift) {
      try {
        const year = parseInt(params.yr, 10);
        const month = parseInt(params.mon, 10) - 1; 
        const shift = params.shift;
        // Perform the unlock using the existing logic, passing the array of IDs
        unlockPersonnelIds(sfcRef, year, month, shift, personnelIds);
        // 1. Send notification to the original requester (APPROVED)
        sendRequesterNotification('APPROVED', personnelIds, lockedRefNums, personnelNames, requesterEmail);
        // <--- PASSING NAMES
        
        // Return Success HTML
        const template = HtmlService.createTemplateFromFile('UnlockStatus');
        template.status = 'SUCCESS';
        template.message = `Successfully unlocked ${summary}. The Print Locks have been removed by Admin (${userEmail}).
        Notification sent to ${requesterEmail}.`;
        return template.evaluate().setTitle('Unlock Status');

      } catch (e) {
        // Return Failure HTML
        const template = HtmlService.createTemplateFromFile('UnlockStatus');
        template.status = 'ERROR';
        template.message = `Failed to unlock ${summary}. Error: ${e.message}`;
        return template.evaluate().setTitle('Unlock Status');
      }
  }
  
  return HtmlService.createHtmlOutput('<h1>Invalid Action</h1><p>The URL provided is incomplete or incorrect.</p>');
}
