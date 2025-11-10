// --- CONFIGURATION: PALITAN ITO NG ID NG INYONG SPREADSHEET at ANG MGA SHEET NAME ---
const SPREADSHEET_ID = '1rQnJGqcWcEBjoyAccjYYMOQj7EkIu1ykXTMLGFzzn2I'; // Main/Master ID (Source)
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; // Target ID for Plans/Employees (Destination)
const CONTRACTS_SHEET_NAME = 'MASTER';
// ------------------------------------------------------------------

// TANDAAN: Para sa MASTER sheet, ang headers ay nagsisimula sa Row 5.
// Para sa ibang sheets (Employees/Plan), ang headers ay nagsisimula sa Row 1.
const MASTER_HEADER_ROW = 5;
// CRITICAL FIX: Baguhin ang Header Row para sa Plan Sheets mula 5 patungong 6
const PLAN_HEADER_ROW = 6; 

/**
 * ENTRY POINT for the Web App.
 * Ito ang function na tinatawag ng Google Apps Script kapag iniload ang URL ng Web App.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Attendance Plan Monitor');
}

/**
 * Helper function para i-include ang iba pang HTML files (Stylesheet.html, JavaScript.html).
 * @param {string} filename Ang pangalan ng HTML file na ii-include (walang .html extension).
 * @return {string} Ang nilalaman ng HTML file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// CRITICAL NEW HELPER: Universal Header Sanitizer
function sanitizeHeader(header) {
    if (!header) return '';
    // Alisin ang lahat ng non-alphanumeric characters (spaces, periods, newlines, etc.)
    return String(header).replace(/[^A-Za-z0-9]/g, '');
}


/**
 * Kinuha ang data mula sa isang sheet, ginagamit ang unang row bilang headers.
 * @param {string} spreadsheetId Ang ID ng Spreadsheet na babasahin.
 * @param {string} sheetName Ang pangalan ng sheet.
 * @return {Array<Object>} Array ng objects, kung saan ang key ay ang header name.
 */
function getSheetData(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId); // Gumagamit na ng dynamic ID
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  let startRow = 1;
  let numRows = sheet.getLastRow();
  let numColumns = sheet.getLastColumn();

  // SPECIAL CASE: Kung MASTER sheet, magsimula sa Row 5.
  if (sheetName === CONTRACTS_SHEET_NAME) {
    startRow = MASTER_HEADER_ROW; // 5
    // Tiyakin na may data na babasahin mula sa startRow
    if (numRows < startRow) {
      Logger.log(`[getSheetData] MASTER sheet has no data starting from Row ${startRow}.`);
      return [];
    }
    // Ayusin ang numRows para sa range (getLastRow() - startRow + 1)
    numRows = sheet.getLastRow() - startRow + 1;
  } 
  // CRITICAL FIX: Add a special case for Employees sheets. (Plan sheets will be accessed by dedicated name)
  else if (sheetName.includes('Employees')) {
      startRow = 1;
  }
  // CRITICAL FIX: Add a special case for AttendancePlan sheets (Horizontal Save).
  else if (sheetName.includes('AttendancePlan')) {
      startRow = PLAN_HEADER_ROW; // FIXED to Row 6
      // Ayusin ang numRows para sa range (getLastRow() - startRow + 1)
      if (numRows < startRow) {
          numRows = 0; // Walang data
      } else {
          numRows = sheet.getLastRow() - startRow + 1;
      }
  }


  // Kung walang data, bumalik na
  if (numRows <= 0 || numColumns === 0) return [];
  // Kumuha ng values, simula sa tamang row at column 1
  const range = sheet.getRange(startRow, 1, numRows, numColumns);
  // CRITICAL FIX: Gamitin ang getDisplayValues() imbes na getValues() para sa STRING consistency
  const values = range.getDisplayValues();
  // Ang unang row ng values array ay ang headers (Row 5 sa Sheets o Row 1)
  const headers = values[0];
  // DEBUGGING LOG: I-log ang raw headers na nabasa ng script
  Logger.log(`[getSheetData] Raw Headers read from ${sheetName} (Starting Row: ${startRow}): ${headers.join(' | ')}`);
  // Linisin ang Headers at I-store para sa mabilis na lookup
  const cleanHeaders = headers.map(header => (header || '').toString().trim());
  Logger.log(`[getSheetData] Cleaned Headers: ${cleanHeaders.join(' | ')}`);

  const data = [];
  // Magsimula sa values[1] dahil values[0] ay ang headers
  for (let i = 1; i < values.length; i++) { 
    const row = values[i];
    // Tiyakin na ang row ay may laman (hindi puro blanko)
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

/**
 * Dynamic Sheet Naming
 * Gumagawa ng sheet name para sa Employees (per SFC Ref#) O Attendance Plan (per SFC Ref#, Month, Shift).
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {string} type 'employees' o 'plan'.
 * @param {number} [year] Ang taon (kinakailangan kung 'plan').
 * @param {number} [month] Ang buwan (0-based, kinakailangan kung 'plan').
 * @param {string} [shift] '1stHalf' o '2ndHalf' (kinakailangan kung 'plan').
 * @return {string} Ang pangalan ng sheet.
 */
function getDynamicSheetName(sfcRef, type, year, month, shift) {
    const safeRef = (sfcRef || '').replace(/[\\/?*[]/g, '_');

    if (type === 'employees') {
        return `${safeRef} - Employees`;
    }
    
    if (type === 'plan' && year !== undefined && month !== undefined && shift) {
        // FIX: Gumawa ng Date object para makuha ang Month Name
        const tempDate = new Date(year, month, 1);
        const monthName = tempDate.toLocaleString('en-US', { month: 'short' });
        // Halimbawa: 2308 - Nov 2025 - 1stHalf AttendancePlan
        return `${safeRef} - ${monthName} ${year} - ${shift} AttendancePlan`;
    }
    
    // Fallback sa generic plan name (ginagamit lang sa updatePlanKeysOnIdChange/saveEmployeeInfoBulk)
    return `${safeRef} - AttendancePlan`; 
}

/**
 * Tinitiyak na ang Employee at Attendance Plan sheets para sa Contract ID ay existing at may tamang headers.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {number} year Ang taon.
 * @param {number} month Ang buwan (0-based).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 * @return {boolean} True kung existing ang BOTH sheets, False kung hindi.
 */
function checkContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef || year === undefined || month === undefined || !shift) return false;
    
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
        const empSheetName = getDynamicSheetName(sfcRef, 'employees');
        // GUMAGAMIT NA NG NEW DYNAMIC SHEET NAME PARA SA PLAN
        const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift); 
        
        // Employees sheet ay laging kailangan, Plan sheet lang ang kailangan i-check for the specific month/shift
        return !!ss.getSheetByName(empSheetName) && !!ss.getSheetByName(planSheetName);
    } catch (e) {
         Logger.log(`[checkContractSheets] ERROR: Failed to open Spreadsheet ID ${TARGET_SPREADSHEET_ID}. Check ID and permissions. Error: ${e.message}`);
         return false;
    }
}


// --- NEW HELPER FUNCTION: To append employee rows to a newly created plan sheet ---
function appendExistingEmployeeRowsToPlan(sfcRef, planSheet, shiftToAppend) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empSheet = ss.getSheetByName(empSheetName);

    if (!empSheet) return;

    // Kunin ang lahat ng Personnel ID mula sa Employee Sheet
    const empData = getSheetData(TARGET_SPREADSHEET_ID, empSheetName);
    const existingIds = empData.map(e => cleanPersonnelId(e['Personnel ID'])).filter(id => id);

    if (existingIds.length === 0) return;

    Logger.log(`[appendExistingEmployeeRowsToPlan] Found ${existingIds.length} existing employees. Populating plan sheet.`);

    const planHeadersCount = 33; // Personnel ID (1), Shift (1), Day 1-31 (31)
    const rowsToAppend = [];
    
    existingIds.forEach(id => {
        // Row for the specific shift
        if (shiftToAppend === '1stHalf') {
            const planRow1 = Array(planHeadersCount).fill('');
            planRow1[0] = id; // Personnel ID
            planRow1[1] = '1stHalf'; // Shift
            rowsToAppend.push(planRow1);
        }
        
        // Row for the specific shift
        if (shiftToAppend === '2ndHalf') {
            const planRow2 = Array(planHeadersCount).fill('');
            planRow2[0] = id; // Personnel ID
            planRow2[1] = '2ndHalf'; // Shift
            rowsToAppend.push(planRow2);
        }
    });

    if (rowsToAppend.length > 0) {
        planSheet.getRange(planSheet.getLastRow() + 1, 1, rowsToAppend.length, planHeadersCount).setValues(rowsToAppend);
        Logger.log(`[appendExistingEmployeeRowsToPlan] Successfully pre-populated ${rowsToAppend.length} plan rows for ${shiftToAppend}.`);
    }
}
// ---------------------------------------------------------------------------------


/**
 * Gumagawa ng Employee at Attendance Plan sheets para sa Contract ID, kasama ang tamang headers.
 * Ngayon ay tumatanggap ng Month/Shift para sa Plan Sheet.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {number} year Ang taon.
 * @param {number} month Ang buwan (0-based).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 */
function createContractSheets(sfcRef, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    
    // --- EMPLOYEES SHEET (STATIC NAME) ---
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
    
    // --- ATTENDANCE PLAN SHEET (DYNAMIC NAME) ---
    // GUMAGAMIT NA NG NEW DYNAMIC SHEET NAME PARA SA PLAN
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift); 
    let planSheet = ss.getSheetByName(planSheetName);

    // CRITICAL FIX: NEW HEADER GENERATION
    const getHorizontalPlanHeaders = (sheetYear, sheetMonth) => {
        const base = ['Personnel ID', 'Shift'];
        
        for (let d = 1; d <= 31; d++) {
            const currentDate = new Date(sheetYear, sheetMonth, d);
            // Check if date is valid for the month
            if (currentDate.getMonth() === sheetMonth) {
                const monthShortRaw = currentDate.toLocaleString('en-US', { month: 'short' });
                // CRITICAL FIX: I-UPPERCASE ang unang letter ng month name AT ALISIN ANG TULDOK
                const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '');
                
                base.push(`${monthShort}${d}`); 
            } else {
                // Para panatilihin ang 31 column position, gagamitin natin ang Day X
                base.push(`Day${d}`); 
            }
        }
        return base;
    };

    if (!planSheet) {
        planSheet = ss.insertSheet(planSheetName);
        planSheet.clear();
        
        // Mag-reserve ng space para sa Contract Info (Row 1-5)
        // CRITICAL FIX: Headers ay nasa Row 6
        const planHeaders = getHorizontalPlanHeaders(year, month); // Pass Year and Month for dynamic date headers
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, planHeaders.length).setValues([planHeaders]);
        planSheet.setFrozenRows(PLAN_HEADER_ROW); // I-freeze ang headers simula Row 6
        
        Logger.log(`[createContractSheets] Created Horizontal Attendance Plan sheet for ${planSheetName} with headers at Row ${PLAN_HEADER_ROW}.`);
        
        // CRITICAL FIX 1: Populate the newly created sheet with existing employee rows
        appendExistingEmployeeRowsToPlan(sfcRef, planSheet, shift); // PASS SHIFT
        
        // CRITICAL FIX 3: Aesthetic Fix - Hide irrelevant columns
        // Column Index: ID (1), Shift (2), Day 1 (3), Day 15 (17), Day 16 (18), Day 31 (33)
        if (shift === '1stHalf') {
            // HIDE Day 16 (Column 18) hanggang Day 31 (Column 33)
            const START_COL_TO_HIDE = 18; 
            const NUM_COLS_TO_HIDE = 16; 
            planSheet.hideColumns(START_COL_TO_HIDE, NUM_COLS_TO_HIDE);
            Logger.log(`[createContractSheets] Hiding Day 16-31 columns for 1stHalf sheet.`);
        } else if (shift === '2ndHalf') {
            // HIDE Day 1 (Column 3) hanggang Day 15 (Column 17)
            const START_COL_TO_HIDE = 3; 
            const NUM_COLS_TO_HIDE = 15; 
            planSheet.hideColumns(START_COL_TO_HIDE, NUM_COLS_TO_HIDE);
            Logger.log(`[createContractSheets] Hiding Day 1-15 columns for 2ndHalf sheet.`);
        }
    } 
}


/**
 * Ito na ngayon ay isang internal helper, ginagamit LAMANG bago mag-save (saveAllData)
 * upang tiyakin na may sheet na mapagsa-save-an.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {number} year Ang taon.
 * @param {number} month Ang buwan (0-based).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 */
function ensureContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required to ensure sheets.");
    
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    // GUMAGAMIT NA NG NEW DYNAMIC SHEET NAME PARA SA PLAN
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift); 

    let needsPlanCreation = false;
    
    // Check Employee Sheet
    if (!ss.getSheetByName(empSheetName)) {
        createContractSheets(sfcRef, year, month, shift); 
        Logger.log(`[ensureContractSheets] Created Employee sheet for ${sfcRef}.`);
        needsPlanCreation = true;
    }
    
    // Check Plan Sheet (for the specific month/shift)
    if (!ss.getSheetByName(planSheetName)) {
        // We only create the Plan sheet here (dynamic name)
        createContractSheets(sfcRef, year, month, shift); 
        Logger.log(`[ensureContractSheets] Created new Plan sheet for ${planSheetName}.`);
        needsPlanCreation = true;
    } 
}


/**
 * Kinukuha ang listahan ng LIVE na Kontrata mula sa MASTER sheet.
 */
function getContracts() {
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE' || !SPREADSHEET_ID) {
    throw new Error("CONFIGURATION ERROR: Pakipalitan ang 'YOUR_SPREADSHEET_ID_HERE' sa Code.gs ng tamang Spreadsheet ID.");
  }
    
  // GUMAGAMIT NG SPREADSHEET_ID para sa MASTER sheet
  const allContracts = getSheetData(SPREADSHEET_ID, CONTRACTS_SHEET_NAME);

  // Helper function para mahanap ang case-insensitive key
  const findKey = (c, search) => {
      const keys = Object.keys(c);
      return keys.find(key => (key || '').trim().toLowerCase() === search.toLowerCase());
  };
    
  const filteredContracts = allContracts.filter((c, index) => {
    // ... (filtering logic remains the same)
    const rowNumber = MASTER_HEADER_ROW + index + 1; 

    const statusKey = findKey(c, 'Status of SFC');
    const contractIdKey = findKey(c, 'CONTRACT GRP ID');
    
    if (!statusKey || !contractIdKey) return false;

    const contractIdValue = (c[contractIdKey] || '').toString().trim();
    if (!contractIdValue) return false; 
    
    const status = (c[statusKey] || '').toString().trim().toLowerCase();
    const isLive = status === 'live' || status === 'on process - live';

    return isLive;
  });

  // MAPPING: Ginagamit ang case-insensitive keys
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
      sfcRef: sfcRefKey ? (c[sfcRefKey] || '').toString() : '', // CRITICAL: Nagdagdag ng SFC Ref#
    };
  });
}

// Helper function for aggressive ID cleaning
function cleanPersonnelId(rawId) {
    let idString = String(rawId || '').trim();
    // Remove all non-digit characters (including spaces, dashes, etc.)
    return idString.replace(/\D/g, ''); 
}


/**
 * Kinukuha ang Employee List at Attendance Plan para sa isang Contract ID.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {number} year Ang taon na kino-convert (e.g., 2025).
 * @param {number} month Ang buwan (0-based, e.g., 10 para sa Nov).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 */
function getAttendancePlan(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    
    // 1. Kumuha ng Employee Data (GUMAGAMIT NG TARGET_SPREADSHEET_ID)
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empData = getSheetData(TARGET_SPREADSHEET_ID, empSheetName);
    
    // 2. Kumuha ng Attendance Plan Data (GUMAGAMIT NG TARGET_SPREADSHEET_ID)
    // GUMAGAMIT NG NEW DYNAMIC SHEET NAME
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(planSheetName);
    
    if (!planSheet) return { employees: empData, planMap: {} };

    // CRITICAL: Read all data rows starting from Row 6
    const HEADER_ROW = PLAN_HEADER_ROW; // FIXED to 6
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();

    if (numRowsToRead <= 0 || numColumns < 33) { 
        // 33 columns = ID, Shift, Day 1 to Day 31
        return { employees: empData, planMap: {} };
    }

    // Read headers (Row 6) and data (Row 7 onwards)
    const planValues = planSheet.getRange(HEADER_ROW, 1, numRowsToRead + 1, numColumns).getValues(); // Use getValues() for raw strings
    const headers = planValues[0];
    const dataRows = planValues.slice(1);

    const personnelIdIndex = headers.indexOf('Personnel ID');
    const shiftIndex = headers.indexOf('Shift');
    
    const planMap = {};
    
    // CRITICAL FIX 1: Gumawa ng sanitized map ng headers
    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        // Linisin ang header sa pamamagitan ng pag-alis ng lahat ng non-alphanumeric
        const cleanedHeader = sanitizeHeader(header); 
        sanitizedHeadersMap[cleanedHeader] = index;
    });
    // END CRITICAL FIX 1
    
    
    // Horizontal to Vertical Conversion (Focusing only on the current shift)
    dataRows.forEach((row, rowIndex) => {
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        const currentShift = String(row[shiftIndex] || '').trim();
        
        // CHECK: Tanging ang row lang na tumutugma sa kasalukuyang shift ang i-ko-convert
        if (currentShift === shift) {
            
            // Mag-loop sa Day 1 hanggang Day 31
            for (let d = 1; d <= 31; d++) {
                // FIX: Gamitin ang year at month na galing sa client
                const dayKey = `${year}-${month + 1}-${d}`; // Month + 1 para sa 1-based month
                
                // CRITICAL FIX: Hanapin ang header gamit ang "MonthDay" format
                const date = new Date(year, month, d);
                
                let lookupHeader = '';
                if (date.getMonth() === month) {
                    const monthShortRaw = date.toLocaleString('en-US', { month: 'short' });
                    // CRITICAL FIX: I-UPPERCASE ang unang letter ng month name AT ALISIN ANG TULDOK AT SPACE
                    const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '').replace(/\s/g, '');
                    
                    // FIX: Ito ang tamang format na mayroon tayo: MonthDay (e.g., Nov15)
                    lookupHeader = `${monthShort}${d}`; 
                } else {
                    // Fallback para sa Day X columns (Hindi dapat ma-trigger, pero safety)
                    lookupHeader = `Day${d}`; 
                }
                
                // CRITICAL FIX 2: Gamitin ang sanitized map
                const dayIndex = sanitizedHeadersMap[sanitizeHeader(lookupHeader)];
                
                if (dayIndex !== undefined && id && currentShift) {
                    // Gamitin ang DisplayValue para makuha ang formatted schedule (e.g., 08:00-17:00)
                    const statusRange = planSheet.getRange(HEADER_ROW + rowIndex + 1, dayIndex + 1);
                    const status = String(statusRange.getDisplayValue() || '').trim();
                    
                    const key = `${id}_${dayKey}_${currentShift}`;
                    
                    if (status) {
                         planMap[key] = status;
                    }
                }
            }
        }
    });
    
    // 3. I-organisa ang Employee Data (Kasama ang 'No.')
    const employees = empData.map((e, index) => {
        // CRITICAL FIX: AGGRESSIVE ID CLEANING
        const id = cleanPersonnelId(e['Personnel ID']);
            
        return {
           no: index + 1, 
            id: id, // CLEANED ID
            name: String(e['Personnel Name'] || '').trim(),
            position: String(e['Position'] || '').trim(),
            area: String(e['Area Posting'] || '').trim(),
        }
    }).filter(e => e.id);
    
    return { employees, planMap };
}

/**
 * Ina-update ang Personnel ID sa AttendancePlan sheet kung nagbago ang ID.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {Array<Object>} employeeChanges Array of {id (new ID), name, position, area, isNew, oldPersonnelId}
 */
function updatePlanKeysOnIdChange(sfcRef, employeeChanges) {
    // ... (This function is complex due to the new dynamic naming and is currently skipped, focusing on the main save issue)
    Logger.log("[updatePlanKeysOnIdChange] Skipped Plan Sheet ID update for ID change due to new dynamic naming.");
}


/**
 * Ina-update ang maramihang entry sa Employee at Attendance Plan sheets, 
 * at sine-save ang Contract Info sa unang 4 rows ng Plan Sheet.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {Object} contractInfo Contract details to save in the plan sheet (Payor, Agency, etc.)
 * @param {Array<Object>} employeeChanges Mga pagbabago sa Employee Info.
 * @param {Array<Object>} attendanceChanges Mga pagbabago sa Attendance Plan.
 * @param {number} year Ang taon.
 * @param {number} month Ang buwan (0-based).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 */
function saveAllData(sfcRef, contractInfo, employeeChanges, attendanceChanges, year, month, shift) {
    Logger.log(`[saveAllData] Starting save for SFC Ref#: ${sfcRef}, Month/Shift: ${month}/${shift}`);
    if (!sfcRef) {
      throw new Error("SFC Ref# is required.");
    }
    // GUMAGAMIT NA NG YEAR, MONTH, SHIFT SA ensure
    ensureContractSheets(sfcRef, year, month, shift); 
    
    // 1. I-save ang Contract Info sa Plan Sheet (for the specific month/shift)
    saveContractInfo(sfcRef, contractInfo, year, month, shift);
    
    // 2. I-save ang Employee Info (Bulk)
    if (employeeChanges && employeeChanges.length > 0) {
        // Pass Year, Month, Shift so that it can find the correct plan sheets for appending new employee rows
        saveEmployeeInfoBulk(sfcRef, employeeChanges, year, month, shift); 
        // updatePlanKeysOnIdChange(sfcRef, employeeChanges); // Disabled for now
    }
    
    // 3. I-save ang Attendance Plan (Bulk)
    if (attendanceChanges && attendanceChanges.length > 0) {
        saveAttendancePlanBulk(sfcRef, attendanceChanges, year, month, shift);
    }
    
    Logger.log(`[saveAllData] Save completed.`);
}


/**
 * Sine-save ang Contract Details sa Row 1-4 ng Attendance Plan sheet.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {Object} info 
 * @param {number} year Ang taon.
 * @param {number} month Ang buwan (0-based).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 */
function saveContractInfo(sfcRef, info, year, month, shift) {
    // GUMAGAMIT NA NG TARGET_SPREADSHEET_ID
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    // GUMAGAMIT NG NEW DYNAMIC SHEET NAME
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift); 
    const planSheet = ss.getSheetByName(planSheetName);
    
    if (!planSheet) throw new Error(`Plan Sheet for ${planSheetName} not found.`);
    
    // Helper para makuha ang start at end date
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
    
    // CRITICAL FIX: Idagdag ang date range sa metadata
    const data = [
        ['PAYOR COMPANY', info.payor],           // Row 1
        ['AGENCY', info.agency],                 // Row 2
        ['SERVICE TYPE', info.serviceType],      // Row 3
        ['TOTAL HEAD COUNT', info.headCount],     // Row 4
        ['PLAN PERIOD', dateRange]               // NEW Row 5
    ];
    
    // 1. Tanggalin ang dating content ng Rows 1-5 (para sa malinis na pag-save)
    planSheet.getRange('A1:B5').clearContent();
    
    // 2. I-set ang values sa A1:B5 (Headers: A, Values: B)
    planSheet.getRange('A1:B5').setValues(data);
    
    // 3. I-update ang frozen rows (Dapat Row 6 na ang header)
    planSheet.setFrozenRows(PLAN_HEADER_ROW);
    
    Logger.log(`[saveContractInfo] Saved metadata and date range for ${planSheetName}.`);
}


/**
 * Ina-update ang maramihang entry sa Attendance Plan sheet (HORIZONTAL SAVE LOGIC).
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {Array<Object>} changes Array of {personnelId, dayKey, shift, status}
 * @param {number} year Ang taon.
 * @param {number} month Ang buwan (0-based).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 */
function saveAttendancePlanBulk(sfcRef, changes, year, month, shift) {
    // GUMAGAMIT NA NG TARGET_SPREADSHEET_ID
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    // GUMAGAMIT NG NEW DYNAMIC SHEET NAME
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift); 
    const planSheet = ss.getSheetByName(planSheetName);

    if (!planSheet) throw new Error(`AttendancePlan Sheet for ${planSheetName} not found.`);
    // TANDAAN: SHIFTED HEADER ROW to 6
    const HEADER_ROW = PLAN_HEADER_ROW; // FIXED to 6
    
    // Tiyakin na ang sheet ay bukas bago magbasa
    planSheet.setFrozenRows(0);
    const lastRow = planSheet.getLastRow();
    
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();
    
    // Kumuha ng buong data set (Headers + Data)
    let values = [];
    let headers = [];
    if (numRowsToRead >= 0 && numColumns > 0) {
         values = planSheet.getRange(HEADER_ROW, 1, numRowsToRead + 1, numColumns).getValues();
         headers = values[0]; 
    } else {
        // Fallback Headers (Should not happen if sheets are created correctly)
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
    
    // CRITICAL FIX 1: Gumawa ng sanitized map ng headers
    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        // Linisin ang header sa pamamagitan ng pag-alis ng lahat ng non-alphanumeric
        const cleanedHeader = sanitizeHeader(header); 
        sanitizedHeadersMap[cleanedHeader] = index;
    });
    
    // Gumawa ng map para sa mabilis na lookup ng existing rows: Key: ID_Shift -> RowIndex (sa values array)
    const rowLookupMap = {}; 
    for (let i = 1; i < values.length; i++) { 
        const rowId = String(values[i][personnelIdIndex] || '').trim();
        const rowShift = String(values[i][shiftIndex] || '').trim();
        if (rowId) {
            rowLookupMap[`${rowId}_${rowShift}`] = i; 
        }
    }
    
    // I-store ang updates: Key: RowIndex (sa values array), Value: {ColIndex: NewStatus}
    const updatesMap = {}; 
    
    changes.forEach(data => {
        const { personnelId, dayKey, shift: dataShift, status } = data;
        
        // Tanging ang changes lang na tumugma sa shift ng sheet ang i-sa-save
        if (dataShift !== shift) return; 
        
        const rowKey = `${personnelId}_${dataShift}`;
        
        // Kunin ang Day Number mula sa DayKey (YYYY-M-D)
        const dayNumber = parseInt(dayKey.split('-')[2], 10);
        
        // CRITICAL FIX 2: Hanapin ang header gamit ang MonthDay format
        const date = new Date(year, month, dayNumber);
        
        let targetLookupHeader = '';
        if (date.getMonth() === month) {
            const monthShortRaw = date.toLocaleString('en-US', { month: 'short' });
            // CRITICAL FIX: I-UPPERCASE ang unang letter ng month name AT ALISIN ANG TULDOK AT SPACE
            const monthShort = (monthShortRaw.charAt(0).toUpperCase() + monthShortRaw.slice(1)).replace('.', '').replace(/\s/g, '');
            
            // FIX: Ito ang tamang format na mayroon tayo: MonthDay (e.g., Nov15)
            targetLookupHeader = `${monthShort}${dayNumber}`; 
        } else {
            // Fallback para sa Day X columns (Hindi dapat ma-trigger, pero safety)
            targetLookupHeader = `Day${dayNumber}`; 
        }
        
        // CRITICAL FIX 3: Maghanap ng index gamit ang Sanitized Map
        const dayColIndex = sanitizedHeadersMap[targetLookupHeader]; 
        
        if (dayColIndex === undefined) {
            // I-log ang FATAL MISS
            Logger.log(`[savePlanBulk] FATAL MISS: Header Lookup '${targetLookupHeader}' failed. Available Sanitized Keys: ${Object.keys(sanitizedHeadersMap).join(' | ')}`);
            return; 
        }

        const rowIndexInValues = rowLookupMap[rowKey]; // Index sa 0-based values array (0 is header)
        
        if (rowIndexInValues !== undefined) {
            // Update existing row
            const sheetRowNumber = rowIndexInValues + HEADER_ROW; // Actual row number sa sheet
            
            // I-store ang update: Sheet Row Number -> [Col Index, New Status]
            if (!updatesMap[sheetRowNumber]) {
                updatesMap[sheetRowNumber] = {};
            }
            updatesMap[sheetRowNumber][dayColIndex + 1] = status; // Col Index + 1 (1-based column number)

        } else {
            // New Employee/Shift combination (This means the Employee Sheet was saved, but the corresponding Plan Row was missing)
            Logger.log(`[savePlanBulk] WARNING: ID/Shift combination not found for row: ${rowKey}. Skipping update for this cell.`);
        }
    });

    // 1. Execute bulk updates
    Object.keys(updatesMap).forEach(sheetRowNumber => {
        const colUpdates = updatesMap[sheetRowNumber];
        
        Object.keys(colUpdates).forEach(colNum => {
            const status = colUpdates[colNum];
            // Update Cell: (Row, Column)
            planSheet.getRange(parseInt(sheetRowNumber), parseInt(colNum)).setValue(status);
        });
    });

    planSheet.setFrozenRows(HEADER_ROW); 
    Logger.log(`[saveAttendancePlanBulk] Completed Horizontal Plan update for ${planSheetName}.`);
}


/**
 * Ina-update ang maramihang entry sa Employee sheet.
 * @param {string} sfcRef Ang SFC Ref# (ang bagong sheet key).
 * @param {Array<Object>} changes Array of {id (new ID), name, position, area, isNew, oldPersonnelId}
 * @param {number} year Ang taon.
 * @param {number} month Ang buwan (0-based).
 * @param {string} shift '1stHalf' o '2ndHalf'.
 */
function saveEmployeeInfoBulk(sfcRef, changes, year, month, shift) {
    // GUMAGAMIT NA NG TARGET_SPREADSHEET_ID
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empSheet = ss.getSheetByName(empSheetName);
    
    // GUMAGAMIT NG NEW DYNAMIC SHEET NAME PARA SA PLAN SHEET
    // Kailangan natin ang sheet name ng 1stHalf at 2ndHalf para makapag-append ng bagong rows ng empleyado
    const planSheetName1st = getDynamicSheetName(sfcRef, 'plan', year, month, '1stHalf'); 
    const planSheet1st = ss.getSheetByName(planSheetName1st);
    const planSheetName2nd = getDynamicSheetName(sfcRef, 'plan', year, month, '2ndHalf'); 
    const planSheet2nd = ss.getSheetByName(planSheetName2nd);

    if (!empSheet) throw new Error(`Employee Sheet for SFC Ref# ${sfcRef} not found.`);
    empSheet.setFrozenRows(0); // Temporarily unfreeze
    
    // Kumuha ng data mula sa Row 1, Col 1 hanggang sa dulo, para makuha ang headers.
    const lastRow = empSheet.getLastRow();
    const numRows = empSheet.getLastRow() > 0 ? empSheet.getLastRow() : 1; 
    const numColumns = empSheet.getLastColumn() > 0 ? empSheet.getLastColumn() : 4;
    // Read values: Dito dapat makita ang headers at ang lahat ng existing data
    const values = empSheet.getRange(1, 1, numRows, numColumns).getValues();
    const headers = values[0]; 
    empSheet.setFrozenRows(1); // Restore freeze
    
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('Position');
    const areaIndex = headers.indexOf('Area Posting');

    const rowsToUpdate = {};
    const rowsToAppend = [];
    
    const planRowsToAppend1st = []; 
    const planRowsToAppend2nd = []; 
    
    // Gumawa ng map para sa mabilis na lookup ng existing rows
    const personnelIdMap = {};
    for (let i = 1; i < values.length; i++) { 
        personnelIdMap[String(values[i][personnelIdIndex] || '').trim()] = i + 1;
    }
    
    changes.forEach((data, changeIndex) => {
        const oldId = String(data.oldPersonnelId || '').trim();
        const newId = String(data.id || '').trim();
        
        // 1. Existing Row Update 
        if (!data.isNew && personnelIdMap[oldId]) { // Use !data.isNew check
             const sheetRowNumber = personnelIdMap[oldId];
            
            if(oldId !== newId && personnelIdMap[newId] && personnelIdMap[newId] !== sheetRowNumber) return; 
            
            rowsToUpdate[sheetRowNumber] = [newId, data.name, data.position, data.area];
            if (oldId !== newId) {
                delete personnelIdMap[oldId];
                personnelIdMap[newId] = sheetRowNumber;
            }

        } 
        // 2. New Row Append 
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
            
            // NEW: Maghanda ng dalawang row para sa Attendance Plan (1stHalf at 2ndHalf)
            const planHeadersCount = 33; // Always assume 33 columns for horizontal save
            
            // Row for 1st Half
            const planRow1 = Array(planHeadersCount).fill('');
            planRow1[0] = newId; // Personnel ID
            planRow1[1] = '1stHalf'; // Shift
            planRowsToAppend1st.push(planRow1);
            
            // Row for 2nd Half
            const planRow2 = Array(planHeadersCount).fill('');
            planRow2[0] = newId; // Personnel ID
            planRow2[1] = '2ndHalf'; // Shift
            planRowsToAppend2nd.push(planRow2);
            
            personnelIdMap[newId] = -1;
        }
    });
    
    // 1. Update existing rows (Employee Sheet)
    Object.keys(rowsToUpdate).forEach(sheetRowNumber => {
        const rowData = rowsToUpdate[sheetRowNumber];
        empSheet.getRange(parseInt(sheetRowNumber), personnelIdIndex + 1, 1, 4).setValues([
            [rowData[0], rowData[1], rowData[2], rowData[3]]
        ]);
    });
    
    // 2. Append new rows (Employee Sheet)
    if (rowsToAppend.length > 0) {
      rowsToAppend.forEach(row => {
          empSheet.appendRow(row); 
      });
    }
    
    // 3. Append new rows (Attendance Plan Sheet - 1st Half)
    // Tanging i-append lang kung may sheet
    if (planRowsToAppend1st.length > 0 && planSheet1st) {
        planSheet1st.getRange(planSheet1st.getLastRow() + 1, 1, planRowsToAppend1st.length, planRowsToAppend1st[0].length).setValues(planRowsToAppend1st);
    }
    
    // 4. Append new rows (Attendance Plan Sheet - 2nd Half)
    // Tanging i-append lang kung may sheet
    if (planRowsToAppend2nd.length > 0 && planSheet2nd) {
        planSheet2nd.getRange(planSheet2nd.getLastRow() + 1, 1, planRowsToAppend2nd.length, planRowsToAppend2nd[0].length).setValues(planRowsToAppend2nd);
    }
}
