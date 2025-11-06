// --- CONFIGURATION: PALITAN ITO NG ID NG INYONG SPREADSHEET at ANG MGA SHEET NAME ---
const SPREADSHEET_ID = '1rQnJGqcWcEBjoyAccjYYMOQj7EkIu1ykXTMLGFzzn2I';
const CONTRACTS_SHEET_NAME = 'MASTER';
// ------------------------------------------------------------------

// TANDAAN: Para sa MASTER sheet, ang headers ay nagsisimula sa Row 5.
// Para sa ibang sheets (Employees/Plan), ang headers ay nagsisimula sa Row 1.
const MASTER_HEADER_ROW = 5;
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

/**
 * Kinuha ang data mula sa isang sheet, ginagamit ang unang row bilang headers.
 * Ngayon, ito ay sumusuporta sa mga header na hindi nagsisimula sa Row 1.
 * @param {string} sheetName Ang pangalan ng sheet.
 * @return {Array<Object>} Array ng objects, kung saan ang key ay ang header name.
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  let startRow = 1;
  let numRows = sheet.getLastRow();
  let numColumns = sheet.getLastColumn();

  // SPECIAL CASE: Kung MASTER sheet, magsimula sa Row 5.
  if (sheetName === CONTRACTS_SHEET_NAME) {
    startRow = MASTER_HEADER_ROW;
    // Tiyakin na may data na babasahin mula sa startRow
    if (numRows < startRow) {
      Logger.log(`[getSheetData] MASTER sheet has no data starting from Row ${startRow}.`);
      return [];
    }
    // Ayusin ang numRows para sa range (getLastRow() - startRow + 1)
    numRows = sheet.getLastRow() - startRow + 1;
  }

  // Kung walang data, bumalik na
  if (numRows <= 0 || numColumns === 0) return [];
  // Kumuha ng values, simula sa tamang row at column 1
  const range = sheet.getRange(startRow, 1, numRows, numColumns);
  const values = range.getValues();
  
  // Ang unang row ng values array ay ang headers (Row 5 sa Sheets)
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
    if (row.some(cell => String(cell).trim() !== '')) { 
        const item = {};
        cleanHeaders.forEach((headerKey, index) => {
          if (headerKey) {
              item[headerKey] = row[index];
          }
        });
        data.push(item);
    }
  }
  
  Logger.log(`[getSheetData] Total data rows processed (excluding header): ${data.length}`);

  return data;
}

/**
 * Dynamic Sheet Naming
 * @param {string} contractId Ang ID ng kontrata.
 * @param {string} type 'employees' o 'plan'.
 * @return {string} Ang pangalan ng sheet.
 */
function getDynamicSheetName(contractId, type) {
    const safeId = (contractId || '').replace(/[\\/?*[]/g, '_');
    // Linisin ang ID
    if (type === 'employees') {
        return `${safeId} - Employees`;
    }
    return `${safeId} - AttendancePlan`;
}

/**
 * Tinitiyak na ang Employee at Attendance Plan sheets para sa Contract ID ay existing at may tamang headers.
 * HINDI na ito auto-creates. Ito ay nagre-return lang ng boolean.
 * @param {string} contractId Ang ID ng kontrata.
 * @return {boolean} True kung existing ang BOTH sheets, False kung hindi.
 */
function checkContractSheets(contractId) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const empSheetName = getDynamicSheetName(contractId, 'employees');
    const planSheetName = getDynamicSheetName(contractId, 'plan');
    
    return !!ss.getSheetByName(empSheetName) && !!ss.getSheetByName(planSheetName);
}

/**
 * Gumagawa ng Employee at Attendance Plan sheets para sa Contract ID, kasama ang tamang headers.
 * @param {string} contractId Ang ID ng kontrata.
 */
function createContractSheets(contractId) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // --- EMPLOYEES SHEET ---
    const empSheetName = getDynamicSheetName(contractId, 'employees');
    let empSheet = ss.getSheetByName(empSheetName);

    if (!empSheet) {
        empSheet = ss.insertSheet(empSheetName);
        empSheet.clear();
        // HEADERS para sa Employees
        const empHeaders = ['Personnel ID', 'Personnel Name', 'Position', 'Area Posting'];
        empSheet.getRange(1, 1, 1, empHeaders.length).setValues([empHeaders]);
        empSheet.setFrozenRows(1);
        Logger.log(`[createContractSheets] Created Employee sheet for ${contractId}`);
    } else {
        Logger.log(`[createContractSheets] Employee sheet for ${contractId} already existed.`);
    }
    
    // --- ATTENDANCE PLAN SHEET ---
    const planSheetName = getDynamicSheetName(contractId, 'plan');
    let planSheet = ss.getSheetByName(planSheetName);

    // Kukunin natin ang attendance plan headers sa isang function para ma-reuse
    const getPlanHeaders = () => ['Personnel ID', 'DayKey', 'Shift', 'Status'];

    if (!planSheet) {
        planSheet = ss.insertSheet(planSheetName);
        planSheet.clear();
        
        // NEW: Mag-reserve ng space para sa Contract Info (Row 1-4)
        // I-set ang Row 5 bilang header ng Attendance Plan
        const planHeaders = getPlanHeaders();
        planSheet.getRange(5, 1, 1, planHeaders.length).setValues([planHeaders]);
        planSheet.setFrozenRows(5); // I-freeze ang headers simula Row 5
        
        Logger.log(`[createContractSheets] Created Attendance Plan sheet for ${contractId} with headers at Row 5.`);
    } else {
        Logger.log(`[createContractSheets] Attendance Plan sheet for ${contractId} already existed.`);
    }
}


/**
 * Ito na ngayon ay isang internal helper, ginagamit LAMANG bago mag-save (saveAllData)
 * upang tiyakin na may sheet na mapagsa-save-an. 
 * @param {string} contractId Ang ID ng kontrata.
 */
function ensureContractSheets(contractId) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const empSheetName = getDynamicSheetName(contractId, 'employees');
    if (!ss.getSheetByName(empSheetName)) {
        createContractSheets(contractId); // Call the creator if missing
        Logger.log(`[ensureContractSheets] Re-created sheets for ${contractId} before saving.`);
    }
}


/**
 * Kinukuha ang listahan ng LIVE na Kontrata mula sa MASTER sheet.
 */
function getContracts() {
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE' || !SPREADSHEET_ID) {
    throw new Error("CONFIGURATION ERROR: Pakipalitan ang 'YOUR_SPREADSHEET_ID_HERE' sa Code.gs ng tamang Spreadsheet ID.");
  }
    
  // Dahil sa pagbabago sa getSheetData, kukunin na nito ang tamang data range
  const allContracts = getSheetData(CONTRACTS_SHEET_NAME);
  Logger.log(`[getContracts] Total data rows available for filtering: ${allContracts.length}`); 

  // Helper function para mahanap ang case-insensitive key
  const findKey = (c, search) => {
      const keys = Object.keys(c);
      return keys.find(key => (key || '').trim().toLowerCase() === search.toLowerCase());
  };
    
  const filteredContracts = allContracts.filter((c, index) => {
    // Ang rowNumber dito ay ang actual row number sa Sheet (base 1, Row 5 + data row index)
    const rowNumber = MASTER_HEADER_ROW + index + 1; 

    const statusKey = findKey(c, 'Status of SFC');
    const contractIdKey = findKey(c, 'CONTRACT GRP ID');
    
    // 1. Filter: Tiyakin na may tamang Headers (Dapat okay na ito dahil inayos na ang getSheetData)
    if (!statusKey) {
 
        Logger.log(`Row ${rowNumber} skipped: 'Status of SFC' key not found after reading headers from Row ${MASTER_HEADER_ROW}. Available keys: ${Object.keys(c).join(', ')}`);
        return false;
    }
    if (!contractIdKey) {
        Logger.log(`Row ${rowNumber} skipped: 'CONTRACT GRP ID' key not found after reading headers from Row ${MASTER_HEADER_ROW}.`);
        return false;
    }

    // 2. Filter: Tiyakin na may Contract ID value
    const contractIdValue = (c[contractIdKey] || 
'').toString().trim();
    if (!contractIdValue) {
        Logger.log(`Row ${rowNumber} skipped: CONTRACT GRP ID value is empty.`);
        return false; 
    }
    
    // 3. Filter: Tiyakin na ang Status ay 'live' o 'on process - live'
    const status = (c[statusKey] || '').toString().trim().toLowerCase();
    const isLive = status === 'live' || status === 'on process - live';
    if (!isLive) {
        Logger.log(`Row ${rowNumber} skipped: Status is '${status}'. Not 'live' or 'on process - live'.`);
    } else {
        Logger.log(`Row ${rowNumber} INCLUDED: ID is '${contractIdValue}', Status is '${status}'.`);
    }

    return isLive;
  });

  Logger.log(`[getContracts] Total contracts LIVE and filtered: ${filteredContracts.length}`);
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
     
      sfcRef: sfcRefKey ? (c[sfcRefKey] || '').toString() : '',              
    };
  });
}

/**
 * Kinukuha ang Employee List at Attendance Plan para sa isang Contract ID.
 * @param {string} contractId Ang Contract Group ID.
 */
function getAttendancePlan(contractId) {
    if (!contractId) throw new Error("Contract ID is required.");
    // TANDAAN: HINDI na tinatawag ang ensureContractSheets() dito. Dapat itong tawagin bago tumawag sa getAttendancePlan.
    
    // 1. Kumuha ng Employee Data
    const empSheetName = getDynamicSheetName(contractId, 'employees');
    const empData = getSheetData(empSheetName);
    // 2. Kumuha ng Attendance Plan Data
    const planSheetName = getDynamicSheetName(contractId, 'plan');
    const planData = getSheetData(planSheetName); // Reads starting from Row 5 (headers)
    const planMap = {};
    planData.forEach(row => {
        const id = row['Personnel ID'] || '';
        const dayKey = row['DayKey'] || '';
        const shift = row['Shift'] || '1stHalf'; // Default shift
        const key = `${id}_${dayKey}_${shift}`;
        planMap[key] = row['Status'] || '';
    });
    // 3. I-organisa ang Employee Data (Kasama ang 'No.')
    const employees = empData.map((e, index) => ({
        no: index + 1, // Auto-incremented No.
        // FIXED: Ginagamit ang String() upang siguraduhin na ang numeric value ay nagiging string bago tawagin ang .trim()
        id: String(e['Personnel ID'] || '').trim(),
        name: String(e['Personnel Name'] || '').trim(),
        position: String(e['Position'] || '').trim(),
        area: String(e['Area Posting'] || '').trim(),
    })).filter(e => e.id);
    // Siguraduhin na may ID

    return { employees, planMap };
}


/**
 * Ina-update ang maramihang entry sa Employee at Attendance Plan sheets, 
 * at sine-save ang Contract Info sa unang 4 rows ng Plan Sheet.
 * @param {string} contractId Ang ID ng kontrata.
 * @param {Object} contractInfo Contract details to save in the plan sheet (Payor, Agency, etc.)
 * @param {Array<Object>} employeeChanges Mga pagbabago sa Employee Info.
 * @param {Array<Object>} attendanceChanges Mga pagbabago sa Attendance Plan.
 */
function saveAllData(contractId, contractInfo, employeeChanges, attendanceChanges) {
    if (!contractId) throw new Error("Contract ID is required.");
    ensureContractSheets(contractId); // Tiyakin na may sheets na mapagsa-save-an
    
    // 1. I-save ang Contract Info sa Plan Sheet
    saveContractInfo(contractId, contractInfo);
    
    // 2. I-save ang Employee Info (Bulk)
    if (employeeChanges && employeeChanges.length > 0) {
        saveEmployeeInfoBulk(contractId, employeeChanges);
    }
    
    // 3. I-save ang Attendance Plan (Bulk)
    if (attendanceChanges && attendanceChanges.length > 0) {
        saveAttendancePlanBulk(contractId, attendanceChanges);
    }
    
    Logger.log(`[saveAllData] Successfully saved Contract Info, ${employeeChanges.length} employee updates and ${attendanceChanges.length} attendance updates for ${contractId}.`);
}


/**
 * Sine-save ang Contract Details sa Row 1-4 ng Attendance Plan sheet.
 * @param {string} contractId 
 * @param {Object} info 
 */
function saveContractInfo(contractId, info) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const planSheetName = getDynamicSheetName(contractId, 'plan');
    const planSheet = ss.getSheetByName(planSheetName);
    
    if (!planSheet) throw new Error(`Plan Sheet for ID ${contractId} not found.`);

    const data = [
        ['PAYOR COMPANY', info.payor],           // Row 1
        ['AGENCY', info.agency],                 // Row 2
        ['SERVICE TYPE', info.serviceType],      // Row 3
        ['TOTAL HEAD COUNT', info.headCount]     // Row 4
    ];

    planSheet.getRange('A1:B4').setValues(data);
    Logger.log(`[saveContractInfo] Saved metadata for ${contractId}.`);
}


/**
 * Ina-update ang maramihang entry sa Attendance Plan sheet.
 * @param {string} contractId Ang ID ng kontrata.
 * @param {Array<Object>} changes Array of {personnelId, dayKey, shift, status}
 */
function saveAttendancePlanBulk(contractId, changes) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const planSheetName = getDynamicSheetName(contractId, 'plan');
    const planSheet = ss.getSheetByName(planSheetName);

    if (!planSheet) throw new Error(`AttendancePlan Sheet for ID ${contractId} not found.`);
    
    const HEADER_ROW = 5; // Attendance Plan headers are now at Row 5
    
    planSheet.setFrozenRows(0); 
    // Magbasa ng data simula sa Row 5. Gumamit ng .getRange() na may tamang start row.
    const lastRow = planSheet.getLastRow();
    
    // --- FIX APPLIED HERE (Line 371) ---
    const numRowsToRead = lastRow - HEADER_ROW + 1; // Removed space in const name
    
    const numColumns = planSheet.getLastColumn();
    
    let values = [];
    let headers = [];

    if (numRowsToRead > 0 && numColumns > 0) {
         // Range: Row 5, Col 1, hanggang sa dulo
         values = planSheet.getRange(HEADER_ROW, 1, numRowsToRead, numColumns).getValues();
         headers = values[0]; 
    } else {
        // Sheet is empty or only has the header row (Row 5)
        headers = ['Personnel ID', 'DayKey', 'Shift', 'Status']; // Fallback
        values.push(headers);
    }

    planSheet.setFrozenRows(HEADER_ROW); 
    
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const dayKeyIndex = headers.indexOf('DayKey');
    const shiftIndex = headers.indexOf('Shift');
    const statusIndex = headers.indexOf('Status');
    
    if (statusIndex === -1 || personnelIdIndex === -1 || dayKeyIndex === -1 || shiftIndex === -1) {
        throw new Error("Missing critical column in AttendancePlan sheet (Personnel ID, DayKey, Shift, or Status).");
    }
    
    const rowsToUpdate = {}; 
    const newRows = [];
    
    changes.forEach(data => {
        let rowFound = false;
        // Mag-loop sa data rows, simula sa index 1 (Row 6 sa sheet)
        for (let i = 1; i < values.length; i++) {
            const row = values[i];
            // TANDAAN: Sheet Row Number = i + HEADER_ROW
            if (
                (String(row[personnelIdIndex] || '').trim() === String(data.personnelId || '').trim()) && 
                (String(row[shiftIndex] || '').trim() === String(data.shift || '').trim()) &&       
                (String(row[dayKeyIndex] || '').trim() === String(data.dayKey || '').trim())         
            ) {
                // Row found, record update
                rowsToUpdate[i + HEADER_ROW] = data.status; // i + 5 is the actual sheet row number
                rowFound = true;
                break;
            }
        }
        
        // Kung walang existing row, i-record ang new row
        if (!rowFound) {
            const newRow = [];
            newRow[personnelIdIndex] = data.personnelId;
            newRow[dayKeyIndex] = data.dayKey;
            newRow[shiftIndex] = data.shift;
            newRow[statusIndex] = data.status;
            
            // I-fill ang gaps para sa appendRow
            const finalRow = [];
            for(let i = 0; i < headers.length; i++) {
                finalRow.push(newRow[i] !== undefined ? newRow[i] : '');
            }
            newRows.push(finalRow);
        }
    });

    // 1. Update existing rows
    Object.keys(rowsToUpdate).forEach(sheetRowNumber => {
        // Sheet Row Number ay tama na
        planSheet.getRange(parseInt(sheetRowNumber), statusIndex + 1).setValue(rowsToUpdate[sheetRowNumber]);
    });
    
    // 2. Append new rows
    if (newRows.length > 0) {
        planSheet.getRange(planSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
}


/**
 * Ina-update ang maramihang entry sa Employee sheet.
 * @param {string} contractId Ang ID ng kontrata.
 * @param {Array<Object>} changes Array of {id (new ID), name, position, area, isNew, oldPersonnelId}
 */
function saveEmployeeInfoBulk(contractId, changes) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const empSheetName = getDynamicSheetName(contractId, 'employees');
    const empSheet = ss.getSheetByName(empSheetName);

    if (!empSheet) throw new Error(`Employee Sheet for ID ${contractId} not found.`);
    
    empSheet.setFrozenRows(0); // Temporarily unfreeze
    const range = empSheet.getDataRange();
    const values = range.getValues();
    const headers = values[0]; 
    empSheet.setFrozenRows(1); // Restore freeze

    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('Position');
    const areaIndex = headers.indexOf('Area Posting');
    
    if (personnelIdIndex === -1 || nameIndex === -1 || positionIndex === -1 || areaIndex === -1) {
        throw new Error("Missing critical column in Employee sheet.");
    }
    
    const rowsToUpdate = {}; // Key: i+1 (sheet row number), Value: [newId, newName, newPosition, newArea]
    const newRows = [];

    // Gumawa ng map para sa mabilis na lookup ng existing rows
    const personnelIdMap = {};
    for (let i = 1; i < values.length; i++) {
        personnelIdMap[String(values[i][personnelIdIndex] || '').trim()] = i + 1; // Store row number
    }
    
    changes.forEach(data => {
        const oldId = String(data.oldPersonnelId || '').trim();
        const newId = String(data.id || '').trim();
        
        // 1. Existing Row Update (Mahanap gamit ang OLD ID)
        if (personnelIdMap[oldId]) {
            const sheetRowNumber = personnelIdMap[oldId];
            
            // Check if a new ID is already being used by another row (Simple validation)
            if(oldId !== newId && personnelIdMap[newId] && personnelIdMap[newId] !== sheetRowNumber) {
                 Logger.log(`SKIPPED: Cannot change ID from ${oldId} to ${newId}. ${newId} already exists in Row ${personnelIdMap[newId]}.`);
                 return; 
            }
            
            // I-record ang update
            rowsToUpdate[sheetRowNumber] = [newId, data.name, data.position, data.area];

            // Kung nagbago ang ID, i-update ang personnelIdMap para maiwasan ang conflict sa ibang changes sa batch na ito
            if (oldId !== newId) {
                delete personnelIdMap[oldId];
                personnelIdMap[newId] = sheetRowNumber;
            }

        } 
        // 2. New Row Append
        else {
            // Check if the new ID already exists in the existing sheet data or in the batch of updates
             if (personnelIdMap[newId]) {
                Logger.log(`SKIPPED: Cannot add new ID ${newId}. It already exists in Row ${personnelIdMap[newId]} or in an earlier update batch.`);
                return;
             }
            
            const newRow = [];
            newRow[personnelIdIndex] = newId;
            newRow[nameIndex] = data.name;
            newRow[positionIndex] = data.position;
            newRow[areaIndex] = data.area;
            
            const finalRow = [];
            for(let i = 0; i < headers.length; i++) {
                finalRow.push(newRow[i] !== undefined ? newRow[i] : '');
            }
            newRows.push(finalRow);

            // Add to map temporarily to prevent duplicates in the same batch
            personnelIdMap[newId] = -1; // Flag as newly added in this batch
        }
    });

    // 1. Update existing rows (using one call per column for efficiency)
    Object.keys(rowsToUpdate).forEach(sheetRowNumber => {
        const rowData = rowsToUpdate[sheetRowNumber];
        empSheet.getRange(parseInt(sheetRowNumber), personnelIdIndex + 1, 1, 4).setValues([
            [rowData[0], rowData[1], rowData[2], rowData[3]]
        ]);
    });
    
    // 2. Append new rows
    if (newRows.length > 0) {
        empSheet.getRange(empSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
}
