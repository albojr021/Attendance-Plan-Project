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
    const safeId = (contractId || '').replace(/[\\/?*[]/g, '_'); // Linisin ang ID
    if (type === 'employees') {
        return `${safeId} - Employees`;
    }
    return `${safeId} - AttendancePlan`;
}

/**
 * Tinitiyak na ang Employee at Attendance Plan sheets para sa Contract ID ay existing at may tamang headers.
 * @param {string} contractId Ang ID ng kontrata.
 */
function ensureContractSheets(contractId) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // --- EMPLOYEES SHEET ---
    const empSheetName = getDynamicSheetName(contractId, 'employees');
    let empSheet = ss.getSheetByName(empSheetName);
    
    // Tandaan: Ang mga auto-generated sheets ay nagsisimula sa Row 1.
    if (!empSheet) {
        empSheet = ss.insertSheet(empSheetName);
        empSheet.clear();
        // HEADERS para sa Employees
        const empHeaders = ['Personnel ID', 'Personnel Name', 'Position', 'Area Posting'];
        empSheet.getRange(1, 1, 1, empHeaders.length).setValues([empHeaders]);
        empSheet.setFrozenRows(1);
    }
    
    // --- ATTENDANCE PLAN SHEET ---
    const planSheetName = getDynamicSheetName(contractId, 'plan');
    let planSheet = ss.getSheetByName(planSheetName);

    if (!planSheet) {
        planSheet = ss.insertSheet(planSheetName);
        planSheet.clear();
        // HEADERS para sa Attendance Plan data
        const planHeaders = ['Personnel ID', 'DayKey', 'Shift', 'Status'];
        planSheet.getRange(1, 1, 1, planHeaders.length).setValues([planHeaders]);
        planSheet.setFrozenRows(1);
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
    const contractIdValue = (c[contractIdKey] || '').toString().trim();
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
    
    // Tiyakin na ang Sheet ay existing
    ensureContractSheets(contractId);
    
    // 1. Kumuha ng Employee Data
    const empSheetName = getDynamicSheetName(contractId, 'employees');
    const empData = getSheetData(empSheetName);

    // 2. Kumuha ng Attendance Plan Data
    const planSheetName = getDynamicSheetName(contractId, 'plan');
    const planData = getSheetData(planSheetName);
    
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
        id: (e['Personnel ID'] || '').trim(),
        name: (e['Personnel Name'] || '').trim(),
        position: (e['Position'] || '').trim(),
        area: (e['Area Posting'] || '').trim(),
    })).filter(e => e.id); // Siguraduhin na may ID

    return { employees, planMap };
}

/**
 * Ina-update ang isang entry sa Attendance Plan sheet.
 * @param {Object} data Ang data na isa-save. Expected: {contractId, personnelId, shift, dayKey, status}
 */
function saveAttendancePlan(data) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const planSheetName = getDynamicSheetName(data.contractId, 'plan');
    const planSheet = ss.getSheetByName(planSheetName);

    if (!planSheet) throw new Error(`AttendancePlan Sheet for ID ${data.contractId} not found.`);

    // Ang mga auto-generated sheets ay nagsisimula sa Row 1, kaya ang headers ay values[0]
    const range = planSheet.getDataRange();
    const values = range.getValues();
    const headers = values[0]; 

    const personnelIdIndex = headers.indexOf('Personnel ID');
    const shiftIndex = headers.indexOf('Shift');
    const dayKeyIndex = headers.indexOf('DayKey');
    const statusIndex = headers.indexOf('Status');
    
    if (statusIndex === -1) throw new Error("Missing 'Status' column in AttendancePlan sheet.");
    if (personnelIdIndex === -1) throw new Error("Missing 'Personnel ID' column in Employee sheet.");

    // Hanapin ang existing row
    let rowFound = false;
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (
            (row[personnelIdIndex] || '').trim() === (data.personnelId || '').trim() && 
            (row[shiftIndex] || '').trim() === (data.shift || '').trim() &&       
            (row[dayKeyIndex] || '').trim() === (data.dayKey || '').trim()         
        ) {
            // Row found, update Status
            planSheet.getRange(i + 1, statusIndex + 1).setValue(data.status);
            rowFound = true;
            break;
        }
    }

    // Kung walang existing row, mag-add ng bago
    if (!rowFound) {
        planSheet.appendRow([
            data.personnelId, 
            data.dayKey, 
            data.shift, 
            data.status
        ]);
    }
}

/**
 * Ina-update ang Employee Info (Name, Position, Area)
 * @param {Object} data Ang data na isa-save. Expected: {contractId, personnelId, name, position, area}
 */
function saveEmployeeInfo(data) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const empSheetName = getDynamicSheetName(data.contractId, 'employees');
    const empSheet = ss.getSheetByName(empSheetName);

    if (!empSheet) throw new Error(`Employee Sheet for ID ${data.contractId} not found.`);

    // Ang mga auto-generated sheets ay nagsisimula sa Row 1, kaya ang headers ay values[0]
    const range = empSheet.getDataRange();
    const values = range.getValues();
    const headers = values[0]; 

    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('Position');
    const areaIndex = headers.indexOf('Area Posting');
    
    if (personnelIdIndex === -1) throw new Error("Missing 'Personnel ID' column in Employee sheet.");

    // Hanapin ang existing row
    let rowFound = false;
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if ((row[personnelIdIndex] || '').trim() === (data.personnelId || '').trim()) {
            // Row found, update fields
            empSheet.getRange(i + 1, nameIndex + 1).setValue(data.name);
            empSheet.getRange(i + 1, positionIndex + 1).setValue(data.position);
            empSheet.getRange(i + 1, areaIndex + 1).setValue(data.area);
            rowFound = true;
            break;
        }
    }

    // Kung walang existing row, mag-add ng bago
    if (!rowFound) {
      // Tiyakin na may Personnel ID at Name bago i-save
      if (!data.personnelId || !data.name) {
          throw new Error("Personnel ID and Name are required to add a new employee.");
      }
      
      const newRow = [];
      newRow[personnelIdIndex] = data.personnelId;
      newRow[nameIndex] = data.name;
      newRow[positionIndex] = data.position || '';
      newRow[areaIndex] = data.area || '';
      
      // I-fill ang gaps para sa appendRow
      const finalRow = [];
      for(let i = 0; i < headers.length; i++) {
          finalRow.push(newRow[i] !== undefined ? newRow[i] : '');
      }
      
      empSheet.appendRow(finalRow);
    }
}
