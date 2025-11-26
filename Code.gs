const SPREADSHEET_ID = '1rQnJGqcWcEBjoyAccjYYMOQj7EkIu1ykXTMLGFzzn2I';
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; 
const CONTRACTS_SHEET_NAME = 'MASTER';

// *** NEW CONSTANTS FOR CENTRALIZED 201 FILE ***
const FILE_201_ID = '19eJ-qC68eazrVMmjPzZjTvfns_N9h03Ha6SDGTvuv0E';
const FILE_201_SHEET_NAME = 'Basic Information'; 
// ----------------------------------------------

// *** NEW CONSTANTS FOR CONSOLIDATED PLAN SHEET ***
const PLAN_SHEET_NAME = 'AttendancePlan_Consolidated';
const PLAN_HEADER_ROW = 1;

// Header is at Row 1
// CONTRACT # to AREA POSTING = 17 Columns (0-based index 0 to 16)
const PLAN_FIXED_COLUMNS = 17;

// *** NEW CONSTANT FOR HALF-MONTH PLANNING ***
const PLAN_MAX_DAYS_IN_HALF = 16; 
// ------------------------------------------------

const MASTER_HEADER_ROW = 5;
const SIGNATORY_MASTER_SHEET = 'SignatoryMaster';

// --- BAGONG CONSTANT PARA SA CONSOLIDATED EMPLOYEE MASTER ---
const EMPLOYEE_MASTER_SHEET_NAME = 'EmployeeMaster_Consolidated'; 
// ------------------------------------------------------------
const ADMIN_EMAILS = ['mcdmarketingstorage@megaworld-lifestyle.com'];

const LOG_SHEET_NAME = 'PrintLog';
const LOG_HEADERS = [
    'Reference #', 
    'SFC Ref#', 
    'Plan Sheet Name (N/A)', // Placeholder
    'Plan Period Display', 
    'Payor Company', 
    'Agency',
    'Sub Property',         
    'Service Type',
    'User Email', 
    'Timestamp',
    'Locked Personnel IDs'  
];

const AUDIT_LOG_SHEET_NAME = 'ScheduleAuditLog';
const AUDIT_LOG_HEADERS = [
    'Timestamp', 
    'User Email', 
    'SFC Ref#', 
    'Personnel ID', 
    'Personnel Name', 
    'Plan Sheet Name (N/A)', // Placeholder
    'Date (YYYY-M-D)', 
    'Shift', 
    'Reference #', 
    'Old Status', 
    'New Status'
];

// --- NEW CONSTANTS FOR UNLOCK REQUEST LOGGING ---
const UNLOCK_LOG_SHEET_NAME = 'UnlockRequestLog';

const UNLOCK_LOG_HEADERS = [
    'SFC Ref#',
    'Personnel ID',
    'Personnel Name',
    'Locked Ref #', // Reference number of the locked print log
    'Requesting User',
    'Request Timestamp',
    'Admin Email',
    'Admin Action Timestamp',
    'Status (APPROVED/REJECTED)',
    'User Action Type', // Hal.'Edit Personal AP, Edit Entire AP, Create AP Plan For an OLD AP',
    'User Action Timestamp'
];
// ---------------------------------------

function doGet(e) {
  if (e.parameter.action) {
    return processAdminUnlockFromUrl(e.parameter);
  }
  
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Attendance Plan Monitor');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function sanitizeHeader(header) {
    if (!header) return '';
    return String(header).replace(/[^A-Za-z00-9#\/]/g, '');
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
  // No longer checking for Plan Sheet name prefix, as it's a single sheet
  else if (sheetName === PLAN_SHEET_NAME || sheetName === EMPLOYEE_MASTER_SHEET_NAME) {
      startRow = 1;
      // Assuming headers are at Row 1 for both consolidated sheets
      if (numRows < startRow) {
          numRows = 0;
      } else {
          numRows = sheet.getLastRow() - startRow + 1;
      }
  }


  if (numRows <= 0 || numColumns === 0) return [];
  const range = sheet.getRange(startRow, 1, numRows, numColumns);
  const values = range.getDisplayValues();
  const headers = values[0];
  const cleanHeaders = headers.map(header => (header || '').toString().trim());

  const data = [];
  for (let i = 1; i < values.length; i++) { 
    const row = values[i];
    if (sheetName === CONTRACTS_SHEET_NAME || sheetName === EMPLOYEE_MASTER_SHEET_NAME) {
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
  
  Logger.log(`[getSheetData] Total data rows processed (excluding header): ${data.length} from ${sheetName}`);
  return data;
}

function checkContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef || year === undefined || month === undefined || !shift) return false;
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        // --- BINAGO: Check na lang ang Consolidated Plan at Employee Master Sheets ---
        const empSheet = ss.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
        const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
        return !!empSheet && !!planSheet;
        // -------------------------------------------------------------------------
    } catch (e) {
         Logger.log(`[checkContractSheets] ERROR: Failed to open Spreadsheet ID ${TARGET_SPREADSHEET_ID}. Error: ${e.message}`);
        return false;
    }
}

function getOrCreateConsolidatedEmployeeMasterSheet(ss) {
    let empSheet = ss.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
    if (!empSheet) {
        empSheet = ss.insertSheet(EMPLOYEE_MASTER_SHEET_NAME);
        empSheet.clear();
        // --- BAGONG HEADERS KASAMA ANG CONTRACT # ---
        const empHeaders = ['CONTRACT #', 'Personnel ID', 'Personnel Name', 'Position', 'Area Posting'];
        empSheet.getRange(1, 1, 1, empHeaders.length).setValues([empHeaders]);
        empSheet.setFrozenRows(1);
        Logger.log(`[getOrCreateConsolidatedEmployeeMasterSheet] Created Consolidated Employee sheet: ${EMPLOYEE_MASTER_SHEET_NAME}`);
    } 
    return empSheet;
}

function createContractSheets(sfcRef, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    // --- EMPLOYEES MASTER SHEET (Consolidated) ---
    getOrCreateConsolidatedEmployeeMasterSheet(ss);
    // --- CONSOLIDATED ATTENDANCE PLAN SHEET ---
    let planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    const getConsolidatedPlanHeaders = () => {
        // Total 33 Columns (17 fixed info + 16 days) 
        const base = [
            'CONTRACT #', 'TOTAL HEADCOUNT', 'PROP OR GRP CODE', 'SERVICE TYPE', 
            'SECTOR', 'PAYOR COMPANY', 'AGENCY', 'MONTH', 'YEAR', 
            'PERIOD / SHIFT', 'GROUP', 'PRINT VERSION', 'Reference #',
       
        
        'Personnel ID', 'Personnel Name', 'POSITION', 'AREA POSTING'
        ];
        // UPDATED: Loop only up to PLAN_MAX_DAYS_IN_HALF (16)
        for (let d = 1; d <= PLAN_MAX_DAYS_IN_HALF; d++) {
            base.push(`DAY${d}`);
        }
        return base;
    };
    if (!planSheet) {
        planSheet = ss.insertSheet(PLAN_SHEET_NAME);
        planSheet.clear();
        const planHeaders = getConsolidatedPlanHeaders();
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, planHeaders.length).setValues([planHeaders]);
        planSheet.setFrozenRows(PLAN_HEADER_ROW); 
        
        // Set number format for fixed info columns to Plain Text ('@')
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, PLAN_FIXED_COLUMNS).setNumberFormat('@');
        Logger.log(`[createContractSheets] Created Consolidated Attendance Plan sheet: ${PLAN_SHEET_NAME} with headers at Row ${PLAN_HEADER_ROW}.`);
    } 
}

function ensureContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required to ensure sheets.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    
    // BINAGO START: Check/create lang ang Consolidated Sheets
    const empSheet = getOrCreateConsolidatedEmployeeMasterSheet(ss);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    
    if (!planSheet) {
        // Since getOrCreateConsolidatedEmployeeMasterSheet is called, we only need to check the plan sheet
        createContractSheets(sfcRef, year, month, shift);
        Logger.log(`[ensureContractSheets] Ensured Consolidated Plan sheet existence.`);
    }
    // BINAGO END
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
    // NEW: Get the keys for the new metadata fields
   
    
    const 
    propOrGrpCodeKey = findKey(c, 'PROP OR GRP CODE'); // Assuming the header is exactly this (Col D)
    const sectorKey = findKey(c, 'SECTOR'); // Assuming the header is exactly this (Col O)
      
    return {
      id: contractIdKey ? (c[contractIdKey] || '').toString() : '',     
      status: statusKey ? (c[statusKey] || '').toString() : '',   
      payorCompany: payorKey ? (c[payorKey] 
|| '').toString() : '', 
      
   
      agency: agencyKey ?
(c[agencyKey] ||
 '').toString() : '',       
      serviceType: serviceTypeKey ?
(c[serviceTypeKey] || '').toString() : '',   
      headCount: parseInt(headCountKey ? c[headCountKey] : 0) ||
0, 
      sfcRef: sfcRefKey ?
(c[sfcRefKey] || '').toString() : '', 
      // NEW FIELDS
      propOrGrpCode: propOrGrpCodeKey ?
(c[propOrGrpCodeKey] || '').toString() : '',
      sector: sectorKey ?
(c[sectorKey] || '').toString() : '',
    };
  });
}

function cleanPersonnelId(rawId) {
    let idString = String(rawId || '').trim();
    return idString.replace(/\D/g, '');
}

/**
 * Reads Personnel ID (CODE), First Name, and Last Name from the 201 Master File.
 * Formats the name as "LastName, FirstName" and cleans the ID.
 */
function get201FileMasterData() {
    if (FILE_201_ID === 'PUNAN_MO_ITO_NG_201_SPREADSHEET_ID') {
        Logger.log('[get201FileMasterData] ERROR: FILE_201_ID is not set.');
        return [];
    }

    try {
        const ss = SpreadsheetApp.openById(FILE_201_ID);
        const sheet = ss.getSheetByName(FILE_201_SHEET_NAME);
        
        if (!sheet) {
            Logger.log(`[get201FileMasterData] ERROR: Sheet ${FILE_201_SHEET_NAME} not found.`);
            return [];
        }

        // Start reading from Row 6 (Index 5)
        const START_ROW = 6; 
        const NUM_ROWS = sheet.getLastRow() - START_ROW + 1;
        const NUM_COLS = 11; // Up to Column K (Last Name)

        if (NUM_ROWS <= 0) return [];
        
        // Read range from Col D (CODE) to Col K (LAST NAME) which is 8 columns wide (D to K)
        // D(4) E(5) F(6) G(7) H(8) I(9) J(10) K(11) -> 8 columns
        const values = sheet.getRange(START_ROW, 4, NUM_ROWS, 8).getDisplayValues(); 

        const masterData = values.map(row => {
            // Indices based on the read range (D is index 0, K is index 7)
            const personnelIdRaw = row[0]; // Col D (CODE)
            const firstName = String(row[4] || '').trim(); // Col H (FIRST NAME) -> Index 4
            const lastName = String(row[7] || '').trim();  // Col K (LAST NAME) -> Index 7
            
            const cleanId = cleanPersonnelId(personnelIdRaw); // Assuming cleanPersonnelId exists and works
            const formattedName = lastName ? `${lastName}, ${firstName}` : firstName;
            
            if (!cleanId || !formattedName) return null; // Skip if ID or Name is blank

            return {
                id: cleanId,
                name: formattedName.toUpperCase(),
                // NOTE: Walang Position/Area dito. Ikakabit ito sa front-end gamit ang SFC Master list.
            };
        }).filter(item => item !== null);

        Logger.log(`[get201FileMasterData] Retrieved ${masterData.length} records from 201 file.`);
        return masterData;

    } catch (e) {
        Logger.log(`[get201FileMasterData] ERROR: ${e.message}`);
        throw new Error(`Failed to access 201 Master File. Please check ID and sheet name in Code.gs. Error: ${e.message}`);
    }
}

function getSignatoryMasterData() {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = getOrCreateSignatoryMasterSheet(ss);

        if (sheet.getLastRow() < 2) return [];
        // Basahin ang 2 columns (Name at Designation)
        const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2);
        const values = range.getDisplayValues();

        // Ibalik ang listahan ng objects na may name at designation
        return values.map(row => ({
            name: String(row[0] || '').trim(), 
            designation: String(row[1] || '').trim() 
        })).filter(item => item.name);
        // Siguraduhin na may pangalan
    } catch (e) {
        Logger.log(`[getSignatoryMasterData] ERROR: ${e.message}`);
        return [];
    }
}

// --- BINAGO: Aalisin na ang dynamic sheet name para sa employees ---
function getDynamicSheetName(sfcRef, type, year, month, shift) {
    const safeRef = (sfcRef || '').replace(/[\\/?*[]/g, '_');
    if (type === 'employees') {
        // Ibalik na lang ang consolidated name
        return EMPLOYEE_MASTER_SHEET_NAME;
    }
    return `${safeRef} - AttendancePlan`; 
}

function getEmployeeMasterData(sfcRef) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");

    // --- NEW LOGIC: 1. Get ALL records from the centralized 201 File ---
    const all201Data = get201FileMasterData();
    const clean201DataMap = {}; // Key: ID
    all201Data.forEach(e => {
        // Ang 201 data ay walang Position/Area, kaya ito ang initial details.
        clean201DataMap[e.id] = { 
            id: e.id, 
            name: e.name, 
            position: '', 
            area: '' 
        };
    });
    // --- END NEW LOGIC ---
    
    // 2. Basahin ang Consolidated sheet (SFC-specific Position/Area)
    const allMasterData = getSheetData(TARGET_SPREADSHEET_ID, EMPLOYEE_MASTER_SHEET_NAME);
    
    const filteredMasterData = allMasterData.filter(e => {
        const contractRef = String(e['CONTRACT #'] || '').trim();
        return contractRef === sfcRef;
    });
    
    // 3. I-Merge ang data: Gamitin ang Position/Area mula sa SFC Master (EmployeeMaster_Consolidated) 
    // at idagdag ang mga record mula sa 201 na wala pa sa SFC master.

    const finalEmployeeList = filteredMasterData.map(e => {
        const id = cleanPersonnelId(e['Personnel ID']);
        const name = String(e['Personnel Name'] || '').trim();
        const position = String(e['Position'] || '').trim();
        const area = String(e['Area Posting'] || '').trim();
        
        // Tanggalin ang ID na ito sa 201 map dahil na-process na ito gamit ang SFC data
        if (clean201DataMap[id]) {
            delete clean201DataMap[id];
        }

        return { id, name, position, area };
    }).filter(e => e.id); // Filter out records na walang ID
    
    // 4. Idagdag ang natitirang records mula sa 201 (ito ang mga record na walang Position/Area sa SFC master)
    Object.values(clean201DataMap).forEach(emp => {
        // Tiyakin lang na walang duplicate na ID (kahit na dapat walang ganyan dahil sa filter)
        if (!finalEmployeeList.some(e => e.id === emp.id)) {
             finalEmployeeList.push(emp);
        }
    });

    Logger.log(`[getEmployeeMasterData] Final Employee List Count: ${finalEmployeeList.length} (Merged 201 + SFC Master)`);
    
    return finalEmployeeList.map((e, index) => ({
        id: e.id,
        name: e.name,
        position: e.position,
        area: e.area,
    })).filter(e => e.id);
}

function getEmployeeNameFromMaster(sfcRef, personnelId) {
    if (!sfcRef || !personnelId) return 'N/A';
    const masterData = getEmployeeMasterData(sfcRef);
    const cleanId = cleanPersonnelId(personnelId);
    const employee = masterData.find(e => e.id === cleanId);
    return employee ?
employee.name : 'N/A';
}

function getSortedPlanSheets(sfcRef, ss) {
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) return [];
    return [{ name: PLAN_SHEET_NAME, date: new Date(2000, 0, 1) }];
}


function getEmployeeSchedulePattern(sfcRef, personnelId) {
    if (!sfcRef || !personnelId) return {};
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const cleanId = cleanPersonnelId(personnelId);
    
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet || planSheet.getLastRow() < PLAN_HEADER_ROW) return {};
    
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - PLAN_HEADER_ROW;
    const numColumns = planSheet.getLastColumn();
    if (numRowsToRead <= 0 || numColumns < (PLAN_FIXED_COLUMNS + 3)) return {};
    const planValues = planSheet.getRange(PLAN_HEADER_ROW, 1, lastRow - PLAN_HEADER_ROW + 1, numColumns).getDisplayValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const shiftIndex = headers.indexOf('PERIOD / SHIFT');
    const sfcRefIndex = headers.indexOf('CONTRACT #');
    const day1Index = headers.indexOf('DAY1');
    const printVersionIndex = headers.indexOf('PRINT VERSION');
    
    if (personnelIdIndex === -1 || shiftIndex === -1 || sfcRefIndex === -1 || day1Index === -1) return {};
    let targetRow = null;
    let latestVersion = 0;
    let latestDate = new Date(0);
    dataRows.forEach(row => {
        const currentId = cleanPersonnelId(row[personnelIdIndex]);
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        
        if (currentId === cleanId && currentSfc === sfcRef) {
            
            const printVersionString = String(row[printVersionIndex] || '').trim();
            const versionParts = printVersionString.split('-');
      
            
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
            
            const monthShort = String(row[headers.indexOf('MONTH')] || '').trim();
            const yearNum = parseInt(row[headers.indexOf('YEAR')] || '0', 10);
            
          
            let planDate = new Date(0);
      
            
            if (monthShort && yearNum) {
                 planDate = new Date(`${monthShort} 1, ${yearNum}`);
            }
            
    
            
            if (version 
> latestVersion || (version === latestVersion && planDate.getTime() > latestDate.getTime())) {
        
      
                latestVersion = version;
                latestDate = planDate;
                targetRow = row;
            }
        }
    });
    if (!targetRow) return {};
    const dayPatternMap = {};
    const targetMonthShort = String(targetRow[headers.indexOf('MONTH')] || '').trim();
    const targetYear = parseInt(targetRow[headers.indexOf('YEAR')] || '0', 10);
    if (!targetMonthShort || !targetYear) return {};
    
    const targetMonth = new Date(`${targetMonthShort} 1, ${targetYear}`).getMonth();
    // NEW LOGIC: Use the shift from the found latest row to correctly map the days
    const targetShift = String(targetRow[shiftIndex] || '').trim();
    const loopLimit = PLAN_MAX_DAYS_IN_HALF; // 16
    const startDayOfMonth = targetShift === '1stHalf' ? 1 : 16;
    const endDayOfMonth = new Date(targetYear, targetMonth + 1, 0).getDate();

    for (let d = 1; d <= loopLimit; d++) { // Loop only up to DAY16 column
        const dayHeader = `DAY${d}`;
        const dayColIndex = headers.indexOf(dayHeader);
        if (dayColIndex === -1) continue; 
        
        const status = String(targetRow[dayColIndex] || '').trim();
        // Calculate actual day of the month
        const actualDay = startDayOfMonth + d - 1;
        if (actualDay > endDayOfMonth) continue; // Skip days outside the month (e.g., Nov 31)

        const currentDate = new Date(targetYear, targetMonth, actualDay);
        if (currentDate.getMonth() !== targetMonth) continue; 

        const dayOfWeek = currentDate.getDay(); 
        const dayKey = dayOfWeek.toString();
        if (status && status !== 'NA') {
             dayPatternMap[dayKey] = status;
        }
    }

    Logger.log(`[getEmployeeSchedulePattern] Final Pattern for ID ${cleanId}: ${JSON.stringify(dayPatternMap)}`);
    return dayPatternMap;
}

/**
 * FIX: This function is updated to filter the PrintLog by sfcRef, year, month, and shift
 * to ensure that only IDs locked for the CURRENT PERIOD are considered locked in the UI.
 */
function getLockedPersonnelIds(ss, sfcRef, year, month, shift) {
    const logSheet = getOrCreateLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    
    if (lastRow < 2) return {};

    const REF_NUM_COL = 1;
    const SFC_REF_COL = 2; 
    const PERIOD_DISPLAY_COL = 4; 
    const LOCKED_IDS_COL = LOG_HEADERS.length;
    
    const LOG_HEADERS_COUNT = LOG_HEADERS.length; 
    
    const values = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS_COUNT).getDisplayValues(); 
    const lockedIdRefMap = {};
    
    // Calculate target period display string for filtering (e.g., 'November 1-15, 2025 (1stHalf)')
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
    
    values.forEach(row => {
        // --- NEW FILTERING LOGIC START (Filter by context) ---
        const currentSfc = String(row[SFC_REF_COL - 1] || '').trim();
        const currentPeriodDisplay = String(row[PERIOD_DISPLAY_COL - 1] || '').trim();

        // Only process rows that match the current SFC and Period/Shift loaded in the UI
        if (currentSfc !== sfcRef || currentPeriodDisplay !== dateRange) {
             return; 
        }
        // --- NEW FILTERING LOGIC END ---
        
        const refNum = String(row[REF_NUM_COL - 1] || '').trim();
        const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim(); 
        
        if (lockedIdsString) {
            const idsList = lockedIdsString.split(',').map(id => id.trim());
            
            idsList.forEach(idWithPrefix => {
                const cleanId = cleanPersonnelId(idWithPrefix);
                 
                if (cleanId.length >= 3 && !idWithPrefix.startsWith('UNLOCKED:')) { 
                    if (!lockedIdRefMap[cleanId]) { 
                         // Map the locked ID to the Reference # (Print Version String)
                         lockedIdRefMap[cleanId] = refNum;
                    }
                }
            });
        }
    });
    return lockedIdRefMap;
}


/**
 * FIX: This function is updated to filter the PrintLog by sfcRef, year, month, and shift
 * to ensure that only the most recent Print Version for the CURRENT PERIOD is used for Audit logging.
 */
function getHistoricalReferenceMap(ss, sfcRef, year, month, shift) {
    const logSheet = ss.getSheetByName('PrintLog');
    if (!logSheet || logSheet.getLastRow() < 2) return {};

    const LOG_HEADERS_COUNT = LOG_HEADERS.length; 
    const REF_NUM_COL = 1;              
    const SFC_REF_COL = 2;
    const PERIOD_DISPLAY_COL = 4;
    const LOCKED_IDS_COL = LOG_HEADERS.length;
    
    const allValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, LOG_HEADERS_COUNT).getDisplayValues();
    const historicalRefMap = {};
    
    // Calculate target period display string for filtering
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

    for (let i = allValues.length - 1; i >= 0; i--) {
        const row = allValues[i];
        
        // --- NEW FILTERING LOGIC START (Filter by context) ---
        const currentSfc = String(row[SFC_REF_COL - 1] || '').trim();
        const currentPeriodDisplay = String(row[PERIOD_DISPLAY_COL - 1] || '').trim();

        // Filter by current SFC and Period/Shift Display
        if (currentSfc !== sfcRef || currentPeriodDisplay !== dateRange) {
             continue; // Skip logs that do not match the current period/shift
        }
        // --- NEW FILTERING LOGIC END ---

        // Reference # column (now storing the print version string)
        const refNumRaw = String(row[REF_NUM_COL - 1] || '').trim();
        const refNum = refNumRaw;
        const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim();
        
        if (refNum) {
            const allIdsInString = lockedIdsString.split(',').map(s => s.trim());
            allIdsInString.forEach(idWithPrefix => {
                const cleanId = cleanPersonnelId(idWithPrefix);
                
                // Only consider unlocked (non-prefixed) IDs
                if (cleanId.length >= 3 && !idWithPrefix.startsWith('UNLOCKED:')) { 
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
    ensureContractSheets(sfcRef, year, month, shift);

    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    // FIX: Pass context parameters to getLockedPersonnelIds
    const lockedIdRefMap = getLockedPersonnelIds(ss, sfcRef, year, month, shift);
    const empMasterData = getEmployeeMasterData(sfcRef);
    const empDetailMap = {};
    empMasterData.forEach(e => {
        if (e.id) {
            empDetailMap[e.id] = { id: e.id, name: e.name, position: e.position, area: e.area };
        }
    });
    if (!planSheet) return { employees: [], planMap: {}, lockedIds: Object.keys(lockedIdRefMap), lockedIdRefMap: lockedIdRefMap };
    
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();
    if (numRowsToRead <= 0 || numColumns < (PLAN_FIXED_COLUMNS + 3)) { 
        return { employees: [], planMap: {}, lockedIds: Object.keys(lockedIdRefMap), lockedIdRefMap: lockedIdRefMap };
    }

    const planValues = planSheet.getRange(HEADER_ROW, 1, lastRow - PLAN_HEADER_ROW + 1, numColumns).getDisplayValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);
    
    const sfcRefIndex = headers.indexOf('CONTRACT #');
    const monthIndex = headers.indexOf('MONTH');
    const yearIndex = headers.indexOf('YEAR');
    const shiftIndex = headers.indexOf('PERIOD / SHIFT');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('POSITION');
    const areaIndex = headers.indexOf('AREA POSTING');
    const printVersionIndex = headers.indexOf('PRINT VERSION');
    const day1Index = headers.indexOf('DAY1');
    if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1 || day1Index === -1) {
         throw new Error("Missing critical column in Consolidated Plan sheet.");
    }
    
    const planMap = {};
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' }); 
    const targetYear = String(year);
    
    const latestVersionMap = {};
    latestVersionMap.group = '1'; // Default group if no data found

    dataRows.forEach(row => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        
   
        
        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift && id) {
            
            const printVersionString = String(row[printVersionIndex] || '').trim();
            const versionParts = printVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
   
            
        
  
            const mapKey = id; 
            const existingRow = latestVersionMap[mapKey];
            
            if (!existingRow ||
version > (parseFloat(existingRow[printVersionIndex].split('-').pop()) ||
 0)) {
                latestVersionMap[mapKey] = row;
            }
        }
    });
    const latestDataRows = Object.values(latestVersionMap).filter(r => r.length > 0); // Filter out the 'group' default property
    const employeesInPlan = new Set();
    const employeesDetails = [];
    
    // NEW LOGIC: Calculate actual days of the month based on shift
    const startDayOfMonth = shift === '1stHalf' ?
1 : 16;
    const endDayOfMonth = new Date(year, month + 1, 0).getDate(); 
    const loopLimit = PLAN_MAX_DAYS_IN_HALF;
    // 16
    
    latestDataRows.forEach((row, index) => {
        const id = cleanPersonnelId(row[personnelIdIndex]);
        
        if (id) {
            employeesInPlan.add(id); 
 
            employeesDetails.push({
                no: 0, 
               
        
    id: id, 
  
                name: String(row[nameIndex] || '').trim(),
                position: String(row[positionIndex] || '').trim(),
                area: String(row[areaIndex] || '').trim(),
            });
           
    
            
            // UPDATED: Loop only up to DAY16 column
            for (let d = 1; d <= loopLimit; d++) {
                const actualDay = startDayOfMonth + d - 1;
                
         
            
                if (actualDay > endDayOfMonth) continue; // Skip day if it exceeds max day of the month

  
                
    const dayKey = `${year}-${month + 1}-${actualDay}`; 
                const dayColIndex = day1Index + d - 1; 

     
                if (dayColIndex < numColumns) {
          
                    const status = String(row[dayColIndex] ||
 '').trim();
          
                    const key = `${id}_${dayKey}_${shift}`;
                    if (status) {
                         planMap[key] = status;
                    }
                }
  
            }
        }
    });
    const employees = employeesDetails.map((e, index) => { 
        return {
           no: index + 1, 
            id: e.id, 
            name: e.name,
            position: e.position,
            area: e.area,
        }
     }).filter(e => e.id);
    return { employees, planMap, lockedIds: Object.keys(lockedIdRefMap), lockedIdRefMap: lockedIdRefMap };
}


// **UPDATED:** Refactored saveAllData to fetch and pass the lockedIdRefMap to saveEmployeeInfoBulk
function saveAllData(sfcRef, contractInfo, employeeChanges, attendanceChanges, year, month, shift, group) { 
    Logger.log(`[saveAllData] Starting save for SFC Ref#: ${sfcRef}, Month/Shift: ${month}/${shift}, Group: ${group}`);
    if (!sfcRef) {
      throw new Error("SFC Ref# is required.");
    }
    ensureContractSheets(sfcRef, year, month, shift);
    
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    // NEW: Get the map of currently locked IDs for contextual logging (PASS CONTEXT)
    const lockedIdRefMap = getLockedPersonnelIds(ss, sfcRef, year, month, shift);
    const lockedIds = Object.keys(lockedIdRefMap);
    
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
    // NEW LOGIC: Identify which employee IDs need their plan rows deleted (i.e., they were marked for deletion in the UI)
    const deletionList = finalEmployeeChanges.filter(c => c.isDeleted).map(c => c.oldPersonnelId);
    if (finalEmployeeChanges && finalEmployeeChanges.length > 0) {
        // **UPDATED:** Pass the lockedIdRefMap for logging context
        saveEmployeeInfoBulk(sfcRef, finalEmployeeChanges, year, month, shift, lockedIdRefMap);
    }
    
    // UPDATED CHECK: Call bulk save if there are attendance changes OR if there are deletions
    if (finalAttendanceChanges && finalAttendanceChanges.length > 0 || deletionList.length > 0) {
        // UPDATED: Pass the deletionList
        saveAttendancePlanBulk(sfcRef, contractInfo, finalAttendanceChanges, year, month, shift, group, deletionList);
    }
    
    // *** NEW LOGIC: Log user action after a successful save ***
    // CRITICAL UPDATE: Pass year, month, and shift to ensure log only applies to the current plan period
    logUserActionAfterUnlock(sfcRef, finalEmployeeChanges, finalAttendanceChanges, Session.getActiveUser().getEmail(), year, month, shift);
    // NEW PARAMETERS
    Logger.log(`[saveAllData] Save completed.`);
}

function getOrCreateAuditLogSheet(ss) {
    let sheet = ss.getSheetByName(AUDIT_LOG_SHEET_NAME);
    if (sheet) {
        return sheet;
    }
    
    try {
        sheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME);
        sheet.getRange(1, 1, 1, AUDIT_LOG_HEADERS.length).setValues([AUDIT_LOG_HEADERS]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, AUDIT_LOG_HEADERS.length, 120); 
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

// *** NEW FUNCTION: Get or Create Unlock Request Log Sheet ***
function getOrCreateUnlockRequestLogSheet(ss) {
    let sheet = ss.getSheetByName(UNLOCK_LOG_SHEET_NAME);
    if (sheet) {
        return sheet;
    }
    
    try {
        sheet = ss.insertSheet(UNLOCK_LOG_SHEET_NAME);
        sheet.getRange(1, 1, 1, UNLOCK_LOG_HEADERS.length).setValues([UNLOCK_LOG_HEADERS]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, UNLOCK_LOG_HEADERS.length, 120); 
        Logger.log(`[getOrCreateUnlockRequestLogSheet] Created Unlock Request Log sheet: ${UNLOCK_LOG_SHEET_NAME}`);
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${UNLOCK_LOG_SHEET_NAME}" already exists`)) {
             Logger.log(`[getOrCreateUnlockRequestLogSheet] WARN: Transient sheet creation failure, retrieving existing sheet.`);
            return ss.getSheetByName(UNLOCK_LOG_SHEET_NAME);
        }
        throw e;
    }
}

// **UPDATED:** Added deletionList parameter and logic for row deletion
function saveAttendancePlanBulk(sfcRef, contractInfo, changes, year, month, shift, group, deletionList) { 
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) throw new Error(`AttendancePlan Sheet for ${PLAN_SHEET_NAME} not found.`);
    const HEADER_ROW = PLAN_HEADER_ROW;
    // FIX: Pass context parameters to getHistoricalReferenceMap
    const historicalRefMap = getHistoricalReferenceMap(ss, sfcRef, year, month, shift);
    
    // 1. Read all data
    planSheet.setFrozenRows(0);
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();
    
    let values = [];
    let headers = [];
    if (numRowsToRead >= 0 && numColumns > 0) {
         values = planSheet.getRange(HEADER_ROW, 1, numRowsToRead + 1, numColumns).getDisplayValues();
        headers = values[0]; 
    } else {
        headers = planSheet.getRange(HEADER_ROW, 1, 1, numColumns).getDisplayValues()[0];
    }
    
    // Find Header Indices
    const sfcRefIndex = headers.indexOf('CONTRACT #');
    const headcountIndex = headers.indexOf('TOTAL HEADCOUNT');
    const propOrGrpCodeIndex = headers.indexOf('PROP OR GRP CODE');
    const serviceTypeIndex = headers.indexOf('SERVICE TYPE');
    const sectorIndex = headers.indexOf('SECTOR');
    const payorIndex = headers.indexOf('PAYOR COMPANY');
    const agencyIndex = headers.indexOf('AGENCY');
    const monthIndex = headers.indexOf('MONTH');
    const yearIndex = headers.indexOf('YEAR');
    const shiftIndex = headers.indexOf('PERIOD / SHIFT');
    const groupIndex = headers.indexOf('GROUP');
    const printVersionIndex = headers.indexOf('PRINT VERSION');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('POSITION');
    const areaIndex = headers.indexOf('AREA POSTING');
    const referenceIndex = headers.indexOf('Reference #');
    const day1Index = headers.indexOf('DAY1');
    
    if (sfcRefIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1 || day1Index === -1) {
        throw new Error("Missing critical column in Consolidated Plan sheet.");
    }
    
    // Filter dataRows by current context
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);
    const dataRows = values.slice(1);
    
    // Find the Latest Version Row for each Personnel ID for the current context
    const latestVersionMap = {};
    // Key: Personnel ID, Value: Latest Row (Array)
    dataRows.forEach(row => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const id = cleanPersonnelId(row[personnelIdIndex]);

        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift 
 
    && id) {
            const printVersionString = String(row[printVersionIndex] || '').trim();
            const versionParts = printVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0; 
            
            const existingRow = latestVersionMap[id];
            if (!existingRow 
 
|| 
    version > (parseFloat(existingRow[printVersionIndex].split('-').pop()) || 0)) {
 
                latestVersionMap[id] = row;
            }
        }
    });
    // --- NEW LOGIC START: Handle Physical Deletion for Marked Employees ---
    const rowsToDeleteMap = {};
    // Key: Sheet Row Number, Value: Personnel ID
    if (deletionList && deletionList.length > 0) {
        deletionList.forEach(deletedId => {
            const latestRow = latestVersionMap[deletedId];
            if (latestRow) {
                // Find the index of this row within the full 'values' array (including header)
               
                // The actual row in the sheet is r_index_in_values + HEADER_ROW 
                const r_index_in_values = values.findIndex(r => r === latestRow);
                
                if (r_index_in_values > 0) { // r_index_in_values > 0 means it's a data row
     
                    const sheetRowIndex = r_index_in_values + HEADER_ROW; 
                    rowsToDeleteMap[sheetRowIndex] = deletedId;
                    delete latestVersionMap[deletedId]; // Remove from map so it's not processed later
       
               
                  }
            }
        });
    }

    const sheetRowsToDelete = Object.keys(rowsToDeleteMap).map(Number);
    
    if (sheetRowsToDelete.length > 0) {
        // Sort in reverse order to avoid row index shifting issues
        sheetRowsToDelete.sort((a, b) => b - a);
        sheetRowsToDelete.forEach(rowNum => {
            planSheet.deleteRow(rowNum);
            Logger.log(`[saveAttendancePlanBulk] Deleted latest plan row for ID ${rowsToDeleteMap[rowNum]} at row ${rowNum}.`);
        });
    }
    // --- NEW LOGIC END ---

    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        sanitizedHeadersMap[sanitizeHeader(header)] = index;
    });
    // Group changes by Personnel ID
    const changesByRow = changes.reduce((acc, change) => {
        const key = change.personnelId;
        if (!acc[key]) acc[key] = [];
        acc[key].push(change);
        return acc;
    }, {});
    const rowsToAppend = [];
    const auditLogSheet = getOrCreateAuditLogSheet(ss);
    const userEmail = Session.getActiveUser().getEmail();
    const masterEmployeeMap = getEmployeeMasterData(sfcRef).reduce((map, emp) => { 
        map[emp.id] = { name: emp.name, position: emp.position, area: emp.area };
        return map;
    }, {});
    changes.sort((a, b) => {
        const dateA = new Date(a.dayKey);
        const dateB = new Date(b.dayKey);
        if (dateA.getTime() !== dateB.getTime()) {
            return dateA.getTime() - dateB.getTime();
        }
        return a.personnelId.localeCompare(b.personnelId);
    });
    Object.keys(changesByRow).forEach(personnelId => {
        const dailyChanges = changesByRow[personnelId];
        const latestVersionRow = latestVersionMap[personnelId];
        let newRow;
        let currentVersion = 0;
        const empDetails = masterEmployeeMap[personnelId] || { name: 'N/A', position: '', area: '' };

        
        let nextGroupToUse = group; // Default to the group passed from the client input (e.g., G2)

     
        
    // --- V1 Creation Logic ---
        if (!latestVersionRow) {
            const planHeadersCount = headers.length; 
            
            newRow = Array(planHeadersCount).fill('');
            
            // Set Fixed Metadata 
   
            
            newRow[sfcRefIndex] = sfcRef;
            newRow[headcountIndex] = contractInfo.headCount;
       
            newRow[propOrGrpCodeIndex] = contractInfo.propOrGrpCode;
            newRow[serviceTypeIndex] = contractInfo.serviceType;
            newRow[sectorIndex] = contractInfo.sector;
            newRow[payorIndex] = contractInfo.payor;
            newRow[agencyIndex] = contractInfo.agency;
            newRow[monthIndex] = targetMonthShort;
            newRow[yearIndex] = targetYear;
            newRow[shiftIndex] = shift;
            newRow[groupIndex] = nextGroupToUse;
            newRow[referenceIndex] = '';
            // Blank on initial save

            // Set Fixed Employee Info
            newRow[personnelIdIndex] = personnelId;
            newRow[nameIndex] = empDetails.name;
            newRow[positionIndex] = empDetails.position;
            newRow[areaIndex] = empDetails.area;
            
            currentVersion = 0;
        } else {
            // Standard Update Scenario (Copy Latest Version)
            newRow = [...latestVersionRow];
            const versionString = latestVersionRow[printVersionIndex].split('-').pop();
            currentVersion = parseFloat(versionString) || 0; 

            newRow[referenceIndex] = '';
            // Ito ang critical reset

            // **NEW LOGIC START: Preserve the old GROUP number for versioning**
            const oldGroup = latestVersionRow[groupIndex];
            if (oldGroup && String(oldGroup).trim().toUpperCase() !== String(group).trim().toUpperCase()) {
                 // Force use of the existing group found in the sheet if it's different from the new requested group (e.g., use G1, ignore G2 from client)
                 Logger.log(`[savePlanBulk] WARNING: Overriding requested group ${group} with existing group ${oldGroup} for versioning ID ${personnelId}.`);
                nextGroupToUse = oldGroup;
            }
            // **NEW LOGIC END**

            // IMPORTANT: Update Group number in the row to the chosen group (usually the old one)
            newRow[groupIndex] = nextGroupToUse;
            // IMPORTANT: Update Employee Info (Name/Position/Area) to latest master data 
            newRow[nameIndex] = empDetails.name;
            newRow[positionIndex] = empDetails.position;
            newRow[areaIndex] = empDetails.area;
        }
        
        let isRowChanged = false;
        // Apply all daily changes to the new row and log audit trail
        dailyChanges.forEach(data => {
            const { dayKey, status: newStatus } = data;
            const dayNumber = parseInt(dayKey.split('-')[2], 10); // Actual day of the month (1-31)
            
            // NEW LOGIC: Map the actual day number to the column index 
 
    // (1-16)
            let dayColumnNumber; // 1-16 (Represents the DAYX column number)
            
            if (shift === '1stHalf') {
                dayColumnNumber = dayNumber; // 1 -> 1, 15 -> 15
            } else { // 2ndHalf
     
    
    dayColumnNumber = dayNumber - 15; // 16 -> 1, 31 -> 16
            }
            
            if (dayColumnNumber < 1 || dayColumnNumber > PLAN_MAX_DAYS_IN_HALF) {
                 Logger.log(`[savePlanBulk] WARNING: Day number ${dayNumber} is out of expected range for shift ${shift}. Skipping update.`);
     
 
                
    
    return;
            }
            
            const dayColIndex = day1Index + dayColumnNumber - 1;
            // 0-based index
           
    
            if (dayColIndex >= day1Index && dayColIndex < numColumns) {
                
                const oldStatus = String((latestVersionRow || newRow)[dayColIndex] || '').trim();
                if (oldStatus !== newStatus) {
   
                    isRowChanged = true;
                    newRow[dayColIndex] = newStatus; 
                    
                    const lockedRefNum = historicalRefMap[personnelId] || ''; 
                    
                    // START OF NEW LOGIC: Only log the historical reference if the old status was not blank or 'NA'.
                    const refToLog = (oldStatus && oldStatus !== 'NA') ? lockedRefNum : '';
                    // END OF NEW LOGIC
            
                    const personnelName = empDetails.name;
                    const logEntry = [
                        new Date(), userEmail, sfcRef, personnelId, personnelName, PLAN_SHEET_NAME, 
                        dayKey, shift, `'${refToLog}`, oldStatus, newStatus // MODIFIED: Use refToLog
       
                    ];
                    auditLogSheet.appendRow(logEntry);
                    Logger.log(`[AuditLog] Change logged for ID ${personnelId}, Day ${dayKey}: ${oldStatus} -> ${newStatus} (Ref: ${refToLog})`);
                    // MODIFIED: Use refToLog
                }
            } else {
                Logger.log(`[savePlanBulk] WARNING: Day column index not found for day ${dayNumber}. Skipping update for this cell.`);
            }
        });
        
        // 5. If any change was detected, append the new row with incremented version
        if (isRowChanged) {
            const nextVersion = (currentVersion + 1).toFixed(1);
            // Keep minor increment for save tracking
            // Format: SFC Ref#-PlanPeriod-shift-group-version
            // IMPORTANT: Use the determined nextGroupToUse here 
            const printVersionString = `${sfcRef}-${targetMonthShort}${targetYear}-${shift}-${nextGroupToUse}-${nextVersion}`;
            newRow[printVersionIndex] = printVersionString;
            rowsToAppend.push(newRow);
        }
    });

    // 6. Bulk append new version rows
    if (rowsToAppend.length > 0) {
        const newRowLength = rowsToAppend[0].length;
        const startRow = planSheet.getLastRow() + 1;
        
        planSheet.getRange(startRow, 1, rowsToAppend.length, newRowLength).setValues(rowsToAppend);
        // Set the format for all fixed columns to Plain Text
        planSheet.getRange(startRow, 1, rowsToAppend.length, PLAN_FIXED_COLUMNS).setNumberFormat('@');
        Logger.log(`[saveAttendancePlanBulk] Appended ${rowsToAppend.length} new version rows for ${PLAN_SHEET_NAME}.`);
    }

    planSheet.setFrozenRows(HEADER_ROW); 
    Logger.log(`[saveAttendancePlanBulk] Completed Attendance Plan update for ${PLAN_SHEET_NAME}.`);
}

/**
 * Updates the Reference # column in AttendancePlan_Consolidated for the latest version rows 
 * that were included in the print action.
 */
function updatePlanSheetReferenceBulk(refNum, sfcRef, year, month, shift, printedPersonnelIds) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) return;

    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    const numColumns = planSheet.getLastColumn();
    
    if (lastRow <= HEADER_ROW) return;
    const planValues = planSheet.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, numColumns).getValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);
    // Hanapin ang mga index ng kailangang columns
    const sfcRefIndex = headers.indexOf('CONTRACT #');
    const monthIndex = headers.indexOf('MONTH');
    const yearIndex = headers.indexOf('YEAR');
    const shiftIndex = headers.indexOf('PERIOD / SHIFT');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const printVersionIndex = headers.indexOf('PRINT VERSION');
    const referenceIndex = headers.indexOf('Reference #'); // Ito ang i-u-update natin

    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);
    
    const latestVersionMap = {}; // Key: Personnel ID, Value: Latest Row Data

    dataRows.forEach((row, rowIndex) => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const id = String(row[personnelIdIndex] || '').trim();
        
     
       // 1. I-filter base sa Context at mga ID na na-print
        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift && printedPersonnelIds.includes(id)) {
            
            const printVersionString = String(row[printVersionIndex] || '').trim();
            const versionParts = printVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 
 1]) || 0;
   
            const existingEntry = latestVersionMap[id];
            
            // 2. Hanapin ang LATEST version row index para sa ID na ito
            if (!existingEntry ||
version > existingEntry.version) {
                latestVersionMap[id] = { 
                    rowArray: row, 
                    sheetRowNumber: rowIndex + HEADER_ROW + 1, // Actual row number sa sheet
                    version: version 
   
             };
}
        }
    });

    const rangesToUpdate = [];
    Object.values(latestVersionMap).forEach(entry => {
        // 3. I-handa ang update: Reference # column (index + 1 para sa 1-based column)
        if (referenceIndex !== -1) {
            rangesToUpdate.push({
                row: entry.sheetRowNumber,
                col: referenceIndex + 1, 
               
                 value: refNum
            });
        }
    });
    // 4. Isagawa ang Batch update
    if (rangesToUpdate.length > 0) {
        planSheet.setFrozenRows(0);
        rangesToUpdate.forEach(update => {
             // Dapat i-set ang format sa @ para ma-preserve ang Reference # string
             planSheet.getRange(update.row, update.col).setNumberFormat('@').setValue(update.value);
        });
        planSheet.setFrozenRows(HEADER_ROW);
        Logger.log(`[updatePlanSheetReferenceBulk] Updated Reference # for ${rangesToUpdate.length} personnel in ${PLAN_SHEET_NAME}.`);
    }
}

// **UPDATED:** Added currentLockRef parameter and logic change
function logScheduleDeletion(sfcRef, planSheet, targetShift, personnelId, userEmail, year, month, currentLockRef) { 
    if (!planSheet || planSheet.getName() !== PLAN_SHEET_NAME) return;
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    
    if (lastRow <= HEADER_ROW) return;
    const numColumns = planSheet.getLastColumn();
    const planValues = planSheet.getRange(HEADER_ROW, 1, lastRow - HEADER_ROW + 1, numColumns).getDisplayValues();
    const headers = planValues[0];
    const dataRows = planValues.slice(1);
    // Find Header Indices
    const sfcRefIndex = headers.indexOf('CONTRACT #');
    const monthIndex = headers.indexOf('MONTH');
    const yearIndex = headers.indexOf('YEAR');
    const shiftIndex = headers.indexOf('PERIOD / SHIFT');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const printVersionIndex = headers.indexOf('PRINT VERSION');
    const day1Index = headers.indexOf('DAY1');
    
    if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1 || day1Index === -1) return;
    // 1. Find the LATEST version row in the target context to be logged/deleted
    let latestVersion = 0;
    let targetRowIndex = -1; 
    let rowToDelete = null;
    
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);

    dataRows.forEach((row, index) => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const currentId = cleanPersonnelId(row[personnelIdIndex]);
        
        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift 
 
    === targetShift && currentId === personnelId) {
            
            const printVersionString = String(row[printVersionIndex] || '').trim();
            const versionParts = printVersionString.split('-').pop();
            const currentVersion = parseFloat(versionParts) || 0; 
            
            if (currentVersion >= 
latestVersion) 
 
            {
    
                latestVersion = currentVersion;
                targetRowIndex = index;
                rowToDelete = row;
            }
        }
    });
    if (targetRowIndex === -1) return;
    const auditLogSheet = getOrCreateAuditLogSheet(planSheet.getParent());
    const logEntries = [];
    
    // **CHANGED LOGIC:** Use the passed lock reference instead of historical map lookup
    const lockedRefNum = currentLockRef ||
''; 
    const employeeName = getEmployeeNameFromMaster(sfcRef, personnelId);
    
    // NEW LOGIC: Iterate only through DAY1 to DAY16 columns
    const loopLimit = PLAN_MAX_DAYS_IN_HALF;
    // 16
    const startDayOfMonth = targetShift === '1stHalf' ? 1 : 16;
    const endDayOfMonth = new Date(year, month + 1, 0).getDate();
    
    // Iterate over days (columns DAY1 to DAY16)
    for (let d = 1; d <= loopLimit; d++) {
        const actualDay = startDayOfMonth + d - 1;
        if (actualDay > endDayOfMonth) continue; // Skip day if it exceeds max day of the month

        const dayKey = `${year}-${month + 1}-${actualDay}`;
        const dayColIndex = day1Index + d - 1; 

        if (dayColIndex < numColumns) {
            const oldStatus = String(rowToDelete[dayColIndex] || '').trim();
            if (oldStatus && oldStatus !== 'NA') {
                 const logEntry = [
                    new Date(), userEmail, sfcRef, personnelId, employeeName, PLAN_SHEET_NAME, 
                    dayKey, targetShift, `'${lockedRefNum}`, oldStatus, 'DELETED_ROW' // Log the current lock ref (or blank)
               
                 ];
                logEntries.push(logEntry);
            }
        }
    }
    
    if (logEntries.length > 0) {
        auditLogSheet.getRange(auditLogSheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
        Logger.log(`[logScheduleDeletion] Logged ${logEntries.length} schedule deletions for ID ${personnelId}.`);
    }
}


// **UPDATED:** Added lockedIdRefMap parameter
function saveEmployeeInfoBulk(sfcRef, changes, year, month, shift, lockedIdRefMap) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const empSheet = getOrCreateConsolidatedEmployeeMasterSheet(ss); // Gagamitin ang Consolidated Sheet
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    const userEmail = Session.getActiveUser().getEmail();
    if (!empSheet) throw new Error(`Employee Consolidated Sheet not found.`);
    empSheet.setFrozenRows(0);
    const numRows = empSheet.getLastRow() > 0 ? empSheet.getLastRow() : 1;
    const numColumns = empSheet.getLastColumn() > 0 ? empSheet.getLastColumn() : 5;
    // CONTRACT # + 4 fields
    const values = empSheet.getRange(1, 1, numRows, numColumns).getValues();
    const headers = values[0];
    empSheet.setFrozenRows(1); 
    
    // BAGONG INDICES (Dahil may CONTRACT # na sa Column A)
    const contractRefIndex = headers.indexOf('CONTRACT #');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('Position');
    const areaIndex = headers.indexOf('Area Posting');
    const rowsToAppend = [];
    const rowsToDelete = [];
    
    const personnelIdMap = {};
    // Gagawa ng map gamit ang composite key: [SFC_ID]
    for (let i = 1; i < values.length; i++) { 
        const sfc = String(values[i][contractRefIndex] || '').trim();
        const id = String(values[i][personnelIdIndex] || '').trim();
        if (sfc === sfcRef) {
            personnelIdMap[id] = i + 1;
            // Row number
        }
    }
    
    changes.forEach((data) => {
        const oldId = String(data.oldPersonnelId || '').trim();
        const newId = String(data.id || '').trim();
        
        if (data.isDeleted && oldId && planSheet) {
            rowsToDelete.push({ id: oldId, isMasterDelete: false }); // Log deletion to Plan Log

     
            // NOTE: Hindi na dinelete sa Master sheet ang existing employee, kundi sa Plan lang.
            // Kung mayaman ang data, hahanapin natin ang row number sa empSheet para burahin.
            
            // Temporary fix: If it's deleted from the UI, we only track the deletion for the plan data.
            
 
            return;
        }

  
        // BINAGO START: Logic for Appending (New/Existing Employee added to the Contract)
        if (newId && !data.isDeleted && !personnelIdMap[newId]) { 
            
            const newRow = [];
            
      
            newRow[contractRefIndex] = sfcRef; // CRITICAL: I-tag ang SFC Ref#
            newRow[personnelIdIndex] = data.id;
            newRow[nameIndex] = data.name;
            newRow[positionIndex] = data.position;
            newRow[areaIndex] = data.area;
            
            const finalRow = [];
            for(let i = 0; i < headers.length; i++) {
                finalRow.push(newRow[i] !== undefined ? newRow[i] : '');
            }
            
            rowsToAppend.push(finalRow);
            personnelIdMap[newId] = -1; // Mark as pending
            
            Logger.log(`[saveEmployeeInfoBulk] Appending new/existing employee to Consolidated Master: ${newId}`);
        }
        // BINAGO END
    });
    // 1. Log schedule deletion for current context
    rowsToDelete.forEach(item => {
        const date = new Date(year, month, 1);
        const logYear = date.getFullYear();
        const logMonth = date.getMonth(); 
        
        // **CHANGED LOGIC:** Get current lock reference from the passed map
        const currentLockRef = lockedIdRefMap[item.id] || '';
        
   
        // **UPDATED CALL:** Pass the current lock reference
        logScheduleDeletion(sfcRef, planSheet, shift, item.id, userEmail, logYear, logMonth, currentLockRef);
    });
    // 2. Append new rows (Employee Sheet)
    if (rowsToAppend.length > 0) {
      rowsToAppend.forEach(row => {
          empSheet.appendRow(row); 
      });
    }
    
    Logger.log(`[saveEmployeeInfoBulk] Completed Employee Info update. Appended ${rowsToAppend.length} rows.`);
}

// *** NEW FUNCTION: Get the next available Group Number ***
/**
 * Finds the latest existing Group Number for the specific SFC/Period and returns the next sequential number (G1, G2, etc.).
 * @param {string} sfcRef 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @returns {string} The next Group identifier (e.g., "G2").
 */
function getNextGroupNumber(sfcRef, year, month, shift) {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);

        if (!planSheet || planSheet.getLastRow() < PLAN_HEADER_ROW + 1) return "G1";
        const headers = planSheet.getRange(PLAN_HEADER_ROW, 1, 1, planSheet.getLastColumn()).getValues()[0];
        const sfcRefIndex = headers.indexOf('CONTRACT #');
        const monthIndex = headers.indexOf('MONTH');
        const yearIndex = headers.indexOf('YEAR');
        const shiftIndex = headers.indexOf('PERIOD / SHIFT');
        const groupIndex = headers.indexOf('GROUP');
        if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || groupIndex === -1) {
            Logger.log("[getNextGroupNumber] Missing required headers in Consolidated Plan Sheet.");
            return "G1"; 
        }

        const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
        const targetYear = String(year);

        // Read only relevant columns: SFC, Month, Year, Shift, Group
        const numRows = planSheet.getLastRow() - PLAN_HEADER_ROW;
        const lookupRange = planSheet.getRange(PLAN_HEADER_ROW + 1, 1, numRows, planSheet.getLastColumn());
        const values = lookupRange.getDisplayValues(); 

        let maxGroupNumber = 0;
        values.forEach(row => {
            const currentSfc = String(row[sfcRefIndex] || '').trim();
            const currentMonth = String(row[monthIndex] || '').trim();
            const currentYear = String(row[yearIndex] || '').trim();
            const currentShift = String(row[shiftIndex] || '').trim();
            const currentGroup = String(row[groupIndex] || '').trim().toUpperCase();

            // 
            if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift) {
                // Extract numeric part from group (e.g., 'G1' -> 1)
                const numericPart = parseInt(currentGroup.replace(/[^\d]/g, ''), 10);
                if (!isNaN(numericPart) && numericPart > maxGroupNumber) {
  
                    
                    maxGroupNumber = numericPart;
                }
            }
        });
        // Return the next sequential group number
        const nextGroupNumber = maxGroupNumber + 1;
        return `G${nextGroupNumber}`;

    } catch (e) {
        Logger.log(`[getNextGroupNumber] ERROR: ${e.message}`);
        return "G1";
    }
}
// *** END NEW FUNCTION ***

function getOrCreateSignatoryMasterSheet(ss) {
    let sheet = ss.getSheetByName(SIGNATORY_MASTER_SHEET);
    if (sheet) {
        return sheet;
    }

    try {
        sheet = ss.insertSheet(SIGNATORY_MASTER_SHEET);
        // --- BINAGO START ---
        const headers = ['Signatory Name', 'Designation'];
        // Idinagdag ang Designation
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 200);
        sheet.setColumnWidth(2, 150);
        // Nagdagdag ng column width para sa Designation
        // --- BINAGO END ---
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

// BINAGO START: Updated logic para tanggapin ang Signatory objects at i-save ang bago.
function updateSignatoryMaster(signatories) {
    if (!signatories) return;
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const sheet = getOrCreateSignatoryMasterSheet(ss);
    // I-extract ang lahat ng approvedBy at checkedBy objects na may laman
    const allSignatories = [];
    if (signatories.approvedBy && signatories.approvedBy.name) {
        allSignatories.push(signatories.approvedBy);
    }
    signatories.checkedBy.forEach(item => {
        if (item.name) allSignatories.push(item);
    });
    const uniqueSignatories = {}; 
    allSignatories.forEach(item => {
        // Gumamit ng combination ng Name at Designation bilang key para maging unique
        const key = (item.name.toUpperCase() + item.designation.toUpperCase()).trim(); 
        if (!uniqueSignatories[key]) {
            uniqueSignatories[key] = {
                name: item.name.trim(),
                designation: item.designation.trim()
  
          };
        }
    });
    const existingMasterData = getSignatoryMasterData(); // Ito ay nagbabalik na ng {name, designation}
    
    const newSignatoriesToAppend = [];
    Object.values(uniqueSignatories).forEach(newSig => {
        const isExisting = existingMasterData.some(existSig => 
            existSig.name.toUpperCase() === newSig.name.toUpperCase() && 
            existSig.designation.toUpperCase() === newSig.designation.toUpperCase()
        );

        if (!isExisting) {
            newSignatoriesToAppend.push([newSig.name, newSig.designation]);
        }
    });
    if (newSignatoriesToAppend.length > 0) {
        // Magsisimula sa Column 1 at Column 2
        sheet.getRange(sheet.getLastRow() + 1, 1, newSignatoriesToAppend.length, 2).setValues(newSignatoriesToAppend);
        Logger.log(`[updateSignatoryMaster] Appended ${newSignatoriesToAppend.length} new signatories (Name and Designation).`);
    }
}
// BINAGO END

function getOrCreateLogSheet(ss) {
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (sheet) {
        return sheet;
    }
    
    try {
        sheet = ss.insertSheet(LOG_SHEET_NAME);
        const headers = ['Reference #', 'SFC Ref#', 'Plan Sheet Name (N/A)', 'Plan Period Display', 'Payor Company', 'Agency', 'Sub Property', 'Service Type', 'User Email', 'Timestamp', 'Locked Personnel IDs'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, headers.length, 120); 
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

function getNextReferenceNumber(logSheet) {
    // NOTE: This numerical index function is kept for backward compatibility, but
    // the printing logic will now generate the reference string.
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return 1;
    const range = logSheet.getRange(2, 1, lastRow - 1, 1);
    const refNumbers = range.getValues();
    
    let maxRef = 0;
    refNumbers.forEach(row => {
        // We rely on the row index if the column cannot be parsed numerically
        const currentRef = parseInt(row[0].toString().replace(/[^\d]/g, '')) || 0;
        if (currentRef > maxRef) {
            maxRef = currentRef;
        }
    });
    return maxRef + 1;
}

function logPrintAction(subProperty, sfcRef, contractInfo, year, month, shift) {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);
        
        const logSheetLastRow = logSheet.getLastRow();
        const numLogRows = logSheetLastRow > 1 ? logSheetLastRow - 1 : 0;
        let maxGroupNumber = 0;
        let baseRefParts = [];
        
        // 1. Hanapin ang Max Group Number mula sa PrintLog
        if (numLogRows > 0) {
             const logValues = logSheet.getRange(2, 1, numLogRows, LOG_HEADERS.length).getDisplayValues();
            logValues.forEach(row => {
                const logRefString = String(row[0] || '').trim();
                const currentMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
                
                // I-check lang ang mga logs na tugma sa SFC, Shift, at Buwan
       
                
                if (logRefString.includes(sfcRef) && logRefString.includes(shift) && logRefString.includes(currentMonthShort)) {
                    const parts = logRefString.split('-');
                    
                   
                    if (parts.length === 5) { // Hal. [SFC, Period, Shift, Group, Version]
    
        
                        const groupPart = parts[3]; // e.g., "G1"
                         const numericPart 
= parseInt(groupPart.replace(/[^\d]/g, ''), 10);
                
                          
                        if (!isNaN(numericPart) && numericPart > maxGroupNumber) {
                         
                            maxGroupNumber = numericPart;
                        // Kuhanin ang base parts: SFC-Period-Shift (para gamitin sa final reference)
                        
                            baseRefParts = parts.slice(0, 3);
                        }
                    }
                }
            });
        }
        
        // 2. I-calculate ang Next Group at Version
        const nextGroupNumeric = maxGroupNumber + 1;
        const nextGroup = `G${nextGroupNumeric}`;
        const nextPrintVersion = '1.0'; // Laging 1.0 para sa bagong Group
        
        // 3. Kung walang nahanap na base, i-construct ang default base
        if (baseRefParts.length === 0) {
            const date = new Date(year, month, 1);
            const monthYear = date.toLocaleString('en-US', { month: 'short' }) + date.getFullYear();
            baseRefParts = [sfcRef, monthYear, shift];
        }

        // 4. I-construct ang Final Print Reference String
        const finalPrintReference = `${baseRefParts.join('-')}-${nextGroup}-${nextPrintVersion}`;
        Logger.log(`[logPrintAction] Calculated Print Reference String (New Group Logic): ${finalPrintReference}.`);
        
        return finalPrintReference;
    } catch (e) {
        Logger.log(`[logPrintAction] FATAL ERROR: ${e.message}`);
        throw new Error(`Failed to generate print reference string. Error: ${e.message}`);
    }
}


// MODIFIED: This function now logs the PRINT VERSION string into the 'Reference #' column (Column A).
function recordPrintLogEntry(refNum, subProperty, signatories, sfcRef, contractInfo, year, month, shift, printedPersonnelIds) {
    // refNum now contains the PRINT VERSION string, e.g., "2308-Nov2025-1stHalf-G1-2.0"
    
    if (!refNum) {
        Logger.log(`[recordPrintLogEntry] ERROR: No Reference String provided.`);
        return;
    }
    
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);

        updateSignatoryMaster(signatories);

        // *** BAGONG HAKBANG: I-update ang Reference # sa Consolidated Plan Sheet ***
        updatePlanSheetReferenceBulk(refNum, sfcRef, year, month, shift, printedPersonnelIds);
        // ***********************************************************************

        // *** Bagong Hakbang 3: I-log ang Re-Print Action sa Unlock Log ***
        const userEmail = Session.getActiveUser().getEmail();
        logUserReprintAction(sfcRef, userEmail, printedPersonnelIds);
        // ***************************************************************

        const planSheetName = PLAN_SHEET_NAME;
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
        
        const logEntry = [
            refNum, // Use the new string reference here (e.g., 2308-Nov2025-1stHalf-G1-2.0)
            sfcRef,
            planSheetName, 
            dateRange,     
            contractInfo.payor,
      
            
            contractInfo.agency,
  
            subProperty,         
            contractInfo.serviceType,
            Session.getActiveUser().getEmail(), 
            new Date(),
            printedPersonnelIds.join(',') 
  
        
        ];
        const lastLoggedRow = logSheet.getLastRow();
        const newRow = lastLoggedRow + 1;
        const LOCKED_IDS_COL = LOG_HEADERS.length;
        const logEntryRange = logSheet.getRange(newRow, 1, 1, LOG_HEADERS.length);
        // Ensure first column (Reference #) is set to plain text to hold the string
        logEntryRange.getCell(1, 1).setNumberFormat('@');
        logEntryRange.getCell(1, LOCKED_IDS_COL).setNumberFormat('@');
        logEntryRange.setValues([logEntry]);
        logSheet.getRange(newRow, 1, 1, LOG_HEADERS.length).setHorizontalAlignment('left');

        Logger.log(`[recordPrintLogEntry] Logged and Locked ${printedPersonnelIds.length} IDs using Reference String ${refNum}.`);
    } catch (e) {
        Logger.log(`[recordPrintLogEntry] FATAL ERROR: Failed to log print action ${refNum}. Error: ${e.message}`);
    }
}


function sendRequesterNotification(status, personnelIds, lockedRefNums, personnelNames, requesterEmail) {
  if (requesterEmail === 'UNKNOWN_REQUESTER' || !requesterEmail) return;
  const totalCount = personnelIds.length;
  const uniqueRefNums = [...new Set(lockedRefNums)].sort();
  const subject = `Unlock Request Status: ${status} for ${totalCount} Personnel Schedules (Ref# ${uniqueRefNums.join(', ')})`;
  const combinedRequests = personnelIds.map((id, index) => ({
    id: id,
    ref: lockedRefNums[index],
    name: personnelNames[index] 
  }));
  combinedRequests.sort((a, b) => a.ref.localeCompare(b.ref)); 

  const idList = combinedRequests.map(item => 
    `<li><b>${item.name}</b> (ID ${item.id}) (Ref #: ${item.ref})</li>` 
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
      htmlBody: body, 
      name: 'Attendance Plan Monitor (Status Update)'
    });
    Logger.log(`[sendRequesterNotification] Status ${status} email sent to requester: ${requesterEmail} for ${totalCount} IDs.`);
  } catch (e) {
    Logger.log(`[sendRequesterNotification] Failed to send status email to ${requesterEmail}: ${e.message}`);
  }
}


function unlockPersonnelIds(sfcRef, year, month, shift, personnelIdsToUnlock) {
    const userEmail = Session.getActiveUser().getEmail();
    if (!ADMIN_EMAILS.includes(userEmail)) {
      throw new Error("AUTHORIZATION ERROR: Only admin users can unlock printed schedules. Contact administrator.");
    }

    if (!personnelIdsToUnlock || personnelIdsToUnlock.length === 0) return;

    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;
    
    const LOCKED_IDS_COL_INDEX = LOG_HEADERS.length - 1;
    const range = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length);
    const values = range.getValues(); 
    const rangeToUpdate = [];
    values.forEach((row, rowIndex) => {
      const rowNumInSheet = rowIndex + 2; 
      const lockedIdsString = String(row[LOCKED_IDS_COL_INDEX] || '').trim();
      
      if (lockedIdsString) {
        
        const currentIdsWithPrefix = lockedIdsString.split(',').map(id => id.trim());
        let updatedLockedIds = [...currentIdsWithPrefix];
        let changed = false;

        personnelIdsToUnlock.forEach(unlockId => {
       
          
          const lockedIndex = updatedLockedIds.indexOf(unlockId);
          
          if (lockedIndex > -1) {
             updatedLockedIds.splice(lockedIndex, 1); 
             
             const unlockedPrefixId = `UNLOCKED:${unlockId}`;
         
             if (!updatedLockedIds.includes(unlockedPrefixId)) 
    
             {
 
                updatedLockedIds.push(unlockedPrefixId);
             }
             changed = true;
          }
        });
     
        
       
      
        if (changed) {
          
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
    Logger.log(`[unlockPersonnelIds] Successfully unlocked ${personnelIdsToUnlock.length} IDs. (History preserved with UNLOCKED: prefix.)`);
}

function requestUnlockEmailNotification(sfcRef, year, month, shift, personnelIds, lockedRefNums, personnelNames) { 
  const requestingUserEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const unlockLogSheet = getOrCreateUnlockRequestLogSheet(ss);
  
  // --- NEW LOGIC: Log PENDING Unlock Request ---
  const logEntries = [];
  personnelIds.forEach((id, index) => {
      const name = personnelNames[index];
      const refNum = lockedRefNums[index];
      
      logEntries.push([
          sfcRef,
          id,
          name,
          refNum,
          requestingUserEmail,
          new Date(),
      
          '', // Admin Email (Blank)
          '', // Admin Action Timestamp (Blank)
          'PENDING', // Status
          '', // User Action Type (Blank)
          '' // User Action Timestamp (Blank)
      ]);
  });
  if (logEntries.length > 0) {
      unlockLogSheet.getRange(unlockLogSheet.getLastRow() + 1, 1, logEntries.length, UNLOCK_LOG_HEADERS.length).setValues(logEntries);
      Logger.log(`[requestUnlockEmailNotification] Logged ${logEntries.length} PENDING unlock requests.`);
  }
  // --- END NEW LOGIC ---

  const adminEmails = ADMIN_EMAILS.join(', ');
  const date = new Date(year, month, 1);
  const planPeriod = date.toLocaleString('en-US', { month: 'long', year: 'numeric' });
  const shiftDisplay = (shift === '1stHalf' ? '1st to 15th' : '16th to End');
  const combinedRequests = personnelIds.map((id, index) => ({
    id: id,
    ref: lockedRefNums[index],
    name: personnelNames[index] 
  }));
  combinedRequests.sort((a, b) => a.ref.localeCompare(b.ref));
  const requestDetails = combinedRequests.map(item => {
    return `<li style="font-size: 14px;"><b>${item.name}</b> (ID ${item.id}) (Ref #: ${item.ref})</li>`; 
  }).join('');
  const uniqueRefNums = [...new Set(lockedRefNums)].sort();
  const subjectRefNums = uniqueRefNums.join(', ');
  const subject = `ATTN: Admin Unlock Request - Ref# ${subjectRefNums} for ${sfcRef}`;
  
  const idsEncoded = encodeURIComponent(personnelIds.join(','));
  const refsEncoded = encodeURIComponent(lockedRefNums.join(','));
  const requesterEmailEncoded = encodeURIComponent(requestingUserEmail);
  
  const webAppUrl = ScriptApp.getService().getUrl();
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
           style="background-color: #10b981; color: white; padding: 10px 20px; text-align: center; text-decoration: none; display: inline-block; border-radius: 5px; font-weight: bold;
margin-right: 10px;">
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
      subject: subject, 
      htmlBody: htmlBody, 
      name: 'Attendance Plan Monitor (Automated Request)'
    });
    Logger.log(`[requestUnlockEmailNotification] Sent request email for ${personnelIds.length} IDs to ${adminEmails}`);
    return { success: true, message: `Unlock request sent to Admin(s): ${adminEmails} for ${personnelIds.length} IDs.` };
  
    
    } catch (e) {
    Logger.log(`[requestUnlockEmailNotification] Failed to send email: ${e.message}`);
    return { success: false, message: `WARNING: Failed to send request email. Error: ${e.message}` };
  }
}

// *** NEW HELPER FUNCTION: Log the Admin's decision (Approve/Reject) ***
function logAdminUnlockAction(status, sfcRef, personnelIds, lockedRefNums, requesterEmail, adminEmail) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;
    // Read only the necessary columns for filtering and updating
    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();
    // Indices for efficiency (0-based) based on UNLOCK_LOG_HEADERS
    const SFC_INDEX = 0;
    const ID_INDEX = 1;
    const REF_INDEX = 3;
    const REQUESTER_INDEX = 4;
    const ADMIN_EMAIL_INDEX = 6;
    const ADMIN_ACTION_TIME_INDEX = 7;
    const STATUS_INDEX = 8;
    let updatedCount = 0;

    personnelIds.forEach((id, index) => {
        const targetRef = lockedRefNums[index];
        
        // Find the latest PENDING entry for this ID/Ref/Requester
        let rowIndexToUpdate = -1;
        
        // Iterate backward to find the latest request
        for (let i = values.length - 1; i >= 0; i--) {
     
           const row = values[i];
            const currentStatus = String(row[STATUS_INDEX] || '').trim();
            
            if (String(row[SFC_INDEX]).trim() === sfcRef &&
                String(row[ID_INDEX]).trim() === id &&
                String(row[REF_INDEX]).trim() === targetRef &&
   
                 String(row[REQUESTER_INDEX]).trim() === requesterEmail &&
                currentStatus === 'PENDING') {
                
                rowIndexToUpdate = i + 2; // +1 for 0-base to 1-base, +1 for header row
           
             break;
       
           }
        }

        if (rowIndexToUpdate !== -1) {
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            // Update Admin Email, Timestamp, and Status
            targetRow.getCell(1, ADMIN_EMAIL_INDEX + 1).setValue(adminEmail);
            targetRow.getCell(1, ADMIN_ACTION_TIME_INDEX + 1).setValue(new Date());
            targetRow.getCell(1, STATUS_INDEX + 1).setValue(status);
            updatedCount++;
        }
    });
    Logger.log(`[logAdminUnlockAction] Logged ${status} for ${updatedCount} unlock requests.`);
}

// **UPDATED:** Added year, month, and shift parameters for contextual filtering.
function logUserActionAfterUnlock(sfcRef, employeeChanges, attendanceChanges, userEmail, year, month, shift) { // UPDATED SIGNATURE
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;
    // --- NEW LOGIC: Determine the current Plan Period Identifier ---
    const targetMonthYear = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' }) + year;
    const targetPeriodIdentifier = `${targetMonthYear}-${shift}`; // e.g., Nov2025-1stHalf
    Logger.log(`[logUserActionAfterUnlock] Target Plan Period: ${targetPeriodIdentifier}`);
    // -------------------------------------------------------------
    
    // Read only the necessary columns for filtering and updating
    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();
    // Indices for efficiency (0-based)
    const SFC_INDEX = 0;
    const ID_INDEX = 1;
    const REF_INDEX = 3;
    // Locked Ref # column
    const REQUESTER_INDEX = 4;
    const STATUS_INDEX = 8;
    const USER_ACTION_TYPE_INDEX = 9;
    const USER_ACTION_TIME_INDEX = 10;
    
    // 1. Determine the type of user action based on the save operation
    // ... (rest of the actionType determination logic remains the same)
    const MIN_ROSTER_FOR_ENTIRE_EDIT = 5;
    let actionType = 'Edit Personal AP (Info Only)';
    if (attendanceChanges.length > 0) {
        
        const planEmployees = getEmployeeMasterData(sfcRef);
        // Get all active employees in master list
        const modifiedIDs = new Set(attendanceChanges.map(c => cleanPersonnelId(c.personnelId)));
        const allPlanIDs = new Set(planEmployees.map(e => e.id));
        
        // Check for creation of a new plan/employee within an existing contract
        if (employeeChanges.some(c => c.isNew)) {
             actionType = 'Create AP Plan For an OLD AP';
        } 
        // **UPDATED HEURISTIC**: Check if total employees is >= MIN_ROSTER_FOR_ENTIRE_EDIT
        else if (allPlanIDs.size >= MIN_ROSTER_FOR_ENTIRE_EDIT && 
                 modifiedIDs.size / allPlanIDs.size > 0.6) {
             actionType = 'Edit Entire AP';
        } else {
             actionType = 'Edit Personal AP (Schedule)';
        }
    } else if (employeeChanges.length > 0) {
        actionType = 'Edit Personal AP (Info Only)';
    } else {
        return;
    }

    // 2. Collect all IDs involved in this save action
    const modifiedIdsInSave = new Set([
        ...employeeChanges.map(c => cleanPersonnelId(c.id || c.oldPersonnelId)),
        ...attendanceChanges.map(c => cleanPersonnelId(c.personnelId))
    ]);
    let loggedCount = 0;
    const processedKeys = new Set(); // Key: ID_SFC_Requester to prevent updating the same unlock request

    for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        const rowId = cleanPersonnelId(row[ID_INDEX]);
        const rowSfc = String(row[SFC_INDEX]).trim();
        const rowStatus = String(row[STATUS_INDEX]).trim();
        const rowActionType = String(row[USER_ACTION_TYPE_INDEX]).trim();
        const rowRequester = String(row[REQUESTER_INDEX]).trim();
        // --- NEW FILTERING CONDITION: Check the Locked Ref # (Column D) for the current period ---
        const lockedRef = String(row[REF_INDEX]).trim();
        const isTargetPeriod = lockedRef.includes(targetPeriodIdentifier);
        // --------------------------------------------------------------------------------------
        
        // Find the latest APPROVED entry for this ID that hasn't been logged with a user action yet AND matches the current period
        if (rowSfc === sfcRef && 
            modifiedIdsInSave.has(rowId) &&
            rowStatus === 'APPROVED' &&
            rowActionType === '' &&
 
            isTargetPeriod // *** CRITICAL NEW CHECK ***
           ) {
            
            const rowKey = `${rowId}_${rowSfc}_${rowRequester}`;
            if (processedKeys.has(rowKey)) continue;

            const rowIndexToUpdate = i + 2; // +1 for 0-base to 1-base, +1 for header row
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            // Update User Action Type and Timestamp
            targetRow.getCell(1, USER_ACTION_TYPE_INDEX + 1).setValue(actionType);
            targetRow.getCell(1, USER_ACTION_TIME_INDEX + 1).setValue(new Date());
            
            processedKeys.add(rowKey);
            loggedCount++;
        }
    }
     Logger.log(`[logUserActionAfterUnlock] Logged action type "${actionType}" for ${loggedCount} recently approved unlock requests matching period ${targetPeriodIdentifier}.`);
}

/**
 * Logs a 'Re-Print Attendance Plan' action in the UnlockRequestLog for personnel 
 * who were recently APPROVED for an unlock and are now included in a new print log.
 */
function logUserReprintAction(sfcRef, userEmail, printedPersonnelIds) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;

    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();

    const SFC_INDEX = 0;
    const ID_INDEX = 1;
    const REQUESTER_INDEX = 4;
    const STATUS_INDEX = 8;
    const USER_ACTION_TYPE_INDEX = 9;
    const USER_ACTION_TIME_INDEX = 10;
    let loggedCount = 0;
    const processedKeys = new Set(); // Key: ID_SFC_Requester

    for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        const rowId = String(row[ID_INDEX]).trim();
        const rowSfc = String(row[SFC_INDEX]).trim();
        const rowStatus = String(row[STATUS_INDEX]).trim();
        const rowActionType = String(row[USER_ACTION_TYPE_INDEX]).trim();
        const rowRequester = String(row[REQUESTER_INDEX]).trim();
        // 1. I-filter: Katugma sa SFC, ID ay na-print, Status ay APPROVED, at User Action ay blangko
        if (rowSfc === sfcRef && 
            printedPersonnelIds.includes(rowId) &&
            rowStatus === 'APPROVED' &&
            rowActionType === '' // Hindi pa na-log ang aksyon ng user
           ) {
         
            const rowKey = `${rowId}_${rowSfc}_${rowRequester}`;
            if (processedKeys.has(rowKey)) continue;

            // 2. I-log ang aksyon
            const rowIndexToUpdate = i + 2;
            // Actual row number sa sheet
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            targetRow.getCell(1, USER_ACTION_TYPE_INDEX + 1).setValue('Re-Print Attendance Plan');
            targetRow.getCell(1, USER_ACTION_TIME_INDEX + 1).setValue(new Date());
            
            processedKeys.add(rowKey);
            loggedCount++;
        }
    }
     Logger.log(`[logUserReprintAction] Logged 'Re-Print Attendance Plan' for ${loggedCount} recently approved unlock requests.`);
}

/**
 * Checks the status of the latest request for a batch of personnel IDs in the UnlockRequestLog.
 * Ito ay ginagamitan ng backward iteration para mahanap ang pinakabagong status para sa ID/Ref combo.
 */
function getBatchRequestStatus(sfcRef, personnelIds, lockedRefNums, requesterEmail) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return 'PENDING';
    // Indices for efficiency (0-based)
    const SFC_INDEX = 0;
    const ID_INDEX = 1;
    const REF_INDEX = 3;
    const REQUESTER_INDEX = 4;
    const STATUS_INDEX = 8; 
    
    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();
    // Map ang ID at Ref# ng mga incoming request (Key: ID_Ref#)
    const incomingKeys = new Set(personnelIds.map((id, index) => `${id}_${lockedRefNums[index]}`));
    // Map na mag-iimbak ng pinakabagong status (Key: ID_Ref# -> Value: Status)
    const latestStatusMap = {};
    // 1. Iterate backward para makuha ang PINAKABAGONG STATUS para sa bawat ID_Ref#
    for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        const rowId = String(row[ID_INDEX]).trim();
        const rowRef = String(row[REF_INDEX]).trim();
        const rowKey = `${rowId}_${rowRef}`;
        // I-filter: Katugma sa SFC at Requester, at kasama sa incoming batch
        if (String(row[SFC_INDEX]).trim() === sfcRef &&
            String(row[REQUESTER_INDEX]).trim() === requesterEmail &&
            incomingKeys.has(rowKey) &&
            !latestStatusMap[rowKey] // Only map the most recent one
           ) {
            latestStatusMap[rowKey] = String(row[STATUS_INDEX] || '').trim();
        }
    }
    
    // 2. I-check ang status ng lahat ng incoming requests
    let blockStatus = 'PENDING';
    incomingKeys.forEach(key => {
        // Ang status ay ang pinakabagong nakuha natin, o PENDING kung hindi nahanap sa log
        const status = latestStatusMap[key] || 'PENDING'; 
        
        // Kung ang pinakabagong status ay processed (APPROVED/REJECTED), i-block ang aksyon
        if (status !== 'PENDING') {
            blockStatus = status; 
        }
 
    });

    return blockStatus;
}

function processAdminUnlockFromUrl(params) {
  const idsString = params.id ? decodeURIComponent(params.id) : '';
  const refsString = params.ref ? decodeURIComponent(params.ref) : '';
  const personnelIds = idsString.split(',').map(s => s.trim()).filter(s => s);
  const lockedRefNums = refsString.split(',').map(s => s.trim()).filter(s => s);
  
  const sfcRef = params.sfc;
  const requesterEmail = params.req_email ? decodeURIComponent(params.req_email) : 'UNKNOWN_REQUESTER';
  const personnelNames = personnelIds.map(id => getEmployeeNameFromMaster(sfcRef, id));
  
  if (personnelIds.length === 0 || lockedRefNums.length === 0 || personnelIds.length !== lockedRefNums.length) {
     return HtmlService.createHtmlOutput('<h1 style="color: red;">INVALID REQUEST</h1><p>The Unlock URL is incomplete or the number of Personnel IDs does not match the number of Reference Numbers.</p>');
  }

  const userEmail = Session.getActiveUser().getEmail();
  if (!ADMIN_EMAILS.includes(userEmail)) {
    return HtmlService.createHtmlOutput('<h1 style="color: red;">AUTHORIZATION FAILED</h1><p>You are not authorized to perform this action. Your email: ' + userEmail + '</p>');
  }
  
  // *** START OF ONE-CLICK GUARDRAIL CHECK ***
  // I-check kung na-proseso na ang request (approved/rejected) sa UnlockRequestLog
  const currentBatchStatus = getBatchRequestStatus(sfcRef, personnelIds, lockedRefNums, requesterEmail);
  if (currentBatchStatus !== 'PENDING') {
      const template = HtmlService.createTemplateFromFile('UnlockStatus');
      template.status = 'INFO';
      template.message = `This unlock request has already been processed as **${currentBatchStatus}** by Admin (${userEmail}). No further action will be taken.`;
      return template.evaluate().setTitle('Request Already Processed');
  }
  // *** END OF ONE-CLICK GUARDRAIL CHECK ***
  
  const summary = `${personnelIds.length} schedules (Ref# ${lockedRefNums.join(', ')})`;
  if (params.action === 'reject_info') {
      sendRequesterNotification('REJECTED', personnelIds, lockedRefNums, personnelNames, requesterEmail);
      logAdminUnlockAction('REJECTED', sfcRef, personnelIds, lockedRefNums, requesterEmail, userEmail);
      
      const template = HtmlService.createTemplateFromFile('UnlockStatus');
      template.status = 'INFO';
      template.message = `Admin (${userEmail}) acknowledged the REJECT click for ${summary}.
      Notification sent to ${requesterEmail}.
      No data was changed.
      The locks remain active.`;
      return template.evaluate().setTitle('Reject Status');
  }
  
  if (params.action === 'unlock' && params.yr && params.mon && params.shift) {
      try {
        const year = parseInt(params.yr, 10);
        const month = parseInt(params.mon, 10) - 1; 
        const shift = params.shift;

        unlockPersonnelIds(sfcRef, year, month, shift, personnelIds);
        sendRequesterNotification('APPROVED', personnelIds, lockedRefNums, personnelNames, requesterEmail);
        
        logAdminUnlockAction('APPROVED', sfcRef, personnelIds, lockedRefNums, requesterEmail, userEmail);
        
        const template = HtmlService.createTemplateFromFile('UnlockStatus');
        template.status = 'SUCCESS';
        template.message = `Successfully unlocked ${summary}.
        The Print Locks have been removed by Admin (${userEmail}). Notification sent to ${requesterEmail}.`;
        return template.evaluate().setTitle('Unlock Status');
      } catch (e) {
        const template = HtmlService.createTemplateFromFile('UnlockStatus');
        template.status = 'ERROR';
        template.message = `Failed to unlock ${summary}. Error: ${e.message}`;
        return template.evaluate().setTitle('Unlock Status');
      }
  }
  
  return HtmlService.createHtmlOutput('<h1>Invalid Action</h1><p>The URL provided is incomplete or incorrect.</p>');
}
