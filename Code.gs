const SPREADSHEET_ID = '1qheN_KURc-sOKSngpzVxLvfkkc8StzGv-1gMvGJZdsc';
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; 
const CONTRACTS_SHEET_NAME = 'RefSeries';
const FILE_201_ID = '1i3ISJGbtRU10MmQ1-YG7esyFpg25-3prOxRa-mpAuJM';
const FILE_201_SHEET_NAME = ['MEG'];
const BLACKLIST_FILE_ID = '1i3ISJGbtRU10MmQ1-YG7esyFpg25-3prOxRa-mpAuJM'; 
const BLACKLIST_SHEET_NAMES = ['MEG'];
const FILE_201_ID_COL_INDEX = 0; // Column A (Personnel ID)
const FILE_201_NAME_COL_INDEX = 1; // Column B (Personnel Name)
const FILE_201_BLACKLIST_STATUS_COL_INDEX = 9;
const PLAN_SHEET_NAME = 'AttendancePlan_Consolidated';
const PLAN_HEADER_ROW = 1;
const PLAN_FIXED_COLUMNS = 18;
const PLAN_MAX_DAYS_IN_HALF = 16; 
const REFSERIES_HEADER_ROW = 8;
const SIGNATORY_MASTER_SHEET = 'SignatoryMaster';
const PRINT_FIELD_MASTER_SHEET = 'PrintFieldMaster';
const EMPLOYEE_MASTER_SHEET_NAME = 'EmployeeMaster_Consolidated'; 
const ADMIN_EMAILS = ['mcdmarketingstorage@megaworld-lifestyle.com'];
const LOG_SHEET_NAME = 'PrintLog';
const LOG_HEADERS = [
    'Reference #', 
    'SFC Ref#', 
    'Plan Sheet Name (N/A)', 
    'Plan Period Display', 
    'Payor Company', 
    'Agency',
    'Sub Property',         
    'Service Type',
    'User Email', 
    'Timestamp',
    'Locked Personnel IDs'  
];
const UNLOCK_LOG_SHEET_NAME = 'UnlockRequestLog';
const UNLOCK_LOG_HEADERS = [
    'SFC Ref#',
    'Personnel ID',
    'Personnel Name',
    'Locked Ref #', 
    'Requesting User',
    'Request Timestamp',
    'Admin Email',
    'Admin Action Timestamp',
    'Status (APPROVED/REJECTED)',
    'User Action Type', 
    'User Action Timestamp'
];
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
    startRow = REFSERIES_HEADER_ROW;
    if (numRows < startRow) {
      Logger.log(`[getSheetData] RefSeries sheet has no data starting from Row ${startRow}.`);
      return [];
    }
    numRows = sheet.getLastRow() - startRow + 1;
  } 
  else if (sheetName === PLAN_SHEET_NAME || sheetName === EMPLOYEE_MASTER_SHEET_NAME) {
      startRow = 1;
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
        const empSheet = ss.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
        const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
        return !!empSheet && !!planSheet;
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
        const empHeaders = ['CONTRACT #', 'Personnel ID', 'Personnel Name', 'Position', 'Area Posting'];
        empSheet.getRange(1, 1, 1, empHeaders.length).setValues([empHeaders]);
        empSheet.setFrozenRows(1);
        Logger.log(`[getOrCreateConsolidatedEmployeeMasterSheet] Created Consolidated Employee sheet: ${EMPLOYEE_MASTER_SHEET_NAME}`);
    } 
    return empSheet;
}

function getOrCreatePrintFieldMasterSheet(ss) {
    let sheet = ss.getSheetByName(PRINT_FIELD_MASTER_SHEET);
    if (sheet) {
        return sheet;
    }

    try {
        sheet = ss.insertSheet(PRINT_FIELD_MASTER_SHEET);
        // Headers: SECTION, DEPARTMENT, REMARKS
        const headers = ['SECTION', 'DEPARTMENT', 'REMARKS'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 150);
        sheet.setColumnWidth(2, 150);
        sheet.setColumnWidth(3, 300);
        Logger.log(`[getOrCreatePrintFieldMasterSheet] Created Print Field Master sheet: ${PRINT_FIELD_MASTER_SHEET}`);
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${PRINT_FIELD_MASTER_SHEET}" already exists`)) {
             Logger.log(`[getOrCreatePrintFieldMasterSheet] WARN: Transient sheet creation failure, retrieving existing sheet.`);
             return ss.getSheetByName(PRINT_FIELD_MASTER_SHEET);
        }
        throw e;
    }
}

function getPrintFieldMasterData() {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = getOrCreatePrintFieldMasterSheet(ss);

        if (sheet.getLastRow() < 2) return { sections: [], departments: [], remarks: [] };
        
        // Read all three columns (1 to 3)
        const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
        const values = range.getDisplayValues();

        const sections = new Set();
        const departments = new Set();
        const remarks = new Set();
        
        values.forEach(row => {
            // Force uppercase for consistency in Datalist
            const section = String(row[0] || '').trim().toUpperCase();
            const department = String(row[1] || '').trim().toUpperCase();
            const remark = String(row[2] || '').trim(); // Remarks retains case/format from sheet
            
            if (section) sections.add(section);
            if (department) departments.add(department);
            if (remark) remarks.add(remark);
        });

        return { 
            sections: Array.from(sections).sort(), 
            departments: Array.from(departments).sort(),
            remarks: Array.from(remarks).sort()
        };
    } catch (e) {
        Logger.log(`[getPrintFieldMasterData] ERROR: ${e.message}`);
        return { sections: [], departments: [], remarks: [] };
    }
}

function updatePrintFieldMaster(printFields) {
    // Section and Department are Required (Client-side validation ensures this)
    if (!printFields || !printFields.section || !printFields.department) return;
    
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const sheet = getOrCreatePrintFieldMasterSheet(ss);
    
    // Normalize new entry
    const newSection = printFields.section.trim().toUpperCase();
    const newDepartment = printFields.department.trim().toUpperCase();
    const newRemarks = (printFields.remarks || '').trim(); 

    if (sheet.getLastRow() < 2) {
        // If sheet is empty, save immediately (it's unique by definition)
        const newEntry = [newSection, newDepartment, newRemarks];
        sheet.appendRow(newEntry);
        Logger.log(`[updatePrintFieldMaster] Appended first entry: ${newSection}/${newDepartment}.`);
        return;
    }

    const allValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
    
    const existingSections = new Set();
    const existingDepartments = new Set();
    const existingRemarks = new Set();
    
    allValues.forEach(row => {
        // Add existing values to Sets for quick lookup
        existingSections.add(String(row[0] || '').trim().toUpperCase());
        existingDepartments.add(String(row[1] || '').trim().toUpperCase());
        existingRemarks.add(String(row[2] || '').trim());
    });
    
    let isSectionNew = !existingSections.has(newSection);
    let isDepartmentNew = !existingDepartments.has(newDepartment);
    let isRemarksNew = !existingRemarks.has(newRemarks);
    
    // Check if at least one of the fields is new.
    if (isSectionNew || isDepartmentNew || isRemarksNew) {
        // Construct the new row, placing only the new values in their respective columns.
        // This ensures the sheet only stores unique entries per column (stand-alone values).
        const newEntry = [
            isSectionNew ? newSection : '',
            isDepartmentNew ? newDepartment : '',
            isRemarksNew ? newRemarks : ''
        ];
        
        sheet.appendRow(newEntry);
        Logger.log(`[updatePrintFieldMaster] Appended new unique entries. Section New: ${isSectionNew}, Dept New: ${isDepartmentNew}, Remarks New: ${isRemarksNew}`);
        
    } else {
        Logger.log(`[updatePrintFieldMaster] All fields already exist individually. Skipping new row creation.`);
    }
}

function createContractSheets(sfcRef, year, month, shift) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    getOrCreateConsolidatedEmployeeMasterSheet(ss);
    let planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    const getConsolidatedPlanHeaders = () => {
        const base = [
            'CONTRACT #', 'TOTAL HEADCOUNT', 'PROP OR GRP CODE', 'SERVICE TYPE', 
            'SECTOR', 'PAYOR COMPANY', 'AGENCY', 'MONTH', 'YEAR', 
            'PERIOD / SHIFT', 
            'SAVE GROUP', 
            'SAVE VERSION', 
            'PRINT GROUP', 
            'Reference #',
            'Personnel ID', 'Personnel Name', 'POSITION', 'AREA POSTING'
        ];
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
        
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, PLAN_FIXED_COLUMNS).setNumberFormat('@');
        Logger.log(`[createContractSheets] Created Consolidated Attendance Plan sheet: ${PLAN_SHEET_NAME} with headers at Row ${PLAN_HEADER_ROW}.`);
    } 
}

function ensureContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required to ensure sheets.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    const empSheet = getOrCreateConsolidatedEmployeeMasterSheet(ss);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) {
        createContractSheets(sfcRef, year, month, shift);
        Logger.log(`[ensureContractSheets] Ensured Consolidated Plan sheet existence.`);
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
    const statusKey = findKey(c, 'Status of Agreement (SFC)');
    const contractIdKey = findKey(c, 'Contract Group ID');
    if (!statusKey || !contractIdKey) return false;
    const contractIdValue = (c[contractIdKey] || '').toString().trim();
    if (!contractIdValue) return false; 
    const status = (c[statusKey] || '').toString().trim().toLowerCase();
    const isLive = status === 'live' || status === 'on process - live' || status === "ongoing temporary augmentation";

    return isLive;
  });
  return filteredContracts.map(c => {
    const contractIdKey = findKey(c, 'Contract Group ID');
    const statusKey = findKey(c, 'Status of Agreement (SFC)');
    const payorKey = findKey(c, 'PAYOR');
    const agencyKey = findKey(c, 'SUPPLIER');
    const serviceTypeKey = findKey(c, 'Kind of Service');
    const headCountKey = findKey(c, 'Headcount');
    const sfcRefKey = findKey(c, 'Ref #');
    const propOrGrpCodeKey = findKey(c, 'PROP OR GRP CODE'); 
    const sectorKey = findKey(c, 'Sector'); 
    const kindOfSfcKey = findKey(c, 'Kind of SFC')
    
    return {
      id: contractIdKey ? (c[contractIdKey] || '').toString() : '',     
      status: statusKey ? (c[statusKey] || '').toString() : '',   
      payorCompany: payorKey ? (c[payorKey] || '').toString() : '', 
      agency: agencyKey ? (c[agencyKey] || '').toString() : '',       
      serviceType: serviceTypeKey ? (c[serviceTypeKey] || '').toString() : '',   
      headCount: parseInt(headCountKey ? c[headCountKey] : 0) || 0, 
      sfcRef: sfcRefKey ? (c[sfcRefKey] || '').toString() : '', 
      propOrGrpCode: propOrGrpCodeKey ? (c[propOrGrpCodeKey] || '').toString() : '',
      sector: sectorKey ? (c[sectorKey] || '').toString() : '',
      kindOfSfc: kindOfSfcKey ? (c[kindOfSfcKey] || '').toString() : ''
    };
  });
}

function cleanPersonnelId(rawId) {
    let idString = String(rawId || '').trim();
    return idString.replace(/\D/g, '');
}

function get201FileAllPersonnelDetails() {
    if (FILE_201_ID === 'PUNAN_MO_ITO_NG_201_SPREADSHEET_ID' || !FILE_201_ID) {
        Logger.log('[get201FileAllPersonnelDetails] ERROR: FILE_201_ID is not set.');
        return [];
    }
    
    try {
        const ss = SpreadsheetApp.openById(FILE_201_ID);
        // Assuming FILE_201_SHEET_NAME is an array like ['MEG']
        const sheet = ss.getSheetByName(FILE_201_SHEET_NAME[0]); // Uses index 0 as per logic in source 70
        if (!sheet) {
            Logger.log(`[get201FileAllPersonnelDetails] ERROR: Sheet ${FILE_201_SHEET_NAME[0]} not found.`);
            return [];
        }

        const START_ROW = 2; 
        const lastRow = sheet.getLastRow();
        const NUM_ROWS = lastRow - START_ROW + 1;
        // Basahin hanggang Column J (index 9) para makuha ang Blacklist Status
        const NUM_COLS_TO_READ = FILE_201_BLACKLIST_STATUS_COL_INDEX + 1;
        
        if (NUM_ROWS <= 0) return [];

        const values = sheet.getRange(START_ROW, 1, NUM_ROWS, NUM_COLS_TO_READ).getDisplayValues();
        const allData = values.map(row => {
            const personnelIdRaw = row[FILE_201_ID_COL_INDEX];     // Column A (index 0)
            const personnelNameRaw = row[FILE_201_NAME_COL_INDEX]; // Column B (index 1)
            // Column J (index 9)
            const status = String(row[FILE_201_BLACKLIST_STATUS_COL_INDEX] || '').trim().toUpperCase(); 
            
            const cleanId = cleanPersonnelId(personnelIdRaw);
            const formattedName = String(personnelNameRaw || '').trim().toUpperCase(); 
            const isBlacklisted = status === 'BLACKLISTED';
      
            if (!cleanId || !formattedName) return null; 

            return {
                id: cleanId,
                name: formattedName,
                isBlacklisted: isBlacklisted,
                position: '', // Placeholder (Position/Area is read from other master sheets)
                area: ''      // Placeholder
            };
        }).filter(item => item !== null);

        Logger.log(`[get201FileAllPersonnelDetails] Retrieved ${allData.length} records from 201 file (Full Details).`);
        return allData;

    } catch (e) {
        Logger.log(`[get201FileAllPersonnelDetails] ERROR: ${e.message}`);
        throw new Error(`Failed to access 201 Master File. Error: ${e.message}`);
    }
}

/**
 * Reads Personnel ID (CODE), First Name, and Last Name from the 201 Master File.
 * Filters out BLACKLISTED employees.
 * Formats the name as "LastName, FirstName" and cleans the ID.
 */
function get201FileMasterData() {
    // Check for config error early (as in original function body)
    if (FILE_201_ID === 'PUNAN_MO_ITO_NG_201_SPREADSHEET_ID') {
        Logger.log('[get201FileMasterData] ERROR: FILE_201_ID is not set.');
        return [];
    }

    try {
        // Uses the new function to get all details
        const allPersonnel = get201FileAllPersonnelDetails();
        
        // Filter out blacklisted and map to the format expected by getEmployeeMasterData
        const masterData = allPersonnel
            .filter(e => !e.isBlacklisted)
            .map(e => ({ id: e.id, name: e.name })); // Keep only ID and Name
    
        Logger.log(`[get201FileMasterData] Retrieved ${masterData.length} NON-BLACKLISTED records from 201 file.`);
        return masterData;

    } catch (e) {
        // Propagate error from get201FileAllPersonnelDetails
        Logger.log(`[get201FileMasterData] ERROR: ${e.message}`);
        throw e;
    }
}

function getBlacklistedEmployeesFrom201() {
    if (FILE_201_ID === 'PUNAN_MO_ITO_NG_201_SPREADSHEET_ID' || !FILE_201_ID) {
        Logger.log('[getBlacklistedEmployeesFrom201] ERROR: FILE_201_ID is not set.');
        return [];
    }

    try {
        const ss = SpreadsheetApp.openById(FILE_201_ID);
        // Assuming FILE_201_SHEET_NAME is an array like ['MEG']
        const sheet = ss.getSheetByName(FILE_201_SHEET_NAME[0]); 
        
        if (!sheet) {
            Logger.log(`[getBlacklistedEmployeesFrom201] ERROR: Sheet ${FILE_201_SHEET_NAME[0]} not found.`);
            return [];
        }

        const START_ROW = 2; // Data starts at Row 2 (skipping header at Row 1)
        const lastRow = sheet.getLastRow();
        const NUM_ROWS = lastRow - START_ROW + 1;
        // Basahin hanggang Column J (index 9) para makuha ang Blacklist Status
        const NUM_COLS_TO_READ = FILE_201_BLACKLIST_STATUS_COL_INDEX + 1; 
        
        if (NUM_ROWS <= 0) return [];

        // Magbasa mula Row 2, Column 1 (A), hanggang Column J (index 9)
        const values = sheet.getRange(START_ROW, 1, NUM_ROWS, NUM_COLS_TO_READ).getDisplayValues();
        const blacklistedEmployees = [];

        values.forEach(row => {
            // Index 9 corresponds to Column J
            const status = String(row[FILE_201_BLACKLIST_STATUS_COL_INDEX] || '').trim().toUpperCase();
            
            if (status === 'BLACKLISTED') {
                const personnelIdRaw = row[FILE_201_ID_COL_INDEX];     // Column A (index 0)
                const personnelNameRaw = row[FILE_201_NAME_COL_INDEX]; // Column B (index 1)
                
                const id = cleanPersonnelId(personnelIdRaw); // Use existing helper
                const name = String(personnelNameRaw || '').trim().toUpperCase(); 

                if (id) {  
                    blacklistedEmployees.push({ id: id, name: name });
                }
            }
        });

        Logger.log(`[getBlacklistedEmployeesFrom201] Retrieved ${blacklistedEmployees.length} blacklisted records from 201 file.`);
        return blacklistedEmployees;
    } catch (e) {
        Logger.log(`[getBlacklistedEmployeesFrom201] ERROR: ${e.message}`);
        // Gamitin ang FILE_201_ID para sa error message
        throw new Error(`Failed to access 201 Master File for blacklist check. Error: ${e.message}`);
    }
}

/**
 * Fetches Personnel IDs and Names of all BLACKLISTED employees from the dedicated sheet(s).
 * @returns {Array<Object>} An array of objects {id: string, name: string}.
 */
function getBlacklistData() { 
    return getBlacklistedEmployeesFrom201();
}

function getEmployeeMasterDataForUI(sfcRef) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    
    // 1. Datalist: Non-blacklisted employees merged with EmployeeMaster_Consolidated.
    const datalistEmployees = getEmployeeMasterData(sfcRef);

    // 2. Blacklisted: Listahan ng mga blacklisted lang.
    const blacklisted = getBlacklistData();
    
    // 3. All 201 Personnel: Para sa client-side check kung ang ID/Name ay galing sa 201.
    const all201Personnel = get201FileAllPersonnelDetails(); 
    
    return {
        datalist: datalistEmployees,
        blacklisted: blacklisted,
        all201Personnel: all201Personnel 
    };
}

/**
 * Checks if a specific personnel ID is on the blacklist.
 * @param {string} personnelId The ID to check.
 * @returns {boolean} True if blacklisted, false otherwise.
 */
function checkBlacklistForPersonnel(personnelId) {
  if (!personnelId) return false;
  const cleanId = cleanPersonnelId(personnelId);
  if (!cleanId) return false;
  
  const blacklistedList = getBlacklistData();
  const blacklistMap = blacklistedList.reduce((map, emp) => {
    map[emp.id] = true;
    return map;
  }, {});
  return !!blacklistMap[cleanId];
}

/**
 * Checks if a specific personnel ID OR Name is on the blacklist.
 * @param {string} personnelId The ID provided in the grid.
 * @param {string} personnelName The Name provided in the grid.
 * @returns {object} Returns {isBlacklisted: boolean, reason: string}
 */
function checkBlacklistByIdOrName(personnelId, personnelName) {
    const cleanId = cleanPersonnelId(personnelId);
    const cleanName = String(personnelName || '').trim().toUpperCase();
    const blacklistedList = getBlacklistData();
    const blacklistIdMap = blacklistedList.reduce((map, emp) => {
        map[emp.id] = emp.name; 
        return map;
    }, {});
    const blacklistNameMap = blacklistedList.reduce((map, emp) => {
        map[emp.name] = emp.id; 
        return map;
    }, {});
    if (blacklistIdMap[cleanId]) {
        return { isBlacklisted: true, reason: `Personnel ID ${cleanId} is BLACKLISTED.` };
    }
    
    if (blacklistNameMap[cleanName]) {
        return { isBlacklisted: true, reason: `Personnel Name "${cleanName}" is BLACKLISTED (Linked ID: ${blacklistNameMap[cleanName]}).` };
    }
    
    return { isBlacklisted: false, reason: '' };
}

function getSignatoryMasterData() {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = getOrCreateSignatoryMasterSheet(ss);

        if (sheet.getLastRow() < 2) return [];
        const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2);
        const values = range.getDisplayValues();

        return values.map(row => ({
            name: String(row[0] || '').trim(), 
            designation: String(row[1] || '').trim() 
        })).filter(item => item.name);
    } catch (e) {
        Logger.log(`[getSignatoryMasterData] ERROR: ${e.message}`);
        return [];
    }
}

function getDynamicSheetName(sfcRef, type, year, month, shift) {
    const safeRef = (sfcRef || '').replace(/[\\/?*[]/g, '_');
    if (type === 'employees') {
        return EMPLOYEE_MASTER_SHEET_NAME;
    }
    return `${safeRef} - AttendancePlan`; 
}

function getEmployeeMasterData(sfcRef) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const all201Data = get201FileMasterData();
    const clean201DataMap = {}; 
    all201Data.forEach(e => {
        clean201DataMap[e.id] = { 
            id: e.id, 
            name: e.name, 
            position: '', 
            area: '' 
        };
    });
    const allMasterData = getSheetData(TARGET_SPREADSHEET_ID, EMPLOYEE_MASTER_SHEET_NAME);
    const filteredMasterData = allMasterData.filter(e => {
        const contractRef = String(e['CONTRACT #'] || '').trim();
        return contractRef === sfcRef;
    });
    const finalEmployeeList = filteredMasterData.map(e => {
        const id = cleanPersonnelId(e['Personnel ID']);
        const name = String(e['Personnel Name'] || '').trim();
        const position = String(e['Position'] || '').trim();
        const area = String(e['Area Posting'] || '').trim();
        
        if (clean201DataMap[id]) 
        {
            delete clean201DataMap[id];
   
        }

        return { id, name, position, area };
    }).filter(e => e.id);
    Object.values(clean201DataMap).forEach(emp => {
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
    return employee ? employee.name : 'N/A';
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
    const saveVersionIndex = headers.indexOf('SAVE VERSION'); 
    
    if (personnelIdIndex === -1 || shiftIndex === -1 || sfcRefIndex === -1 || day1Index === -1) return {};
    let targetRow = null;
    let latestVersion = 0;
    let latestDate = new Date(0);
    dataRows.forEach(row => {
        const currentId = cleanPersonnelId(row[personnelIdIndex]);
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        
        if (currentId === cleanId && currentSfc === sfcRef) {
            const saveVersionString = String(row[saveVersionIndex] || '').trim(); 
            const versionParts = saveVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
            const monthShort = String(row[headers.indexOf('MONTH')] || '').trim();
            const yearNum = parseInt(row[headers.indexOf('YEAR')] || '0', 10);
            
            let planDate = new Date(0);
      
            if (monthShort && yearNum) {
             
                planDate = new Date(`${monthShort} 1, ${yearNum}`);
            }
            
            if (version > latestVersion || (version === latestVersion && planDate.getTime() > latestDate.getTime())) {
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
    const targetShift = String(targetRow[shiftIndex] || '').trim();
    const loopLimit = PLAN_MAX_DAYS_IN_HALF; 
    const startDayOfMonth = targetShift === '1stHalf' ? 1 : 16;
    const endDayOfMonth = new Date(targetYear, targetMonth + 1, 0).getDate();

    for (let d = 1; d <= loopLimit; d++) { 
        const dayHeader = `DAY${d}`;
        const dayColIndex = headers.indexOf(dayHeader);
        if (dayColIndex === -1) continue; 
        const status = String(targetRow[dayColIndex] || '').trim();
        const actualDay = startDayOfMonth + d - 1;
        if (actualDay > endDayOfMonth) continue; 
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
    const values = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getDisplayValues();
    const lockedIdRefMap = {};
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
        const currentSfc = String(row[SFC_REF_COL - 1] || '').trim();
        const currentPeriodDisplay = String(row[PERIOD_DISPLAY_COL - 1] || '').trim();

        if (currentSfc !== sfcRef || currentPeriodDisplay !== dateRange) {
        return; 
        }
        
        const refNum = String(row[REF_NUM_COL - 1] || '').trim();
    
        const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim(); 
        
        if (lockedIdsString) {
        const idsList = lockedIdsString.split(',').map(id => id.trim());
            
            idsList.forEach(idWithPrefix => {
                const cleanId = cleanPersonnelId(idWithPrefix);

                if (cleanId.length >= 3 && !idWithPrefix.startsWith('UNLOCKED:')) { 
                    if (!lockedIdRefMap[cleanId]) {                         
                        lockedIdRefMap[cleanId] = refNum;
                    }
                }
            });
        }
    });
    return lockedIdRefMap;
}

/**
 * NEW FUNCTION: Retrieves all IDs/Names/Refs locked under a specific list of reference numbers
 * for the current planning period/SFC.
 * @param {string} sfcRef 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @param {string[]} refNumsToFind List of unique Reference Numbers (e.g., ['2308-Nov2025-1stHalf-0001-P1']).
 * @returns {Array<{id: string, name: string, ref: string}>}
 */
function getAllLockedIdsByRefs(sfcRef, year, month, shift, refNumsToFind) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return [];

    const REF_NUM_COL = 1;
    const SFC_REF_COL = 2; 
    const PERIOD_DISPLAY_COL = 4;
    const LOCKED_IDS_COL = LOG_HEADERS.length;
    const values = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getDisplayValues();
    const lockedList = [];
    const empNameMap = getEmployeeMasterData(sfcRef).reduce((map, emp) => {
        map[emp.id] = emp.name;
        return map;
    }, {});
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
        const currentSfc = String(row[SFC_REF_COL - 1] || '').trim();
        const currentPeriodDisplay = String(row[PERIOD_DISPLAY_COL - 1] || '').trim();
        const refNum = String(row[REF_NUM_COL - 1] || '').trim();

        if (currentSfc === sfcRef && currentPeriodDisplay === dateRange && refNumsToFind.includes(refNum)) {
   
             const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim(); 

            if (lockedIdsString) {
                const idsList = lockedIdsString.split(',').map(id => id.trim());
        
                idsList.forEach(idWithPrefix => {

                    const cleanId = cleanPersonnelId(idWithPrefix);
                     
                    if (cleanId.length >= 3 && !idWithPrefix.startsWith('UNLOCKED:')) {            
                        if (!lockedList.some(item => item.id === cleanId && item.ref === refNum)) {
    
                             lockedList.push({
                           
                                 id: cleanId,              
                                 name: empNameMap[cleanId] || 'N/A',
                                 ref: refNum
                             });
                        }
                    }
                });
            }
        }
    });
    return lockedList;
}

/**
 * FIX: This function is updated to filter the PrintLog by sfcRef, year, month, and shift
 * to ensure that only the most recent Print Version for the CURRENT PERIOD is used for Audit logging.
 */
function getHistoricalReferenceMap(ss, sfcRef, year, month, shift) {
    const logSheet = ss.getSheetByName('PrintLog');
    if (!logSheet || logSheet.getLastRow() < 2) return {};

    const REF_NUM_COL = 1;              
    const SFC_REF_COL = 2;
    const PERIOD_DISPLAY_COL = 4;
    const LOCKED_IDS_COL = LOG_HEADERS.length;
    
    const allValues = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, LOG_HEADERS.length).getDisplayValues();
    const historicalRefMap = {};
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
        const currentSfc = String(row[SFC_REF_COL - 1] || '').trim();
        const currentPeriodDisplay = String(row[PERIOD_DISPLAY_COL - 1] || '').trim();
        if (currentSfc !== sfcRef || currentPeriodDisplay !== dateRange) {
             continue;
        }

        const refNumRaw = String(row[REF_NUM_COL - 1] || '').trim();
        const refNum = refNumRaw;
        const lockedIdsString = String(row[LOCKED_IDS_COL - 1] || '').trim();
        if (refNum) {
            const allIdsInString = lockedIdsString.split(',').map(s => s.trim());
            allIdsInString.forEach(idWithPrefix => {
                const cleanId = cleanPersonnelId(idWithPrefix);
                
                if (cleanId.length >= 3) { 
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
    const saveVersionIndex = headers.indexOf('SAVE VERSION'); 
    const day1Index = headers.indexOf('DAY1');
    if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1 || day1Index === -1) {
         throw new Error("Missing critical column in Consolidated Plan sheet.");
    }
    
    const planMap = {};
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' }); 
    const targetYear = String(year);
    
    const latestVersionMap = {};
    latestVersionMap.group = '1'; 

    dataRows.forEach(row => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        
        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift && id) {
            
            const saveVersionString = String(row[saveVersionIndex] || '').trim(); 
            const versionParts = saveVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
            const mapKey = id; 
       
            const existingRow = latestVersionMap[mapKey];
    
            if (!existingRow || version > (parseFloat(existingRow[saveVersionIndex].split('-').pop()) || 0)) { 
                latestVersionMap[mapKey] = row;
            }
        }
    });
    const latestDataRows = Object.values(latestVersionMap).filter(r => r.length > 0); 
    const employeesInPlan = new Set();
    const employeesDetails = [];
    const startDayOfMonth = shift === '1stHalf' ? 1 : 16;
    const endDayOfMonth = new Date(year, month + 1, 0).getDate();
    const loopLimit = PLAN_MAX_DAYS_IN_HALF;
    
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
            
            for (let d = 1; d <= loopLimit; d++) {
 
                const actualDay = startDayOfMonth + d - 1;    
                if (actualDay > endDayOfMonth) continue; 
                const dayKey = `${year}-${month + 1}-${actualDay}`; 
                const dayColIndex = day1Index + d - 1; 
   
                if (dayColIndex < numColumns) {
                    const status = String(row[dayColIndex] || '').trim();
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
    const regularEmployees = employees.filter(e => e.position !== 'RELIEVER' || e.area !== 'RELIEVER');
    const relieverEmployees = employees.filter(e => e.position === 'RELIEVER' && e.area === 'RELIEVER');

    return { 
        employees: regularEmployees, // Tanging regular employees
        relieverPersonnelList: relieverEmployees, // Listahan para sa Reliever table
        planMap, 
        lockedIds: Object.keys(lockedIdRefMap), 
        lockedIdRefMap: lockedIdRefMap 
    };
}
/**
 * Retrieves the latest version of the Attendance Plan data for a specified period (used for Copy Previous Plan feature).
 * @param {string} sfcRef 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @returns {object} {employees: [], planMap: {}}
 */
function getPlanDataForPeriod(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);

    if (!planSheet || planSheet.getLastRow() <= PLAN_HEADER_ROW) {
        return { employees: [], planMap: {} };
    }
    
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    const numColumns = planSheet.getLastColumn();
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
    const saveVersionIndex = headers.indexOf('SAVE VERSION'); 
    const day1Index = headers.indexOf('DAY1');
    if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1 || day1Index === -1) {
         throw new Error("Missing critical column in Consolidated Plan sheet.");
    }
    
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);
    
    const latestVersionMap = {};

    dataRows.forEach(row => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const rawId = row[personnelIdIndex];
        const id = cleanPersonnelId(rawId);
        
        
        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift && id) {
            
            const saveVersionString = String(row[saveVersionIndex] || '').trim(); 
            const versionParts = saveVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
          
        
            const mapKey = id; 

            const existingRow = latestVersionMap[mapKey];
    
            if (!existingRow || version > (parseFloat(existingRow[saveVersionIndex].split('-').pop()) || 0)) { 
                latestVersionMap[mapKey] = row;
            }
        }
    });
    const latestDataRows = Object.values(latestVersionMap).filter(r => r.length > 0); 
    const employees = [];
    const planMap = {};
    latestDataRows.forEach(row => {
        const id = cleanPersonnelId(row[personnelIdIndex]);
        
        if (id) {
            
            employees.push({
                id: id, 
                name: String(row[nameIndex] || '').trim(),
        
                position: String(row[positionIndex] || '').trim(),
                area: String(row[areaIndex] || '').trim(),
            });
        }
    });
    
    // --- NEW LOGIC START ---
    // Filter out employees where Position and Area are both 'RELIEVER' (case-insensitive)
    const regularEmployees = employees.filter(e => {
        const position = e.position.toUpperCase();
        const area = e.area.toUpperCase();
        // Keep the employee if they are NOT a Reliever OR if one of the fields is not 'RELIEVER'
        return position !== 'RELIEVER' || area !== 'RELIEVER';
    });
    
    Logger.log(`[getPlanDataForPeriod] Retrieved ${employees.length} total employee records; filtered down to ${regularEmployees.length} non-reliever employees.`);
    // --- NEW LOGIC END ---
    
    return { employees: regularEmployees, planMap }; 
}

function saveAllData(sfcRef, contractInfo, employeeChanges, relieverChanges, attendanceChanges, year, month, shift, group) { // ADDED relieverChanges
    Logger.log(`[saveAllData] Starting save for SFC Ref#: ${sfcRef}, Month/Shift: ${month}/${shift}, SAVE GROUP: ${group}`);
    if (!sfcRef) {
      throw new Error("SFC Ref# is required.");
    }
    ensureContractSheets(sfcRef, year, month, shift);
    
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const lockedIdRefMap = getLockedPersonnelIds(ss, sfcRef, year, month, shift);
    const lockedIds = Object.keys(lockedIdRefMap);
    
    // 1. Filter Regular Employee Info Changes (excluding locked/deleted)
    const finalEmployeeChanges = employeeChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.id || change.oldPersonnelId);
        if (lockedIds.includes(idToCheck) && !change.isDeleted) {
             Logger.log(`[saveAllData] Skipping regular employee info update for locked ID: ${idToCheck}`);
            return false;
        }
        return true;
    });
    
    // 2. Filter Reliever Changes (excluding locked)
    const finalRelieverChanges = relieverChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.id);
        if (lockedIds.includes(idToCheck)) {
            Logger.log(`[saveAllData] Skipping reliever entry for locked ID: ${idToCheck}`);
            return false;
        }
        return true;
    });
    
    // 3. Filter Attendance Changes (excluding locked)
    const finalAttendanceChanges = attendanceChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.personnelId);
        if (lockedIds.includes(idToCheck)) {
            Logger.log(`[saveAllData] Skipping attendance plan update for locked ID: ${idToCheck}`);
            return false;
        }
        return true;
    });
    
    // --- UPDATED LOGIC: Create unified deletion list (Regular + Reliever) ---
    const regularDeletions = finalEmployeeChanges.filter(c => c.isDeleted).map(c => c.oldPersonnelId);
    const relieverDeletions = finalRelieverChanges.filter(c => c.isDeleted).map(c => c.oldPersonnelId);
    
    // Combined list of IDs to delete from Plan Sheet
    const deletionList = regularDeletions.concat(relieverDeletions); 
    // --- END UPDATED LOGIC ---

    // CRITICAL: Regular Employee Info Changes MUST be saved to EmployeeMaster_Consolidated
    const regularEmployeeInfoChanges = finalEmployeeChanges.filter(c => !c.isDeleted);
    if (regularEmployeeInfoChanges && regularEmployeeInfoChanges.length > 0) {
        saveEmployeeInfoBulk(sfcRef, regularEmployeeInfoChanges, year, month, shift, lockedIdRefMap);
    }
    
    // CRITICAL: Pass Reliever Changes to saveAttendancePlanBulk for row versioning
    const newRelieverEntries = finalRelieverChanges.filter(c => !c.isDeleted);

    if (finalAttendanceChanges.length > 0 || deletionList.length > 0 || newRelieverEntries.length > 0) {
        // deletionList now includes IDs from deleted relievers
        saveAttendancePlanBulk(sfcRef, contractInfo, finalAttendanceChanges, newRelieverEntries, year, month, shift, group, deletionList);
    }
    
    logUserActionAfterUnlock(sfcRef, finalEmployeeChanges, finalAttendanceChanges, Session.getActiveUser().getEmail(), year, month, shift);
    Logger.log(`[saveAllData] Save completed.`);
}

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

function saveAttendancePlanBulk(sfcRef, contractInfo, changes, relieverChanges, year, month, shift, group, deletionList) { // ADDED relieverChanges
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) throw new Error(`AttendancePlan Sheet for ${PLAN_SHEET_NAME} not found.`);
    const HEADER_ROW = PLAN_HEADER_ROW;
    const historicalRefMap = getHistoricalReferenceMap(ss, sfcRef, year, month, shift);
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
    const saveGroupIndex = headers.indexOf('SAVE GROUP');
    const saveVersionIndex = headers.indexOf('SAVE VERSION');
    const printGroupIndex = headers.indexOf('PRINT GROUP');
    const referenceIndex = headers.indexOf('Reference #'); 
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('POSITION');
    const areaIndex = headers.indexOf('AREA POSTING');
    const day1Index = headers.indexOf('DAY1');
    if (sfcRefIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1 || day1Index === -1) {
        throw new Error("Missing critical column in Consolidated Plan sheet.");
    }
    
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);
    const dataRows = values.slice(1);
    
    const latestVersionMap = {};
    dataRows.forEach(row => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const id = cleanPersonnelId(row[personnelIdIndex]);

        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift && id) {
         
            const saveVersionString = String(row[saveVersionIndex] || '').trim(); 
            const versionParts = saveVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0; 
            const existingRow = latestVersionMap[id];
            if (!existingRow || version > (parseFloat(existingRow[saveVersionIndex].split('-').pop()) || 0)) { 
           
                 latestVersionMap[id] = row;
  
            }
        }
    });

    const rowsToDeleteMap = {};
    if (deletionList && deletionList.length > 0) {
        deletionList.forEach(deletedId => {
            const latestRow = latestVersionMap[deletedId];
            if (latestRow) {
                const r_index_in_values = values.findIndex(r => r === latestRow);
                if (r_index_in_values > 0) { 
       
                    const sheetRowIndex = r_index_in_values + HEADER_ROW; 
                    rowsToDeleteMap[sheetRowIndex] = deletedId;
                    delete latestVersionMap[deletedId]; 
                  }
            
            }
       
 
        });
    }

    const sheetRowsToDelete = Object.keys(rowsToDeleteMap).map(Number);
    if (sheetRowsToDelete.length > 0) {
        sheetRowsToDelete.sort((a, b) => b - a);
        sheetRowsToDelete.forEach(rowNum => {
            planSheet.deleteRow(rowNum);
            Logger.log(`[saveAttendancePlanBulk] Deleted latest plan row for ID ${rowsToDeleteMap[rowNum]} at row ${rowNum}.`);
        });
    }

    const sanitizedHeadersMap = {};
    headers.forEach((header, index) => {
        sanitizedHeadersMap[sanitizeHeader(header)] = index;
    });
    
    // Convert attendance changes into a map for easier processing
    const changesByRow = changes.reduce((acc, change) => {
        const key = change.personnelId;
        if (!acc[key]) acc[key] = [];
        acc[key].push(change);
        return acc;
    }, {});
    
    const rowsToAppend = [];
    const userEmail = Session.getActiveUser().getEmail();
    const masterEmployeeMap = getEmployeeMasterData(sfcRef).reduce((map, emp) => { 
        map[emp.id] = { name: emp.name, position: emp.position, area: emp.area };
        return map;
    }, {});
    
    // --------------------------------------------------------
    // NEW LOGIC: 1. Process NEW RELIEVER ENTRIES 
    // --------------------------------------------------------
    relieverChanges.forEach(reliever => {
        const personnelId = reliever.id;
        
        // Skip if there's already a saved version OR if there are already attendance changes for this ID
        // Note: Relievers are typically new entries, so we expect this to run once.
        if (latestVersionMap[personnelId] || changesByRow[personnelId]) {
             Logger.log(`[saveAttendancePlanBulk] WARNING: Skipping new reliever entry for existing ID/pending schedule: ${personnelId}`);
             return;
        }
        
        const planHeadersCount = headers.length; 
        
        let newRow = Array(planHeadersCount).fill('');            
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
        newRow[saveGroupIndex] = group;
        newRow[printGroupIndex] = '';
        newRow[referenceIndex] = '';
        newRow[personnelIdIndex] = personnelId;
        newRow[nameIndex] = reliever.name;
        newRow[positionIndex] = 'RELIEVER'; 
        newRow[areaIndex] = 'RELIEVER';     

        // Set DAY columns to empty (Walang schedule)
        for (let d = day1Index; d < numColumns; d++) {
             newRow[d] = ''; 
        }

        const nextVersion = '1.0';
        newRow[saveVersionIndex] = `${sfcRef}-${targetMonthShort}${targetYear}-${shift}-${group}-${nextVersion}`;
        
        rowsToAppend.push(newRow);
        
        // Mark as latest version
        latestVersionMap[personnelId] = newRow; 
        
        Logger.log(`[saveAttendancePlanBulk] Appending new RELIEVER row: ${personnelId} with version ${nextVersion}`);
    });
    // --------------------------------------------------------
    // END NEW LOGIC
    // --------------------------------------------------------


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
        
        // Use details from Employee Master (Regular) or the saved Plan (Reliever is already in latestVersionMap)
        const empDetails = masterEmployeeMap[personnelId] || { 
            name: (latestVersionRow ? latestVersionRow[nameIndex] : 'N/A'), 
            position: (latestVersionRow ? latestVersionRow[positionIndex] : ''), 
            area: (latestVersionRow ? latestVersionRow[areaIndex] : '')
        };    
        let nextGroupToUse = group; 

        if (!latestVersionRow) {
          
            const planHeadersCount = headers.length; 
            
            newRow = Array(planHeadersCount).fill('');            
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
            newRow[saveGroupIndex] = nextGroupToUse;
            newRow[printGroupIndex] = '';
            newRow[referenceIndex] = '';
            newRow[personnelIdIndex] = personnelId;
            newRow[nameIndex] = empDetails.name;
            newRow[positionIndex] = empDetails.position;
            newRow[areaIndex] = empDetails.area;
  
            currentVersion = 0;
        } else {
            newRow = [...latestVersionRow];
            const versionString = latestVersionRow[saveVersionIndex].split('-').pop(); 
            currentVersion = parseFloat(versionString) || 0; 

            newRow[referenceIndex] = '';
            newRow[printGroupIndex] = ''; 
            const oldGroup = latestVersionRow[saveGroupIndex];
            if (oldGroup && String(oldGroup).trim().toUpperCase() !== String(group).trim().toUpperCase()) {
                Logger.log(`[savePlanBulk] WARNING: Overriding requested group ${group} with existing group ${oldGroup} for versioning ID ${personnelId}.`);
                nextGroupToUse = oldGroup;
            }

            newRow[saveGroupIndex] = nextGroupToUse;
            // Update Name/Position/Area only if it's NOT a Reliever, OR if employee info change was also requested
            if (empDetails.position !== 'RELIEVER') {
                newRow[nameIndex] = empDetails.name;
                newRow[positionIndex] = empDetails.position;
                newRow[areaIndex] = empDetails.area;
            }
        }
        
        let isRowChanged = false;
        dailyChanges.forEach(data => {
            const { dayKey, status: newStatus } = data;
            const dayNumber = parseInt(dayKey.split('-')[2], 10); 
            
            let dayColumnNumber; 
            
            if (shift === '1stHalf') {
        
                dayColumnNumber = dayNumber; 
            } else { 
                dayColumnNumber = dayNumber - 15; 
            }
            
            if (dayColumnNumber < 1 || dayColumnNumber > PLAN_MAX_DAYS_IN_HALF) {
 
                Logger.log(`[savePlanBulk] WARNING: Day number ${dayNumber} is out of expected range for shift ${shift}. Skipping update.`);
            return;
            }
            
            const dayColIndex = day1Index + dayColumnNumber - 1;
        
       
            if (dayColIndex >= day1Index && dayColIndex < numColumns) {
 
                const oldStatus = String((latestVersionRow || newRow)[dayColIndex] || '').trim();
                if (oldStatus !== newStatus) {
                    isRowChanged = true;
                    newRow[dayColIndex] = newStatus; 
                    
                    Logger.log(`[saveAttendancePlanBulk] Change applied for ID ${personnelId}, Day ${dayKey}: ${oldStatus} -> ${newStatus} (User: ${userEmail})`);
                }
            } else {
                Logger.log(`[savePlanBulk] WARNING: Day column index not found for day ${dayNumber}. Skipping update for this cell.`);
            }
        });
        
        if (isRowChanged) {
            const nextVersion = (currentVersion + 1).toFixed(1);
            newRow[saveVersionIndex] = `${sfcRef}-${targetMonthShort}${targetYear}-${shift}-${nextGroupToUse}-${nextVersion}`;
            rowsToAppend.push(newRow);
        }
    });

    if (rowsToAppend.length > 0) {
        const newRowLength = rowsToAppend[0].length;
        const startRow = planSheet.getLastRow() + 1;
        
        planSheet.getRange(startRow, 1, rowsToAppend.length, newRowLength).setValues(rowsToAppend);
        planSheet.getRange(startRow, 1, rowsToAppend.length, PLAN_FIXED_COLUMNS).setNumberFormat('@');
        Logger.log(`[saveAttendancePlanBulk] Appended ${rowsToAppend.length} new version rows (including relievers) for ${PLAN_SHEET_NAME}.`);
    }

    planSheet.setFrozenRows(HEADER_ROW); 
    Logger.log(`[saveAttendancePlanBulk] Completed Attendance Plan update for ${PLAN_SHEET_NAME}.`);
}

function updatePlanSheetReferenceBulk(refNum, printGroup, sfcRef, year, month, shift, printedPersonnelIds) { 
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
    const sfcRefIndex = headers.indexOf('CONTRACT #');
    const monthIndex = headers.indexOf('MONTH');
    const yearIndex = headers.indexOf('YEAR');
    const shiftIndex = headers.indexOf('PERIOD / SHIFT');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const saveVersionIndex = headers.indexOf('SAVE VERSION');
    const printGroupIndex = headers.indexOf('PRINT GROUP');
    const referenceIndex = headers.indexOf('Reference #');
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);
    
    const latestVersionMap = {};
    dataRows.forEach((row, rowIndex) => {
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const id = String(row[personnelIdIndex] || '').trim();
      
        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift && printedPersonnelIds.includes(id)) 
        { 
            const saveVersionString = String(row[saveVersionIndex] || '').trim(); 
            const versionParts = saveVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
            const existingEntry = latestVersionMap[id];
      
            if (!existingEntry || version > existingEntry.version) 
            {
        
                latestVersionMap[id] = { 
                    rowArray: row, 
                    sheetRowNumber: rowIndex + HEADER_ROW + 1, 
                    version: version 
   
                 };
            }
        }
    });

    const rangesToUpdate = [];
    Object.values(latestVersionMap).forEach(entry => {
        if (referenceIndex !== -1) {
            rangesToUpdate.push({
                row: entry.sheetRowNumber,
                col: referenceIndex + 1, 
                value: refNum
            });
        
        }
        if (printGroupIndex !== -1) {
            rangesToUpdate.push({
                row: entry.sheetRowNumber,
                col: printGroupIndex + 1, 
                value: printGroup
            });
 
          }
  
    });
    if (rangesToUpdate.length > 0) {
        planSheet.setFrozenRows(0);
        rangesToUpdate.forEach(update => {
             planSheet.getRange(update.row, update.col).setNumberFormat('@').setValue(update.value);
        });
        planSheet.setFrozenRows(HEADER_ROW);
        Logger.log(`[updatePlanSheetReferenceBulk] Updated Reference # and PRINT GROUP for ${rangesToUpdate.length / 2} personnel in ${PLAN_SHEET_NAME}.`);
    }
}

function saveEmployeeInfoBulk(sfcRef, changes, year, month, shift, lockedIdRefMap) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const empSheet = getOrCreateConsolidatedEmployeeMasterSheet(ss);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    const userEmail = Session.getActiveUser().getEmail();
    const historicalRefMapForAudit = getHistoricalReferenceMap(ss, sfcRef, year, month, shift);
    if (!empSheet) throw new Error(`Employee Consolidated Sheet not found.`);
    empSheet.setFrozenRows(0);
    const numRows = empSheet.getLastRow() > 0 ? empSheet.getLastRow() : 1;
    const numColumns = empSheet.getLastColumn() > 0 ? empSheet.getLastColumn() : 5;
    const values = empSheet.getRange(1, 1, numRows, numColumns).getValues();
    const headers = values[0];
    empSheet.setFrozenRows(1); 
    
    const contractRefIndex = headers.indexOf('CONTRACT #');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('Position');
    const areaIndex = headers.indexOf('Area Posting');
    const rowsToAppend = [];
    const rowsToDelete = [];
    const personnelIdMap = {};
    for (let i = 1; i < values.length; i++) { 
        const sfc = String(values[i][contractRefIndex] || '').trim();
        const id = String(values[i][personnelIdIndex] || '').trim();
        if (sfc === sfcRef) {
            personnelIdMap[id] = i + 1;
        }
    }
    
    changes.forEach((data) => {
        const oldId = String(data.oldPersonnelId || '').trim();
        const newId = String(data.id || '').trim();
        
        if (data.isDeleted && oldId && planSheet) {
            rowsToDelete.push({ id: oldId, isMasterDelete: false }); 
            return;
        
        }
        if (newId && !data.isDeleted && !personnelIdMap[newId]) {    
            const newRow = [];
            newRow[contractRefIndex] = sfcRef;
            newRow[personnelIdIndex] = data.id;
            newRow[nameIndex] = data.name;
            newRow[positionIndex] = data.position;
   
            newRow[areaIndex] = data.area;
            
            const finalRow = [];
            for(let i = 0; i < headers.length; i++) {
                finalRow.push(newRow[i] !== undefined ? newRow[i] : '');
         
            }

            rowsToAppend.push(finalRow);
    
            personnelIdMap[newId] = -1;
            Logger.log(`[saveEmployeeInfoBulk] Appending new/existing employee to Consolidated Master: ${newId}`);
        }
    });
    
    if (rowsToAppend.length > 0) {
      rowsToAppend.forEach(row => {
          empSheet.appendRow(row); 
      });
    }
    Logger.log(`[saveEmployeeInfoBulk] Completed Employee Info update. Appended ${rowsToAppend.length} rows.`);
}

function getNextGroupNumber(sfcRef, year, month, shift) {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
        if (!planSheet || planSheet.getLastRow() < PLAN_HEADER_ROW + 1) return "S1";
        const headers = planSheet.getRange(PLAN_HEADER_ROW, 1, 1, planSheet.getLastColumn()).getValues()[0];
        const sfcRefIndex = headers.indexOf('CONTRACT #');
        const monthIndex = headers.indexOf('MONTH');
        const yearIndex = headers.indexOf('YEAR');
        const shiftIndex = headers.indexOf('PERIOD / SHIFT');
        const saveGroupIndex = headers.indexOf('SAVE GROUP');
        if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || saveGroupIndex === -1) {
            Logger.log("[getNextGroupNumber] Missing required headers in Consolidated Plan Sheet.");
            return "S1"; 
        }

        const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
        const targetYear = String(year);

        const numRows = planSheet.getLastRow() - PLAN_HEADER_ROW;
        const lookupRange = planSheet.getRange(PLAN_HEADER_ROW + 1, 1, numRows, planSheet.getLastColumn());
        const values = lookupRange.getDisplayValues(); 

        let maxGroupNumber = 0;
        values.forEach(row => {
            const currentSfc = String(row[sfcRefIndex] || '').trim();
            const currentMonth = String(row[monthIndex] || '').trim();
            const currentYear = String(row[yearIndex] || '').trim();
            const currentShift = String(row[shiftIndex] || '').trim();
            const currentSaveGroup = String(row[saveGroupIndex] || '').trim().toUpperCase(); 

    
            if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift) {
                const numericPart = parseInt(currentSaveGroup.replace(/[^\d]/g, ''), 10);
                if (!isNaN(numericPart) && numericPart > maxGroupNumber) {
                    maxGroupNumber = numericPart;
       
                }
            }
        });
        const nextGroupNumber = maxGroupNumber + 1;
        return `S${nextGroupNumber}`; 

    } catch (e) {
        Logger.log(`[getNextGroupNumber] ERROR: ${e.message}`);
        return "S1";
    }
}

function getOrCreateSignatoryMasterSheet(ss) {
    let sheet = ss.getSheetByName(SIGNATORY_MASTER_SHEET);
    if (sheet) {
        return sheet;
    }

    try {
        sheet = ss.insertSheet(SIGNATORY_MASTER_SHEET);
        const headers = ['Signatory Name', 'Designation'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 200);
        sheet.setColumnWidth(2, 150);
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

function updateSignatoryMaster(signatories) {
    if (!signatories) return;
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const sheet = getOrCreateSignatoryMasterSheet(ss);
    const allSignatories = [];
    if (signatories.approvedBy && signatories.approvedBy.name) {
        allSignatories.push(signatories.approvedBy);
    }
    signatories.checkedBy.forEach(item => {
        if (item.name) allSignatories.push(item);
    });
    const uniqueSignatories = {}; 
    allSignatories.forEach(item => {
        const key = (item.name.toUpperCase() + item.designation.toUpperCase()).trim(); 
        if (!uniqueSignatories[key]) {
            uniqueSignatories[key] = {
                name: item.name.trim(),
                designation: item.designation.trim()
            };
        }
  
    });
    const existingMasterData = getSignatoryMasterData(); 
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
        sheet.getRange(sheet.getLastRow() + 1, 1, newSignatoriesToAppend.length, 2).setValues(newSignatoriesToAppend);
        Logger.log(`[updateSignatoryMaster] Appended ${newSignatoriesToAppend.length} new signatories (Name and Designation).`);
    }
}

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

/**
 * Finds the maximum sequential number used in the base reference 
 * (SFC-MonthYear-Shift-000X) across ALL TIME for the current SFC and returns the next number.
 * PADDING IS REMOVED once the number exceeds 9999.
 */
function getNextSequentialNumber(logSheet, sfcRef) { 
    const logValues = logSheet.getLastRow() > 1 ?
        logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues() : [];
    
    let maxSequentialNumber = 0;
    
    logValues.forEach(row => {
        const logRefString = String(row[0] || '').trim();
        const parts = logRefString.split('-');
        
        let currentSfc = '';
        let sequentialPart = '0';

        if (parts.length === 5) {
            currentSfc = parts[0];
            sequentialPart = parts[3];
        } 
        else if (parts.length === 6) {
            currentSfc = parts[1]; // SFC is now index 1
            sequentialPart = parts[4]; // Sequence is now index 4
        }

        // Compare using the detected SFC
        if (currentSfc === sfcRef) { 
            const numericPart = parseInt(sequentialPart, 10);
            if (!isNaN(numericPart)) {
                if (numericPart > maxSequentialNumber) {
                    maxSequentialNumber = numericPart;
                }
            }
        }
    });

    const nextNumber = maxSequentialNumber + 1;
    // Apply padding only up to 9999.
    if (nextNumber <= 9999) {
        return String(nextNumber).padStart(4, '0');
    }
    
    // Return as a full number string if > 9999
    return String(nextNumber);
}

function logPrintAction(subProperty, sfcRef, contractInfo, year, month, shift) { 
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);
        
        const logSheetLastRow = logSheet.getLastRow();
        const numLogRows = logSheetLastRow > 1 ? logSheetLastRow - 1 : 0;
        const date = new Date(year, month, 1);
        const monthYear = date.toLocaleString('en-US', { month: 'short' }) + date.getFullYear();
        
        let maxGroupNumber = 0;
        let finalSequentialNumber = '';
        let foundUnlockedBaseRefSequential = '';
        const unlockedSequentialNumbers = new Set();
        
        if (numLogRows > 0) {
            const logValues = logSheet.getRange(2, 1, numLogRows, LOG_HEADERS.length).getDisplayValues();
            logValues.forEach(row => {
                const logRefString = String(row[0] || '').trim();
                const parts = logRefString.split('-');
                
                let isMatch = false;
                let sequentialPart = '';
                let groupPart = '';

                if (parts.length === 5 && parts[0] === sfcRef && parts[1] === monthYear && parts[2] === shift) {
                    sequentialPart = parts[3];
                    groupPart = parts[4];
                    isMatch = true;
                }

                else if (parts.length === 6 && parts[1] === sfcRef && parts[2] === monthYear && parts[3] === shift) {
                    sequentialPart = parts[4];
                    groupPart = parts[5];
                    isMatch = true;
                }
                
                if (isMatch) {
                    const lockedIdsString = String(row[LOG_HEADERS.length - 1] || '').trim();
                    const isExplicitlyUnlocked = lockedIdsString.includes('UNLOCKED:'); 

                    if (isExplicitlyUnlocked) {
                        unlockedSequentialNumbers.add(sequentialPart);
                        const numericPart = parseInt(groupPart.replace(/[^\d]/g, ''), 10);
                        if (!isNaN(numericPart) && numericPart > maxGroupNumber) {
                            maxGroupNumber = numericPart;
                            foundUnlockedBaseRefSequential = sequentialPart;
                        }
                    }
                }
            });
        }
        
        let nextPrintGroupNumeric = 1;
        if (unlockedSequentialNumbers.size > 1) {
            nextPrintGroupNumeric = 1;
            finalSequentialNumber = getNextSequentialNumber(logSheet, sfcRef);
            Logger.log(`[logPrintAction] MULTIPLE UNLOCKS. Generating NEW sequential: ${finalSequentialNumber}.`);
        } else if (unlockedSequentialNumbers.size === 1) {
            nextPrintGroupNumeric = maxGroupNumber + 1;
            finalSequentialNumber = foundUnlockedBaseRefSequential;
            Logger.log(`[logPrintAction] SINGLE UNLOCK. Reusing sequential ${finalSequentialNumber}, Group P${nextPrintGroupNumeric}`);
        } else {
            nextPrintGroupNumeric = 1;
            finalSequentialNumber = getNextSequentialNumber(logSheet, sfcRef); 
            Logger.log(`[logPrintAction] New sequential: ${finalSequentialNumber}.`);
        }
        
        const finalPrintGroup = `P${nextPrintGroupNumeric}`;
        
        // --- UPDATED LOGIC: RA Check Only + Fallback ---
        let rawKind = (contractInfo.kindOfSfc || 'SFC').toString().trim().toUpperCase();
        let kindOfSfc = '';

        if (rawKind.startsWith('RA') || rawKind.includes(' RA')) {
            kindOfSfc = 'RA';
        } else {

            kindOfSfc = rawKind.replace(/[^A-Z0-9]/g, '');
            
            if (kindOfSfc.length > 8) {
                kindOfSfc = kindOfSfc.substring(0, 8); 
            }
        }
        
        const baseRef = [kindOfSfc, sfcRef, monthYear, shift, finalSequentialNumber];
        const finalPrintReference = `${baseRef.join('-')}-${finalPrintGroup}`;

        Logger.log(`[logPrintAction] Generated Ref: ${finalPrintReference}.`);
        return { refNum: finalPrintReference, printGroup: finalPrintGroup };
        
    } catch (e) {
        Logger.log(`[logPrintAction] FATAL ERROR: ${e.message}`);
        throw new Error(`Failed to generate print reference string. Error: ${e.message}`);
    }
}

function recordPrintLogEntry(refNum, printGroup, subProperty, printFields, signatories, sfcRef, contractInfo, year, month, shift, printedPersonnelIds) { 
    
    if (!refNum) {
        Logger.log(`[recordPrintLogEntry] ERROR: No Reference String provided.`);
        return;
    }
    
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);
        
        // NEW: Update the Print Field Master sheet
        updatePrintFieldMaster(printFields); 
        
        updateSignatoryMaster(signatories);

        updatePlanSheetReferenceBulk(refNum, printGroup, sfcRef, year, month, shift, printedPersonnelIds);
        
        logUserReprintAction(sfcRef, Session.getActiveUser().getEmail(), printedPersonnelIds);

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
            refNum, 
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
    
    const date = new Date(year, month, 1);
    const monthName = date.toLocaleString('en-US', { month: 'long' });
    const yearNum = date.getFullYear();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    let targetPeriodDisplay = '';
    if (shift === '1stHalf') {
        targetPeriodDisplay = `${monthName} 1-15, ${yearNum} (${shift})`;
    } else {
        targetPeriodDisplay = `${monthName} 16-${daysInMonth}, ${yearNum} (${shift})`;
    }

    const LOCKED_IDS_COL_INDEX = LOG_HEADERS.length - 1;
    const SFC_REF_COL_INDEX = 1;
    const PERIOD_DISPLAY_COL_INDEX = 3;
    const values = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
    const rangeToUpdate = [];
    values.forEach((row, rowIndex) => {
      const currentSfc = String(row[SFC_REF_COL_INDEX] || '').trim();
      const currentPeriodDisplay = String(row[PERIOD_DISPLAY_COL_INDEX] || '').trim();
      
      if (currentSfc !== sfcRef || currentPeriodDisplay !== targetPeriodDisplay) {
          return; 
      }
      
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
         
            if (!updatedLockedIds.includes(unlockedPrefixId)) {
                updatedLockedIds.push(unlockedPrefixId);
            }
             changed = true;
            }
        });
    
        if (changed) {
          const newLockedIdsString = updatedLockedIds.filter(id => id.length > 0).join(',');
          rangeToUpdate.push({
              row: rowNumInSheet,
              col: LOCKED_IDS_COL_INDEX + 1, 
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
    Logger.log(`[unlockPersonnelIds] Successfully unlocked ${personnelIdsToUnlock.length} IDs for period ${targetPeriodDisplay}. (No historical locks removed.)`);
}

/**
 * MODIFIED: Ngayon, tumatanggap na lang ng array ng UNIQUE lockedRefNums (strings) mula sa client.
 * Palalawakin (i-e-expand) ng function na ito ang Ref#s para maging listahan ng IDs para sa email/log.
 * @param {string} sfcRef 
 * @param {number} year 
 * @param {number} month 
 * @param {string} shift 
 * @param {string[]} lockedRefNums (UNIQUE list of Reference # strings to unlock)
 * @returns 
 */
function requestUnlockEmailNotification(sfcRef, year, month, shift, lockedRefNums) { 
    if (!lockedRefNums || lockedRefNums.length === 0) {
        return { success: false, message: 'ERROR: No Reference Numbers selected for unlock.'
};
    }

    const combinedRequests = getAllLockedIdsByRefs(sfcRef, year, month, shift, lockedRefNums);
    if (combinedRequests.length === 0) {
        return { success: false, message: 'ERROR: Could not find any locked schedules associated with the selected Reference Numbers for this period.'
};
    }

    const requestingUserEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const unlockLogSheet = getOrCreateUnlockRequestLogSheet(ss);
    combinedRequests.sort((a, b) => a.ref.localeCompare(b.ref));
    const logEntries = [];
    combinedRequests.forEach(item => {
        logEntries.push([
            sfcRef,
            item.id,
            item.name,
            item.ref,
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

    const personnelIds = combinedRequests.map(item => item.id);
    const personnelNames = combinedRequests.map(item => item.name);
    const expandedRefNums = combinedRequests.map(item => item.ref);
    
    const adminEmails = ADMIN_EMAILS.join(', ');
    const date = new Date(year, month, 1);
    const planPeriod = date.toLocaleString('en-US', { month: 'long', year: 'numeric' });
    const shiftDisplay = (shift === '1stHalf' ? '1st to 15th' : '16th to End');
    const requestDetails = combinedRequests.map(item => {
        return `<li style="font-size: 14px;"><b>${item.name}</b> (ID ${item.id}) (Ref #: ${item.ref})</li>`; 
    }).join('');
    const uniqueRefNums = lockedRefNums.sort(); 
    const subjectRefNums = uniqueRefNums.join(', ');
    const subject = `ATTN: Admin Unlock Request - Ref# ${subjectRefNums} for ${sfcRef}`;
    const idsEncoded = encodeURIComponent(personnelIds.join(','));
    const refsEncoded = encodeURIComponent(expandedRefNums.join(',')); // Expanded list
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
        <ul style="list-style-type: 
        none; padding-left: 0;">
            ${requestDetails} </ul>
        <hr style="margin: 10px 0;">
        <h3 style="color: #1e40af;">Admin Action Required:</h3>
      
        <div style="margin-top: 15px;">
            <a href="${unlockUrl}" target="_blank" 
               style="background-color: #10b981; color: white; padding: 10px 20px; text-align: 
        center; text-decoration: none; display: inline-block; border-radius: 5px;
        font-weight: bold;
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

function logAdminUnlockAction(status, sfcRef, personnelIds, lockedRefNums, requesterEmail, adminEmail) 
{
    
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;
    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();
    
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
        let rowIndexToUpdate = -1;
        
        for (let i = values.length - 1; i >= 0; i--) {            
     
            const row = values[i];
    
            const currentStatus = String(row[STATUS_INDEX] || '').trim();
            
            if (String(row[SFC_INDEX]).trim() === sfcRef &&
                String(row[ID_INDEX]).trim() === id &&
                String(row[REF_INDEX]).trim() === targetRef &&
              
                String(row[REQUESTER_INDEX]).trim() === requesterEmail && currentStatus === 'PENDING') {
                rowIndexToUpdate = i + 2; 
                break;
              }
        }
     
        if (rowIndexToUpdate !== -1) 
        {
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            targetRow.getCell(1, ADMIN_EMAIL_INDEX + 1).setValue(adminEmail);
            targetRow.getCell(1, ADMIN_ACTION_TIME_INDEX + 1).setValue(new Date());
            targetRow.getCell(1, STATUS_INDEX + 1).setValue(status);
            updatedCount++;
        }
    });
    Logger.log(`[logAdminUnlockAction] Logged ${status} for ${updatedCount} unlock requests.`);
}

function logUserActionAfterUnlock(sfcRef, employeeChanges, attendanceChanges, userEmail, year, month, shift) { 
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;
    const targetMonthYear = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' }) + year;
    const targetPeriodIdentifier = `${targetMonthYear}-${shift}`;
    Logger.log(`[logUserActionAfterUnlock] Target Plan Period: ${targetPeriodIdentifier}`);
    
    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();
    const SFC_INDEX = 0;
    const ID_INDEX = 1;
    const REF_INDEX = 3;
    const REQUESTER_INDEX = 4;
    const STATUS_INDEX = 8;
    const USER_ACTION_TYPE_INDEX = 9;
    const USER_ACTION_TIME_INDEX = 10;
    
    const MIN_ROSTER_FOR_ENTIRE_EDIT = 5;
    let actionType = 'Edit Personal AP (Info Only)';
    if (attendanceChanges.length > 0) {
        
        const planEmployees = getEmployeeMasterData(sfcRef);
        const modifiedIDs = new Set(attendanceChanges.map(c => cleanPersonnelId(c.personnelId)));
        const allPlanIDs = new Set(planEmployees.map(e => e.id));
        if (employeeChanges.some(c => c.isNew)) {
             actionType = 'Create AP Plan For an OLD AP';
        } 
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

    const modifiedIdsInSave = new Set([
        ...employeeChanges.map(c => cleanPersonnelId(c.id || c.oldPersonnelId)),
        ...attendanceChanges.map(c => cleanPersonnelId(c.personnelId))
    ]);
    let loggedCount = 0;
    const processedKeys = new Set(); 

    for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        const rowId = cleanPersonnelId(row[ID_INDEX]);
        const rowSfc = String(row[SFC_INDEX]).trim();
        const rowStatus = String(row[STATUS_INDEX]).trim();
        const rowActionType = String(row[USER_ACTION_TYPE_INDEX]).trim();
        const rowRequester = String(row[REQUESTER_INDEX]).trim();
        const lockedRef = String(row[REF_INDEX]).trim();
        const isTargetPeriod = lockedRef.includes(targetPeriodIdentifier);
        
        if (rowSfc === sfcRef && 
            modifiedIdsInSave.has(rowId) &&
            rowStatus === 'APPROVED' &&
            rowActionType === '' &&
 
            isTargetPeriod 
           ) {
            
            const rowKey = `${rowId}_${rowSfc}_${rowRequester}`;
            if (processedKeys.has(rowKey)) continue;

            const rowIndexToUpdate = i + 2;
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            targetRow.getCell(1, USER_ACTION_TYPE_INDEX + 1).setValue(actionType);
            targetRow.getCell(1, USER_ACTION_TIME_INDEX + 1).setValue(new Date());
            processedKeys.add(rowKey);
            loggedCount++;
        }
    }
     Logger.log(`[logUserActionAfterUnlock] Logged action type "${actionType}" for ${loggedCount} recently approved unlock requests matching period ${targetPeriodIdentifier}.`);
}

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
    const processedKeys = new Set(); 

    for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        const rowId = String(row[ID_INDEX]).trim();
        const rowSfc = String(row[SFC_INDEX]).trim();
        const rowStatus = String(row[STATUS_INDEX]).trim();
        const rowActionType = String(row[USER_ACTION_TYPE_INDEX]).trim();
        const rowRequester = String(row[REQUESTER_INDEX]).trim();
        if (rowSfc === sfcRef && 
            printedPersonnelIds.includes(rowId) &&
            rowStatus === 'APPROVED' &&
            rowActionType === '' 
           ) {
         
            const rowKey = `${rowId}_${rowSfc}_${rowRequester}`;
            if (processedKeys.has(rowKey)) continue;
            
            const rowIndexToUpdate = i + 2;
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            targetRow.getCell(1, USER_ACTION_TYPE_INDEX + 1).setValue('Re-Print Attendance Plan');
            targetRow.getCell(1, USER_ACTION_TIME_INDEX + 1).setValue(new Date());
            
            processedKeys.add(rowKey);
            loggedCount++;
        }
    }
     Logger.log(`[logUserReprintAction] Logged 'Re-Print Attendance Plan' for ${loggedCount} recently approved unlock requests.`);
}

function getBatchRequestStatus(sfcRef, personnelIds, lockedRefNums, requesterEmail) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return 'PENDING';
    
    const SFC_INDEX = 0;
    const ID_INDEX = 1;
    const REF_INDEX = 3;
    const REQUESTER_INDEX = 4;
    const STATUS_INDEX = 8;
    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();
    
    const incomingKeys = new Set(personnelIds.map((id, index) => `${id}_${lockedRefNums[index]}`));
    const latestStatusMap = {};
    
    for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        const rowId = String(row[ID_INDEX]).trim();
        const rowRef = String(row[REF_INDEX]).trim();
        const rowKey = `${rowId}_${rowRef}`;
        if (String(row[SFC_INDEX]).trim() === sfcRef &&
            String(row[REQUESTER_INDEX]).trim() === requesterEmail &&
            incomingKeys.has(rowKey) &&
            !latestStatusMap[rowKey] 
           ) {
            latestStatusMap[rowKey] = String(row[STATUS_INDEX] || '').trim();
        }
    }
    
    let blockStatus = 'PENDING';
    incomingKeys.forEach(key => {
        const status = latestStatusMap[key] || 'PENDING';     
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
  
  const currentBatchStatus = getBatchRequestStatus(sfcRef, personnelIds, lockedRefNums, requesterEmail);
  if (currentBatchStatus !== 'PENDING') {
      const template = HtmlService.createTemplateFromFile('UnlockStatus');
      template.status = 'INFO';
      template.message = `This unlock request has already been processed as **${currentBatchStatus}** by Admin (${userEmail}). No further action will be taken.`;
      return template.evaluate().setTitle('Request Already Processed');
  }
  
  const summary = `${personnelIds.length} schedules (Ref# ${[...new Set(lockedRefNums)].join(', ')})`;
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
