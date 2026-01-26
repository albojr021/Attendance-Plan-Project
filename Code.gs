// ============================================================================
// 1. CONFIGURATION & CONSTANTS
// ============================================================================

const SPREADSHEET_ID = '1qheN_KURc-sOKSngpzVxLvfkkc8StzGv-1gMvGJZdsc'; // Source Contracts
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; // Target Database
const FILE_201_ID = '1hk4UX4tBFh2-_udnPrlii5j07G4MT0Sl007ZEBCWOc8'; // 201 File Master
const PDF_FOLDER_ID = '1_CfNZlLDfWW5UBxRDubbDxeN5vZdtNs2';

// Sheet Names
const CONTRACTS_SHEET_NAME = 'RefSeries';
const PLAN_SHEET_NAME = 'AttendancePlan_Consolidated';
const EMPLOYEE_MASTER_SHEET_NAME = 'EmployeeMaster_Consolidated';
const SIGNATORY_MASTER_SHEET = 'SignatoryMaster';
const PRINT_FIELD_MASTER_SHEET = 'PrintFieldMaster';
const LOG_SHEET_NAME = 'PrintLog';
const UNLOCK_LOG_SHEET_NAME = 'UnlockRequestLog';
const SECURITY_PLAN_SHEET_NAME = 'SecurityPlan_Details';
const FILE_201_SHEET_NAME = ['MEG', 'MALL'];
const BLACKLIST_SHEET_NAMES = ['MEG', 'MALL'];

// Headers & Columns
const REFSERIES_HEADER_ROW = 8;
const PLAN_HEADER_ROW = 1;
const PLAN_FIXED_COLUMNS = 18;
const PLAN_MAX_DAYS_IN_HALF = 16;
const FILE_201_ID_COL_INDEX = 0; // Column A
const FILE_201_NAME_COL_INDEX = 1; // Column B
const FILE_201_BLACKLIST_STATUS_COL_INDEX = 2; // Column C

const ADMIN_EMAILS = ['mcdmarketingstorage@megaworld-lifestyle.com'];

const LOG_HEADERS = [
    'Reference #', 'SFC Ref#', 'Plan Sheet Name (N/A)', 'Plan Period Display', 
    'Payor Company', 'Agency', 'Sub Property', 'Service Type',
    'User Email', 'Timestamp', 'Locked Personnel IDs'
];

const UNLOCK_LOG_HEADERS = [
    'SFC Ref#', 'Personnel ID', 'Personnel Name', 'Locked Ref #', 
    'Requesting User', 'Request Timestamp', 'Admin Email', 
    'Admin Action Timestamp', 'Status (APPROVED/REJECTED)', 
    'User Action Type', 'User Action Timestamp'
];

// ============================================================================
// 2. UI SERVING FUNCTIONS
// ============================================================================

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

// ============================================================================
// 3. UTILITY FUNCTIONS
// ============================================================================

function sanitizeHeader(header) {
    if (!header) return '';
    return String(header).replace(/[^A-Za-z00-9#\/]/g, '');
}

function cleanPersonnelId(rawId) {
    let idString = String(rawId || '').trim();
    return idString.replace(/\D/g, '');
}

function getNextSequentialNumber(logSheet, sfcRef) { 
    const logValues = logSheet.getLastRow() > 1 ?
        logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues() : [];
    
    let maxSequentialNumber = 0;

    logValues.forEach(row => {
        const logRefString = String(row[0] || '').trim();
        const parts = logRefString.split('-');

        if (parts.length >= 5) {
            const sequentialPart = parts[parts.length - 2]; 
            const numericPart = parseInt(sequentialPart, 10);

            if (!isNaN(numericPart) && numericPart > maxSequentialNumber) {
                 maxSequentialNumber = numericPart;
            }
        }
    });

    const nextNumber = maxSequentialNumber + 1;

    if (nextNumber <= 9999) {
        return String(nextNumber).padStart(4, '0');
    }
    return String(nextNumber);
}

function getDynamicSheetName(sfcRef, type, year, month, shift) {
    const safeRef = (sfcRef || '').replace(/[\\/?*[]/g, '_');
    if (type === 'employees') {
        return EMPLOYEE_MASTER_SHEET_NAME;
    }
    return `${safeRef} - AttendancePlan`; 
}

function savePrintToPdfAndLog(htmlContent, fileName, sfcRef, year, month, shift, personnelIds) {
  try {
    if (!PDF_FOLDER_ID || PDF_FOLDER_ID === 'GOOGLE_DRIVE_FOLDER_ID') {
      throw new Error("CONFIGURATION ERROR: Set The PDF_FOLDER_ID in Code.gs.");
    }

    // 1. Create PDF Blob
    const blob = Utilities.newBlob(htmlContent, MimeType.HTML, fileName).getAs(MimeType.PDF);
    
    // 2. Save to Drive
    const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl(); 

    // 3. Log the Link to the Sheet
    updateLinkFileColumn(sfcRef, year, month, shift, personnelIds, fileUrl);

    return { success: true, url: fileUrl, message: 'PDF Saved and Linked successfully.' };

  } catch (e) {
    Logger.log(`[savePrintToPdfAndLog] ERROR: ${e.message}`);
    return { success: false, message: `Failed to save PDF: ${e.message}` };
  }
}

function updateLinkFileColumn(sfcRef, year, month, shift, personnelIds, fileUrl) {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    
    if (!planSheet) return;

    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    
    const headers = planSheet.getRange(HEADER_ROW, 1, 1, planSheet.getLastColumn()).getValues()[0];
    
    const sfcIdx = headers.indexOf('CONTRACT #');
    const monthIdx = headers.indexOf('MONTH');
    const yearIdx = headers.indexOf('YEAR');
    const shiftIdx = headers.indexOf('PERIOD / SHIFT');
    const idIdx = headers.indexOf('Personnel ID');
    const linkFileIdx = headers.indexOf('LINK FILE');

    if (linkFileIdx === -1) {
        throw new Error("Column 'LINK FILE' not found. Please add it manually to the sheet or regenerate headers.");
    }

    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);

    // Get Data 
    const dataRange = planSheet.getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, planSheet.getLastColumn());
    const values = dataRange.getValues();
    const updates = [];

    values.forEach((row, rIndex) => {
        const rSfc = String(row[sfcIdx]).trim();
        const rMonth = String(row[monthIdx]).trim();
        const rYear = String(row[yearIdx]).trim();
        const rShift = String(row[shiftIdx]).trim();
        const rId = cleanPersonnelId(row[idIdx]);

        // Check if row matches the criteria and the personnel ID is in the list
        if (rSfc === sfcRef && 
            rMonth === targetMonthShort && 
            rYear === targetYear && 
            rShift === shift && 
            personnelIds.includes(rId)) {
            
            // Push update: rIndex + HEADER_ROW + 1 (dahil 0-based ang array at may header)
            updates.push({
                row: rIndex + HEADER_ROW + 1,
                col: linkFileIdx + 1,
                val: fileUrl
            });
        }
    });

    // Batch Update
    if (updates.length > 0) {
        updates.forEach(u => {
            planSheet.getRange(u.row, u.col).setValue(u.val);
        });
    }
}

// ============================================================================
// 4. SHEET HELPERS (GET/CREATE/ENSURE)
// ============================================================================

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
  } else if (sheetName === PLAN_SHEET_NAME || sheetName === EMPLOYEE_MASTER_SHEET_NAME) {
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
      if (!row.some(cell => String(cell).trim() !== '')) continue;
    }
    
    const item = {};
    cleanHeaders.forEach((headerKey, index) => {
      if (headerKey) item[headerKey] = row[index];
    });
    data.push(item);
  }
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
         Logger.log(`[checkContractSheets] ERROR: Failed to open Spreadsheet. Error: ${e.message}`);
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
    } 
    return empSheet;
}

function getOrCreatePrintFieldMasterSheet(ss) {
    let sheet = ss.getSheetByName(PRINT_FIELD_MASTER_SHEET);
    if (sheet) return sheet;

    try {
        sheet = ss.insertSheet(PRINT_FIELD_MASTER_SHEET);
        const headers = ['SECTION', 'DEPARTMENT', 'REMARKS'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 150);
        sheet.setColumnWidth(2, 150);
        sheet.setColumnWidth(3, 300);
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${PRINT_FIELD_MASTER_SHEET}" already exists`)) {
             return ss.getSheetByName(PRINT_FIELD_MASTER_SHEET);
        }
        throw e;
    }
}

function getOrCreateSignatoryMasterSheet(ss) {
    let sheet = ss.getSheetByName(SIGNATORY_MASTER_SHEET);
    if (sheet) return sheet;

    try {
        sheet = ss.insertSheet(SIGNATORY_MASTER_SHEET);
        const headers = ['Signatory Name', 'Designation'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidth(1, 200);
        sheet.setColumnWidth(2, 150);
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${SIGNATORY_MASTER_SHEET}" already exists`)) {
             return ss.getSheetByName(SIGNATORY_MASTER_SHEET);
        }
        throw e;
    }
}

function getOrCreateLogSheet(ss) {
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (sheet) return sheet;
    
    try {
        sheet = ss.insertSheet(LOG_SHEET_NAME);
        sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, LOG_HEADERS.length, 120); 
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${LOG_SHEET_NAME}" already exists`)) {
             return ss.getSheetByName(LOG_SHEET_NAME);
        }
        throw e;
    }
}

function getOrCreateUnlockRequestLogSheet(ss) {
    let sheet = ss.getSheetByName(UNLOCK_LOG_SHEET_NAME);
    if (sheet) return sheet;
    
    try {
        sheet = ss.insertSheet(UNLOCK_LOG_SHEET_NAME);
        sheet.getRange(1, 1, 1, UNLOCK_LOG_HEADERS.length).setValues([UNLOCK_LOG_HEADERS]);
        sheet.setFrozenRows(1);
        sheet.setColumnWidths(1, UNLOCK_LOG_HEADERS.length, 120); 
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${UNLOCK_LOG_SHEET_NAME}" already exists`)) {
             return ss.getSheetByName(UNLOCK_LOG_SHEET_NAME);
        }
        throw e;
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
            'PERIOD / SHIFT', 'SAVE GROUP', 'SAVE VERSION', 
            'PRINT GROUP', 'Reference #',
            'Personnel ID', 'Personnel Name', 'POSITION', 'AREA POSTING'
        ];
        for (let d = 1; d <= PLAN_MAX_DAYS_IN_HALF; d++) {
            base.push(`DAY${d}`);
        }
        base.push('LINK FILE');
        return base;
    };
    
    if (!planSheet) {
        planSheet = ss.insertSheet(PLAN_SHEET_NAME);
        planSheet.clear();
        const planHeaders = getConsolidatedPlanHeaders();
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, planHeaders.length).setValues([planHeaders]);
        planSheet.setFrozenRows(PLAN_HEADER_ROW); 
        planSheet.getRange(PLAN_HEADER_ROW, 1, 1, PLAN_FIXED_COLUMNS).setNumberFormat('@');
    } 
}

function ensureContractSheets(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required to ensure sheets.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID); 
    getOrCreateConsolidatedEmployeeMasterSheet(ss);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) {
        createContractSheets(sfcRef, year, month, shift);
    }
}

function getOrCreateSecurityPlanSheet(ss) {
    let sheet = ss.getSheetByName(SECURITY_PLAN_SHEET_NAME);
    if (sheet) return sheet;

    try {
        sheet = ss.insertSheet(SECURITY_PLAN_SHEET_NAME);
        // Headers specific for Security Details
        const headers = [
            'CONTRACT #', 'MONTH', 'YEAR', 'PERIOD / SHIFT', 
            'Personnel ID', 'Personnel Name', 
            'Time of Shift', 'Firearm Type', 'Firearm Make', 
            'Firearm Caliber', 'Firearm Serial', 'License Validity'
        ];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);
        // Format columns as text/string to avoid date conversion issues
        sheet.getRange(1, 1, 1, headers.length).setNumberFormat('@'); 
        return sheet;
    } catch (e) {
        if (e.message.includes(`sheet with the name "${SECURITY_PLAN_SHEET_NAME}" already exists`)) {
             return ss.getSheetByName(SECURITY_PLAN_SHEET_NAME);
        }
        throw e;
    }
}

// ============================================================================
// 5. MASTER DATA FETCHING (CONTRACTS, EMPLOYEES, FIELDS)
// ============================================================================

function getContracts() {
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE' || !SPREADSHEET_ID) {
    throw new Error("CONFIGURATION ERROR: Pakipalitan ang Spreadsheet ID.");
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
    const kindOfSfcKey = findKey(c, 'Kind of SFC');
  
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

function fetchSupportingContracts(targetGroupId) {
    if (!targetGroupId) return [];
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONTRACTS_SHEET_NAME);
    if (!sheet) return [];

    const startRow = REFSERIES_HEADER_ROW;
    const lastRow = sheet.getLastRow();
    if (lastRow < startRow) return [];

    const numRows = lastRow - startRow + 1;
    const numCols = sheet.getLastColumn();
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    const values = dataRange.getDisplayValues();
    const headers = values[0];
    
    const findColIndex = (searchName) => headers.findIndex(h => String(h).trim().toLowerCase() === searchName.toLowerCase());
    const colGrpId = findColIndex('Contract Group ID');
    const colKindSfc = findColIndex('Kind of SFC');
    const colRef = findColIndex('Ref #');
    const colBallWith = findColIndex('SFC Ball with?');

    if (colGrpId === -1) return [];

    const relatedContracts = [];
    for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const currentGrpId = String(row[colGrpId] || '').trim();
        if (currentGrpId === String(targetGroupId).trim()) {
            relatedContracts.push({
                grpId: currentGrpId,
                kindSfc: colKindSfc !== -1 ? String(row[colKindSfc] || '') : '',
                refNum: colRef !== -1 ? String(row[colRef] || '') : '',
                ballWith: colBallWith !== -1 ? String(row[colBallWith] || '') : ''
            });
        }
    }
    return relatedContracts;
}

function getPrintFieldMasterData() {
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = getOrCreatePrintFieldMasterSheet(ss);

        if (sheet.getLastRow() < 2) return { sections: [], departments: [], remarks: [] };
        
        const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
        const values = range.getDisplayValues();

        const sections = new Set();
        const departments = new Set();
        const remarks = new Set();

        values.forEach(row => {
            const section = String(row[0] || '').trim().toUpperCase();
            const department = String(row[1] || '').trim().toUpperCase();
            const remark = String(row[2] || '').trim(); 
            
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
    if (!printFields || !printFields.section || !printFields.department) return;
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const sheet = getOrCreatePrintFieldMasterSheet(ss);
    
    const newSection = printFields.section.trim().toUpperCase();
    const newDepartment = printFields.department.trim().toUpperCase();
    const newRemarks = (printFields.remarks || '').trim();

    if (sheet.getLastRow() < 2) {
        sheet.appendRow([newSection, newDepartment, newRemarks]);
        return;
    }

    const allValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
    const existingSections = new Set();
    const existingDepartments = new Set();
    const existingRemarks = new Set();

    allValues.forEach(row => {
        existingSections.add(String(row[0] || '').trim().toUpperCase());
        existingDepartments.add(String(row[1] || '').trim().toUpperCase());
        existingRemarks.add(String(row[2] || '').trim());
    });

    let isSectionNew = !existingSections.has(newSection);
    let isDepartmentNew = !existingDepartments.has(newDepartment);
    let isRemarksNew = !existingRemarks.has(newRemarks);

    if (isSectionNew || isDepartmentNew || isRemarksNew) {
        const newEntry = [
            isSectionNew ? newSection : '',
            isDepartmentNew ? newDepartment : '',
            isRemarksNew ? newRemarks : ''
        ];
        sheet.appendRow(newEntry);
    }
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
    }
}

function get201FileAllPersonnelDetails() {
    if (!FILE_201_ID) return [];
    try {
        const ss = SpreadsheetApp.openById(FILE_201_ID);
        let allData = [];

        FILE_201_SHEET_NAME.forEach(sheetName => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet) return; // Skip kung wala ang sheet

            const START_ROW = 2; 
            const lastRow = sheet.getLastRow();
            const NUM_ROWS = lastRow - START_ROW + 1;
            const NUM_COLS_TO_READ = FILE_201_BLACKLIST_STATUS_COL_INDEX + 1;

            if (NUM_ROWS > 0) {
                const values = sheet.getRange(START_ROW, 1, NUM_ROWS, NUM_COLS_TO_READ).getDisplayValues();
                const sheetData = values.map(row => {
                    const personnelIdRaw = row[FILE_201_ID_COL_INDEX]; 
                    const personnelNameRaw = row[FILE_201_NAME_COL_INDEX];
                    const status = String(row[FILE_201_BLACKLIST_STATUS_COL_INDEX] || '').trim().toUpperCase(); 
                    
                    const cleanId = cleanPersonnelId(personnelIdRaw);
                    const formattedName = String(personnelNameRaw || '').trim().toUpperCase(); 
        
                    const isBlacklisted = status === 'BLACKLISTED';
            
                    if (!cleanId || !formattedName) return null; 
                    return {
                        id: cleanId,
                        name: formattedName,
                        isBlacklisted: isBlacklisted,
                        position: '',
                        area: ''      
                    };
                }).filter(item => item !== null);
                
                allData = allData.concat(sheetData);
            }
        });

        return allData;
    } catch (e) {
        throw new Error(`Failed to access 201 Master File. Error: ${e.message}`);
    }
}

function get201FileMasterData() {
    if (!FILE_201_ID) return [];
    try {
        const allPersonnel = get201FileAllPersonnelDetails();
        return allPersonnel
            .filter(e => !e.isBlacklisted)
            .map(e => ({ id: e.id, name: e.name }));
    } catch (e) {
        throw e;
    }
}

function getBlacklistedEmployeesFrom201() {
    if (!FILE_201_ID) return [];
    try {
        const ss = SpreadsheetApp.openById(FILE_201_ID);
        const blacklistedEmployees = [];

        FILE_201_SHEET_NAME.forEach(sheetName => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet) return;

            const START_ROW = 2;
            const lastRow = sheet.getLastRow();
            const NUM_ROWS = lastRow - START_ROW + 1;
            const NUM_COLS_TO_READ = FILE_201_BLACKLIST_STATUS_COL_INDEX + 1;

            if (NUM_ROWS > 0) {
                const values = sheet.getRange(START_ROW, 1, NUM_ROWS, NUM_COLS_TO_READ).getDisplayValues();

                values.forEach(row => {
                    const status = String(row[FILE_201_BLACKLIST_STATUS_COL_INDEX] || '').trim().toUpperCase();
                    if (status === 'BLACKLISTED') {
                        const personnelIdRaw = row[FILE_201_ID_COL_INDEX]; 
                        const personnelNameRaw = row[FILE_201_NAME_COL_INDEX];
                        const id = cleanPersonnelId(personnelIdRaw);
                        const name = String(personnelNameRaw || '').trim().toUpperCase(); 

                        if (id) {  
                            blacklistedEmployees.push({ id: id, name: name });
                        }
                    }
                });
            }
        });
        
        return blacklistedEmployees;
    } catch (e) {
        throw new Error(`Failed to access 201 Master File for blacklist check. Error: ${e.message}`);
    }
}

function getBlacklistData() { 
    return getBlacklistedEmployeesFrom201();
}

function getSecurityFieldSuggestions() {
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const suggestions = {
        designations: new Set(),
        places: new Set(),
        types: new Set(),
        makes: new Set(),
        calibers: new Set(),
        serials: new Set()
    };

    // 1. Get Designation (Position) and Place (Area) from Employee Master
    const empSheet = ss.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
    if (empSheet && empSheet.getLastRow() > 1) {
        const data = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).getValues();
        const headers = empSheet.getRange(1, 1, 1, empSheet.getLastColumn()).getValues()[0];
        const posIdx = headers.indexOf('Position');
        const areaIdx = headers.indexOf('Area Posting');

        data.forEach(row => {
            if (posIdx > -1 && row[posIdx]) suggestions.designations.add(String(row[posIdx]).trim().toUpperCase());
            if (areaIdx > -1 && row[areaIdx]) suggestions.places.add(String(row[areaIdx]).trim().toUpperCase());
        });
    }

    // 2. Get Firearm Details from Security Plan History
    const secSheet = ss.getSheetByName(SECURITY_PLAN_SHEET_NAME);
    if (secSheet && secSheet.getLastRow() > 1) {
        const data = secSheet.getRange(2, 1, secSheet.getLastRow() - 1, secSheet.getLastColumn()).getValues();
        const headers = secSheet.getRange(1, 1, 1, secSheet.getLastColumn()).getValues()[0];
        
        const typeIdx = headers.indexOf('Firearm Type');
        const makeIdx = headers.indexOf('Firearm Make');
        const calIdx = headers.indexOf('Firearm Caliber');
        const serialIdx = headers.indexOf('Firearm Serial');

        data.forEach(row => {
            if (typeIdx > -1 && row[typeIdx]) suggestions.types.add(String(row[typeIdx]).trim().toUpperCase());
            if (makeIdx > -1 && row[makeIdx]) suggestions.makes.add(String(row[makeIdx]).trim().toUpperCase());
            if (calIdx > -1 && row[calIdx]) suggestions.calibers.add(String(row[calIdx]).trim().toUpperCase());
            if (serialIdx > -1 && row[serialIdx]) suggestions.serials.add(String(row[serialIdx]).trim().toUpperCase());
        });
    }

    // Convert Sets to Arrays and Sort
    return {
        designations: Array.from(suggestions.designations).sort(),
        places: Array.from(suggestions.places).sort(),
        types: Array.from(suggestions.types).sort(),
        makes: Array.from(suggestions.makes).sort(),
        calibers: Array.from(suggestions.calibers).sort(),
        serials: Array.from(suggestions.serials).sort()
    };
}

function getSupplierAddress(agencyName) {
  if (!agencyName) return '';
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('dvSupplierPayee');
    
    if (!sheet) return '';

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return '';

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); 

    const targetAgency = String(agencyName).trim().toUpperCase();
    
    const foundRow = data.find(row => String(row[2]).trim().toUpperCase() === targetAgency);
    
    if (foundRow) {
      return String(foundRow[6] || '').trim(); // Return Address from Col G
    }
    
    return ''; // Return empty if not found
  } catch (e) {
    Logger.log(`[getSupplierAddress] Error: ${e.message}`);
    return '';
  }
}

// ============================================================================
// 6. ATTENDANCE PLAN READ OPERATIONS
// ============================================================================

function getEmployeeMasterDataForUI(sfcRef) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const datalistEmployees = getEmployeeMasterData(sfcRef);
    const blacklisted = getBlacklistData();
    const all201Personnel = get201FileAllPersonnelDetails();
    return {
        datalist: datalistEmployees,
        blacklisted: blacklisted,
        all201Personnel: all201Personnel 
    };
}

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

function checkCrossContractConflict(personnelName, currentSfcRef, year, month, shift) {
    if (!personnelName) return { hasConflict: false };
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    
    if (!planSheet || planSheet.getLastRow() <= PLAN_HEADER_ROW) {
        return { hasConflict: false };
    }

    const targetNameClean = String(personnelName).toUpperCase().replace(/[^A-Z\s.,-]/g, '').trim();
    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);
    
    const lastRow = planSheet.getLastRow();
    const numColumns = planSheet.getLastColumn();
    const headers = planSheet.getRange(PLAN_HEADER_ROW, 1, 1, numColumns).getDisplayValues()[0];
    const sfcIndex = headers.indexOf('CONTRACT #');
    const nameIndex = headers.indexOf('Personnel Name');
    const monthIndex = headers.indexOf('MONTH');
    const yearIndex = headers.indexOf('YEAR');
    const shiftIndex = headers.indexOf('PERIOD / SHIFT');

    if (sfcIndex === -1 || nameIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1) {
        return { hasConflict: false };
    }

    const dataValues = planSheet.getRange(PLAN_HEADER_ROW + 1, 1, lastRow - PLAN_HEADER_ROW, numColumns).getDisplayValues();
    for (let i = 0; i < dataValues.length; i++) {
        const row = dataValues[i];
        const rowSfc = String(row[sfcIndex] || '').trim();
        const rowName = String(row[nameIndex] || '').toUpperCase().replace(/[^A-Z\s.,-]/g, '').trim();
        const rowMonth = String(row[monthIndex] || '').trim();
        const rowYear = String(row[yearIndex] || '').trim();
        const rowShift = String(row[shiftIndex] || '').trim();
        
        if (rowName === targetNameClean && 
            rowMonth === targetMonthShort && 
            rowYear === targetYear && 
            rowShift === shift && 
            rowSfc !== currentSfcRef) {
            
            return { 
                hasConflict: true, 
                conflictDetails: `Already scheduled in Contract ${rowSfc}` 
            };
        }
    }
    return { hasConflict: false };
}

function getEmployeeMasterData(sfcRef) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const all201Data = get201FileMasterData();
    const clean201DataMap = {}; 
    all201Data.forEach(e => {
        clean201DataMap[e.id] = { 
            id: e.id, name: e.name, position: '', area: '' 
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
        
        if (clean201DataMap[id]) delete clean201DataMap[id];
        return { id, name, position, area };
    }).filter(e => e.id);

    Object.values(clean201DataMap).forEach(emp => {
        if (!finalEmployeeList.some(e => e.id === emp.id)) {
             finalEmployeeList.push(emp);
        }
    });

    return finalEmployeeList.map((e) => ({
        id: e.id, name: e.name, position: e.position, area: e.area,
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

function getAttendancePlan(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    ensureContractSheets(sfcRef, year, month, shift);

    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    const lockedIdRefMap = getLockedPersonnelIds(ss, sfcRef, year, month, shift);
    
    // --- FETCH SECURITY DETAILS (FIXED) ---
    let securityDetailsMap = {};
    const securitySheet = ss.getSheetByName(SECURITY_PLAN_SHEET_NAME);
    
    // Check if sheet exists and has data beyond header
    if (securitySheet && securitySheet.getLastRow() > 1) {
        // CRITICAL FIX: Use getDisplayValues() instead of getValues() 
        // This ensures dates like '12/31/2025' come in as strings, preventing UI crashes.
        const secValues = securitySheet.getRange(1, 1, securitySheet.getLastRow(), securitySheet.getLastColumn()).getDisplayValues();
        const secHeaders = secValues[0];
        
        // Map headers
        const sIdx = secHeaders.indexOf('CONTRACT #');
        const mIdx = secHeaders.indexOf('MONTH');
        const yIdx = secHeaders.indexOf('YEAR');
        const shIdx = secHeaders.indexOf('PERIOD / SHIFT');
        const iIdx = secHeaders.indexOf('Personnel ID');
        
        // Data Column Indices
        const timeIdx = secHeaders.indexOf('Time of Shift');
        const typeIdx = secHeaders.indexOf('Firearm Type');
        const makeIdx = secHeaders.indexOf('Firearm Make');
        const calIdx = secHeaders.indexOf('Firearm Caliber');
        const serialIdx = secHeaders.indexOf('Firearm Serial');
        const validIdx = secHeaders.indexOf('License Validity');
        
        const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
        const targetYear = String(year);

        if (sIdx > -1 && iIdx > -1) {
             for (let i = 1; i < secValues.length; i++) {
                 const row = secValues[i];
                 // Check if row matches current context
                 if (String(row[sIdx]) === sfcRef && 
                     String(row[mIdx]) === targetMonthShort && 
                     String(row[yIdx]) === targetYear && 
                     String(row[shIdx]) === shift) {
                     
                     const pId = cleanPersonnelId(row[iIdx]);
                     
                     if (pId) {
                         securityDetailsMap[pId] = {
                             timeOfShift: timeIdx > -1 ? String(row[timeIdx] || '') : '',
                             faType: typeIdx > -1 ? String(row[typeIdx] || '') : '',
                             faMake: makeIdx > -1 ? String(row[makeIdx] || '') : '',
                             faCaliber: calIdx > -1 ? String(row[calIdx] || '') : '',
                             faSerial: serialIdx > -1 ? String(row[serialIdx] || '') : '',
                             licenseValidity: validIdx > -1 ? String(row[validIdx] || '') : ''
                         };
                     }
                 }
             }
        }
    }

    // Return empty structure if Consolidated Plan is missing/empty
    if (!planSheet) return { employees: [], planMap: {}, lockedIds: Object.keys(lockedIdRefMap), lockedIdRefMap: lockedIdRefMap, securityDetails: securityDetailsMap };
    
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();

    if (numRowsToRead <= 0 || numColumns < (PLAN_FIXED_COLUMNS + 3)) { 
        return { employees: [], planMap: {}, lockedIds: Object.keys(lockedIdRefMap), lockedIdRefMap: lockedIdRefMap, securityDetails: securityDetailsMap };
    }

    // Use getDisplayValues here as well for consistency
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
    const employeesDetails = [];
    const startDayOfMonth = shift === '1stHalf' ? 1 : 16;
    const endDayOfMonth = new Date(year, month + 1, 0).getDate();
    const loopLimit = PLAN_MAX_DAYS_IN_HALF;

    latestDataRows.forEach((row, index) => {
        const id = cleanPersonnelId(row[personnelIdIndex]);
        if (id) {
            employeesDetails.push({
                no: 0, id: id, 
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
                    if (status) planMap[key] = status;
                }
            }
        }
    });

    const employees = employeesDetails.map((e, index) => ({
           no: index + 1, id: e.id, name: e.name, position: e.position, area: e.area,
    })).filter(e => e.id);

    const regularEmployees = employees.filter(e => e.position !== 'RELIEVER' || e.area !== 'RELIEVER');
    const relieverEmployees = employees.filter(e => e.position === 'RELIEVER' && e.area === 'RELIEVER');

    // Return merged data
    return { 
        employees: regularEmployees, 
        relieverPersonnelList: relieverEmployees, 
        planMap, 
        lockedIds: Object.keys(lockedIdRefMap), 
        lockedIdRefMap: lockedIdRefMap,
        securityDetails: securityDetailsMap // Pass the security details map to frontend
    };
}

function getPlanDataForPeriod(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);

    if (!planSheet || planSheet.getLastRow() <= PLAN_HEADER_ROW) {
        return { employees: [], planMap: {}, securityDetails: {} };
    }
    
    // 1. GET BASIC ATTENDANCE PLAN DATA
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

    if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1) {
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

    const regularEmployees = employees.filter(e => {
        const position = e.position.toUpperCase();
        const area = e.area.toUpperCase();
        return position !== 'RELIEVER' || area !== 'RELIEVER';
    });

    // 2. GET SECURITY DETAILS
    let securityDetailsMap = {};
    const securitySheet = ss.getSheetByName(SECURITY_PLAN_SHEET_NAME);
    
    if (securitySheet && securitySheet.getLastRow() > 1) {
        const secValues = securitySheet.getRange(1, 1, securitySheet.getLastRow(), securitySheet.getLastColumn()).getDisplayValues();
        const secHeaders = secValues[0];
        
        const sIdx = secHeaders.indexOf('CONTRACT #');
        const mIdx = secHeaders.indexOf('MONTH');
        const yIdx = secHeaders.indexOf('YEAR');
        const shIdx = secHeaders.indexOf('PERIOD / SHIFT');
        const iIdx = secHeaders.indexOf('Personnel ID');
        
        const timeIdx = secHeaders.indexOf('Time of Shift');
        const typeIdx = secHeaders.indexOf('Firearm Type');
        const makeIdx = secHeaders.indexOf('Firearm Make');
        const calIdx = secHeaders.indexOf('Firearm Caliber');
        const serialIdx = secHeaders.indexOf('Firearm Serial');
        const validIdx = secHeaders.indexOf('License Validity');

        if (sIdx > -1 && iIdx > -1) {
             for (let i = 1; i < secValues.length; i++) {
                 const row = secValues[i];
                 if (String(row[sIdx]) === sfcRef && 
                     String(row[mIdx]) === targetMonthShort && 
                     String(row[yIdx]) === targetYear && 
                     String(row[shIdx]) === shift) {
                     
                     const pId = cleanPersonnelId(row[iIdx]);
                     if (pId) {
                         securityDetailsMap[pId] = {
                             timeOfShift: timeIdx > -1 ? String(row[timeIdx] || '') : '',
                             faType: typeIdx > -1 ? String(row[typeIdx] || '') : '',
                             faMake: makeIdx > -1 ? String(row[makeIdx] || '') : '',
                             faCaliber: calIdx > -1 ? String(row[calIdx] || '') : '',
                             faSerial: serialIdx > -1 ? String(row[serialIdx] || '') : '',
                             licenseValidity: validIdx > -1 ? String(row[validIdx] || '') : ''
                         };
                     }
                 }
             }
        }
    }

    // Return combined data
    return { employees: regularEmployees, planMap: {}, securityDetails: securityDetailsMap };
}

// ============================================================================
// 7. SAVING & WRITING OPERATIONS
// ============================================================================

function saveAllData(sfcRef, contractInfo, employeeChanges, relieverChanges, attendanceChanges, securityDetails, year, month, shift, group) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    ensureContractSheets(sfcRef, year, month, shift);
    
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    
    const lockedIdRefMap = getLockedPersonnelIds(ss, sfcRef, year, month, shift);
    const lockedIds = Object.keys(lockedIdRefMap);

    const finalEmployeeChanges = employeeChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.id || change.oldPersonnelId);
        if (lockedIds.includes(idToCheck) && !change.isDeleted) {
            return false;
        }
        return true;
    });

    const finalRelieverChanges = relieverChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.id);
        if (lockedIds.includes(idToCheck)) return false;
        return true;
    });

    const finalAttendanceChanges = attendanceChanges.filter(change => {
        const idToCheck = cleanPersonnelId(change.personnelId);
        if (lockedIds.includes(idToCheck)) return false;
        return true;
    });

    const regularDeletions = finalEmployeeChanges.filter(c => c.isDeleted).map(c => c.oldPersonnelId);
    const relieverDeletions = finalRelieverChanges.filter(c => c.isDeleted).map(c => c.oldPersonnelId);
    const deletionList = regularDeletions.concat(relieverDeletions);

    if (securityDetails && securityDetails.length > 0) {
        const startDayOfMonth = shift === '1stHalf' ? 1 : 16;
        const dayKey = `${year}-${month + 1}-${startDayOfMonth}`; // Day 1 of the shift
        
        securityDetails.forEach(sec => {
            if (sec.timeOfShift && sec.id && !lockedIds.includes(String(sec.id))) {
                const existingChangeIndex = finalAttendanceChanges.findIndex(ac => 
                    String(ac.personnelId) === String(sec.id) && 
                    ac.dayKey === dayKey && 
                    ac.shift === shift
                );

                if (existingChangeIndex > -1) {
                    finalAttendanceChanges[existingChangeIndex].status = sec.timeOfShift;
                } else {
                    finalAttendanceChanges.push({
                        personnelId: sec.id,
                        dayKey: dayKey,
                        shift: shift,
                        status: sec.timeOfShift
                    });
                }
            }
        });
    }

    if ((securityDetails && securityDetails.length > 0) || deletionList.length > 0) {
        saveSecurityPlanBulk(sfcRef, securityDetails, year, month, shift, ss, deletionList);
    }

    const regularEmployeeInfoChanges = finalEmployeeChanges.filter(c => !c.isDeleted);
    
    if (regularEmployeeInfoChanges.length > 0 || deletionList.length > 0) {
        saveEmployeeInfoBulk(sfcRef, regularEmployeeInfoChanges, ss, deletionList);
    }
    
    const newRelieverEntries = finalRelieverChanges.filter(c => !c.isDeleted);

    if (finalAttendanceChanges.length > 0 || deletionList.length > 0 || newRelieverEntries.length > 0 || regularEmployeeInfoChanges.length > 0) {
        saveAttendancePlanBulk(sfcRef, contractInfo, finalAttendanceChanges, newRelieverEntries, regularEmployeeInfoChanges, year, month, shift, group, deletionList, ss);
    }
    
    logUserActionAfterUnlock(sfcRef, finalEmployeeChanges, finalAttendanceChanges, Session.getActiveUser().getEmail(), year, month, shift);
}

function saveSecurityPlanBulk(sfcRef, securityDetails, year, month, shift, ss, deletionList) {
    const sheet = getOrCreateSecurityPlanSheet(ss);
    const lastRow = sheet.getLastRow();
    
    if ((!securityDetails || securityDetails.length === 0) && (!deletionList || deletionList.length === 0)) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Map headers to indices
    const sfcIdx = headers.indexOf('CONTRACT #');
    const monthIdx = headers.indexOf('MONTH');
    const yearIdx = headers.indexOf('YEAR');
    const shiftIdx = headers.indexOf('PERIOD / SHIFT');
    const idIdx = headers.indexOf('Personnel ID');
    
    // Data columns
    const nameIdx = headers.indexOf('Personnel Name');
    const timeIdx = headers.indexOf('Time of Shift');
    const faTypeIdx = headers.indexOf('Firearm Type');
    const faMakeIdx = headers.indexOf('Firearm Make');
    const faCalIdx = headers.indexOf('Firearm Caliber');
    const faSerialIdx = headers.indexOf('Firearm Serial');
    const validIdx = headers.indexOf('License Validity');

    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);

    if (deletionList && deletionList.length > 0 && lastRow > 1) {
        const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        for (let i = data.length - 1; i >= 0; i--) {
            const row = data[i];
            const rSfc = String(row[sfcIdx]).trim();
            const rMonth = String(row[monthIdx]).trim();
            const rYear = String(row[yearIdx]).trim();
            const rShift = String(row[shiftIdx]).trim();
            const rId = String(row[idIdx]).trim();

            if (rSfc === sfcRef && rMonth === targetMonthShort && rYear === targetYear && rShift === shift) {
                if (deletionList.some(delId => String(delId).trim() === rId)) {
                    sheet.deleteRow(i + 2); // +2 dahil sa header at 0-based index
                }
            }
        }
    }

    if (!securityDetails || securityDetails.length === 0) return;

    // Refresh lastRow after deletion
    const currentLastRow = sheet.getLastRow();
    let existingData = [];
    if (currentLastRow > 1) {
        existingData = sheet.getRange(2, 1, currentLastRow - 1, sheet.getLastColumn()).getValues();
    }

    const rowMap = new Map(); 
    existingData.forEach((row, index) => {
        const rSfc = String(row[sfcIdx]).trim();
        const rMonth = String(row[monthIdx]).trim();
        const rYear = String(row[yearIdx]).trim();
        const rShift = String(row[shiftIdx]).trim();
        const rId = String(row[idIdx]).trim();

        if (rSfc === sfcRef && rMonth === targetMonthShort && rYear === targetYear && rShift === shift) {
             rowMap.set(rId, index);
        }
    });

    const rowsToAppend = [];
    const updates = []; 

    securityDetails.forEach(detail => {
        const id = String(detail.id).trim();
        if (!id) return;
        
        if (deletionList && deletionList.some(delId => String(delId).trim() === id)) return;

        const valName = String(detail.name || '').toUpperCase().trim();
        const valTime = String(detail.timeOfShift || '').trim();
        const valType = String(detail.faType || '').toUpperCase().trim();
        const valMake = String(detail.faMake || '').toUpperCase().trim();
        const valCal = String(detail.faCaliber || '').toUpperCase().trim();
        const valSerial = String(detail.faSerial || '').toUpperCase().trim();
        const valValid = String(detail.licenseValidity || '').trim();

        if (rowMap.has(id)) {
            // Update Existing Row
            const rowIndex = rowMap.get(id) + 2; 
            updates.push({ r: rowIndex, c: nameIdx + 1, v: valName });
            updates.push({ r: rowIndex, c: timeIdx + 1, v: valTime });
            updates.push({ r: rowIndex, c: faTypeIdx + 1, v: valType });
            updates.push({ r: rowIndex, c: faMakeIdx + 1, v: valMake });
            updates.push({ r: rowIndex, c: faCalIdx + 1, v: valCal });
            updates.push({ r: rowIndex, c: faSerialIdx + 1, v: valSerial });
            updates.push({ r: rowIndex, c: validIdx + 1, v: valValid });
        } else {
            // Append New Row
            const newRow = new Array(headers.length).fill('');
            newRow[sfcIdx] = sfcRef;
            newRow[monthIdx] = targetMonthShort;
            newRow[yearIdx] = targetYear;
            newRow[shiftIdx] = shift;
            newRow[idIdx] = id;
            newRow[nameIdx] = valName;
            newRow[timeIdx] = valTime;
            newRow[faTypeIdx] = valType;
            newRow[faMakeIdx] = valMake;
            newRow[faCalIdx] = valCal;
            newRow[faSerialIdx] = valSerial;
            newRow[validIdx] = valValid;
            rowsToAppend.push(newRow);
        }
    });

    // Apply Updates
    if (updates.length > 0) {
        updates.forEach(u => {
            sheet.getRange(u.r, u.c).setValue(u.v);
        });
    }

    // Apply Appends
    if (rowsToAppend.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
    }
}

function saveAttendancePlanBulk(sfcRef, contractInfo, changes, relieverChanges, regularEmployeeInfoChanges, year, month, shift, group, deletionList, ss) {
    const planSheet = ss.getSheetByName(PLAN_SHEET_NAME);
    if (!planSheet) throw new Error(`AttendancePlan Sheet for ${PLAN_SHEET_NAME} not found.`);
    
    const HEADER_ROW = PLAN_HEADER_ROW;
    
    const range = planSheet.getDataRange();
    let values = range.getDisplayValues();
    
    if (values.length < HEADER_ROW) return;
    
    const headers = values[HEADER_ROW - 1]; 
    
    // Kunin ang mga index
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
    const numColumns = headers.length;

    if (sfcRefIndex === -1 || shiftIndex === -1 || personnelIdIndex === -1 || day1Index === -1) {
        throw new Error("Missing critical column in Consolidated Plan sheet.");
    }

    const targetMonthShort = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' });
    const targetYear = String(year);

    const retainedRows = [];
    const latestVersionMap = {}; 

    for (let i = HEADER_ROW; i < values.length; i++) {
        const row = values[i];
        const currentSfc = String(row[sfcRefIndex] || '').trim();
        const currentMonth = String(row[monthIndex] || '').trim();
        const currentYear = String(row[yearIndex] || '').trim();
        const currentShift = String(row[shiftIndex] || '').trim();
        const id = cleanPersonnelId(row[personnelIdIndex]);

        const isTargetRow = (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift);
        
        if (isTargetRow && deletionList.includes(id)) {
            continue; 
        }

        if (isTargetRow && id) {
             const saveVersionString = String(row[saveVersionIndex] || '').trim(); 
             const versionParts = saveVersionString.split('-');
             const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
             
             const existingRow = latestVersionMap[id];
             if (!existingRow || version > (parseFloat(existingRow[saveVersionIndex].split('-').pop()) || 0)) {
                 latestVersionMap[id] = row;
             }
             
             const infoChange = regularEmployeeInfoChanges.find(c => c.id === id);
             if (infoChange) {
                 if (infoChange.name) row[nameIndex] = infoChange.name;
                 if (infoChange.position) row[positionIndex] = infoChange.position;
                 if (infoChange.area) row[areaIndex] = infoChange.area;
             }
        }

        retainedRows.push(row);
    }

    const rowsToAppend = [];
    
    const masterEmployeeMap = getEmployeeMasterData(sfcRef).reduce((map, emp) => { 
        map[emp.id] = { name: emp.name, position: emp.position, area: emp.area };
        return map;
    }, {});

    relieverChanges.forEach(reliever => {
        const personnelId = reliever.id;

        if (latestVersionMap[personnelId]) return; 

        const newRow = Array(numColumns).fill('');
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
        newRow[personnelIdIndex] = personnelId;
        newRow[nameIndex] = reliever.name;
        newRow[positionIndex] = 'RELIEVER';
        newRow[areaIndex] = 'RELIEVER';
        newRow[saveVersionIndex] = `${sfcRef}-${targetMonthShort}${targetYear}-${shift}-${group}-1.0`;

        rowsToAppend.push(newRow);
        latestVersionMap[personnelId] = newRow; // Update map
    });

    const changesByRow = changes.reduce((acc, change) => {
        const key = change.personnelId;
        if (!acc[key]) acc[key] = [];
        acc[key].push(change);
        return acc;
    }, {});

    Object.keys(changesByRow).forEach(personnelId => {
        const dailyChanges = changesByRow[personnelId];
        const latestVersionRow = latestVersionMap[personnelId];
        let newRow;
        let currentVersion = 0;
        let nextGroupToUse = group;

        const empDetails = masterEmployeeMap[personnelId] || { 
            name: (latestVersionRow ? latestVersionRow[nameIndex] : 'N/A'), 
            position: (latestVersionRow ? latestVersionRow[positionIndex] : ''), 
            area: (latestVersionRow ? latestVersionRow[areaIndex] : '')
        };

        if (!latestVersionRow) {
            // New entry na wala pa sa sheet (regular employee added via copy or new input)
            newRow = Array(numColumns).fill('');
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
            newRow[personnelIdIndex] = personnelId;
            newRow[nameIndex] = empDetails.name;
            newRow[positionIndex] = empDetails.position;
            newRow[areaIndex] = empDetails.area;
            currentVersion = 0;
        } else {
            // Existing row, copy it
            newRow = [...latestVersionRow];
            const versionString = String(latestVersionRow[saveVersionIndex] || '').split('-').pop();
            currentVersion = parseFloat(versionString) || 0;
            
            // Clear reference for new version
            newRow[referenceIndex] = '';
            newRow[printGroupIndex] = '';
            
            // Update Group logic
            const oldGroup = latestVersionRow[saveGroupIndex];
            if (oldGroup && String(oldGroup).trim().toUpperCase() !== String(group).trim().toUpperCase()) {
                nextGroupToUse = oldGroup;
            }
            newRow[saveGroupIndex] = nextGroupToUse;

            // Update info fields just in case
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
            if (shift === '1stHalf') dayColumnNumber = dayNumber;
            else dayColumnNumber = dayNumber - 15;

            if (dayColumnNumber >= 1 && dayColumnNumber <= PLAN_MAX_DAYS_IN_HALF) {
                const dayColIndex = day1Index + dayColumnNumber - 1;
                if (dayColIndex < numColumns) {
                    const oldStatus = String((latestVersionRow || newRow)[dayColIndex] || '').trim();
                    if (oldStatus !== newStatus) {
                        isRowChanged = true;
                        newRow[dayColIndex] = newStatus;
                    }
                }
            }
        });

        if (isRowChanged || !latestVersionRow) {
            const nextVersion = (currentVersion + 1).toFixed(1);
            newRow[saveVersionIndex] = `${sfcRef}-${targetMonthShort}${targetYear}-${shift}-${nextGroupToUse}-${nextVersion}`;
            rowsToAppend.push(newRow);
        }
    });

    const topRows = values.slice(0, HEADER_ROW); 
    
    const finalData = [...topRows, ...retainedRows, ...rowsToAppend];
    
    planSheet.clearContents();
    
    if (finalData.length > 0) {
        planSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
        // Re-apply number formatting for fixed columns to ensure IDs are strings
        planSheet.getRange(HEADER_ROW, 1, finalData.length - HEADER_ROW + 1, PLAN_FIXED_COLUMNS).setNumberFormat('@');
    }
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
      
        if (currentSfc === sfcRef && currentMonth === targetMonthShort && currentYear === targetYear && currentShift === shift && printedPersonnelIds.includes(id)) { 
            const saveVersionString = String(row[saveVersionIndex] || '').trim(); 
            const versionParts = saveVersionString.split('-');
            const version = parseFloat(versionParts[versionParts.length - 1]) || 0;
            const existingEntry = latestVersionMap[id];
      
            if (!existingEntry || version > existingEntry.version) {
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
                row: entry.sheetRowNumber, col: referenceIndex + 1, value: refNum
            });
        }
        if (printGroupIndex !== -1) {
            rangesToUpdate.push({
                row: entry.sheetRowNumber, col: printGroupIndex + 1, value: printGroup
            });
         }
    });
    
    if (rangesToUpdate.length > 0) {
        planSheet.setFrozenRows(0);
        rangesToUpdate.forEach(update => {
             planSheet.getRange(update.row, update.col).setNumberFormat('@').setValue(update.value);
        });
        planSheet.setFrozenRows(HEADER_ROW);
    }
}

function saveEmployeeInfoBulk(sfcRef, changes, ss, deletionList) {
    const empSheet = getOrCreateConsolidatedEmployeeMasterSheet(ss);
    let numRows = empSheet.getLastRow();
    
    if (deletionList && deletionList.length > 0 && numRows > 1) {
        const values = empSheet.getDataRange().getValues();
        const headers = values[0];
        const contractRefIndex = headers.indexOf('CONTRACT #');
        const personnelIdIndex = headers.indexOf('Personnel ID');

        // Loop backwards
        for (let i = values.length - 1; i >= 1; i--) { // Start from last row down to 1 (skip header)
            const row = values[i];
            const rowSfc = String(row[contractRefIndex] || '').trim();
            const rowId = String(row[personnelIdIndex] || '').trim();

            if (rowSfc === sfcRef) {
                // Check if ID is in deletion list
                if (deletionList.some(delId => String(delId).trim() === rowId)) {
                    empSheet.deleteRow(i + 1); // +1 dahil 0-based ang array index pero 1-based ang sheet row
                }
            }
        }
        // Refresh numRows after deletion
        numRows = empSheet.getLastRow();
    }

    if (!changes || changes.length === 0) return;

    const values = empSheet.getDataRange().getValues();
    const headers = values[0];
    const contractRefIndex = headers.indexOf('CONTRACT #');
    const personnelIdIndex = headers.indexOf('Personnel ID');
    const nameIndex = headers.indexOf('Personnel Name');
    const positionIndex = headers.indexOf('Position');
    const areaIndex = headers.indexOf('Area Posting');
    
    const existingIds = new Set();
    for (let i = 1; i < values.length; i++) {
        const rowSfc = String(values[i][contractRefIndex] || '').trim();
        const rowId = String(values[i][personnelIdIndex] || '').trim();
        if (rowSfc === sfcRef) {
            existingIds.add(rowId);
        }
    }

    const rowsToAppend = [];
    changes.forEach((data) => {
        if (data.isDeleted) return; 

        const newId = String(data.id || '').trim();
        
        if (newId && !existingIds.has(newId)) {
            const newRow = [];
            for(let k=0; k<headers.length; k++) newRow.push('');

            newRow[contractRefIndex] = sfcRef;
            newRow[personnelIdIndex] = data.id;
            newRow[nameIndex] = data.name;
            newRow[positionIndex] = data.position;
            newRow[areaIndex] = data.area;
            
            rowsToAppend.push(newRow);
            existingIds.add(newId); 
        }
    });

    if (rowsToAppend.length > 0) {
        empSheet.getRange(numRows + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }
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
        
        if (sfcRefIndex === -1 || monthIndex === -1 || yearIndex === -1 || shiftIndex === -1 || saveGroupIndex === -1) return "S1"; 

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
        return `S${maxGroupNumber + 1}`; 
    } catch (e) {
        return "S1";
    }
}

// ============================================================================
// 8. PRINTING & LOGGING OPERATIONS
// ============================================================================

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

        if (currentSfc !== sfcRef || currentPeriodDisplay !== dateRange) return; 
        
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
        if (currentSfc !== sfcRef || currentPeriodDisplay !== dateRange) continue;

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
        } else if (unlockedSequentialNumbers.size === 1) {
            nextPrintGroupNumeric = maxGroupNumber + 1;
            finalSequentialNumber = foundUnlockedBaseRefSequential;
        } else {
            nextPrintGroupNumeric = 1;
            finalSequentialNumber = getNextSequentialNumber(logSheet, sfcRef); 
        }
        
        const finalPrintGroup = `P${nextPrintGroupNumeric}`;
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
        return { refNum: finalPrintReference, printGroup: finalPrintGroup };
    } catch (e) {
        throw new Error(`Failed to generate print reference string. Error: ${e.message}`);
    }
}

function recordPrintLogEntry(refNum, printGroup, subProperty, printFields, signatories, sfcRef, contractInfo, year, month, shift, printedPersonnelIds) { 
    if (!refNum) return;
    try {
        const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const logSheet = getOrCreateLogSheet(ss);
        
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
            refNum, sfcRef, planSheetName, dateRange, contractInfo.payor, contractInfo.agency,
            subProperty, contractInfo.serviceType, Session.getActiveUser().getEmail(), 
            new Date(), printedPersonnelIds.join(',') 
        ];
        const lastLoggedRow = logSheet.getLastRow();
        const newRow = lastLoggedRow + 1;
        const LOCKED_IDS_COL = LOG_HEADERS.length;
        const logEntryRange = logSheet.getRange(newRow, 1, 1, LOG_HEADERS.length);
        logEntryRange.getCell(1, 1).setNumberFormat('@');
        logEntryRange.getCell(1, LOCKED_IDS_COL).setNumberFormat('@');
        logEntryRange.setValues([logEntry]);
        logSheet.getRange(newRow, 1, 1, LOG_HEADERS.length).setHorizontalAlignment('left');
    } catch (e) {
        Logger.log(`[recordPrintLogEntry] FATAL ERROR: ${e.message}`);
    }
}

// ============================================================================
// 9. ADMIN & UNLOCK OPERATIONS
// ============================================================================

function sendRequesterNotification(status, personnelIds, lockedRefNums, personnelNames, requesterEmail) {
  if (requesterEmail === 'UNKNOWN_REQUESTER' || !requesterEmail) return;
  const totalCount = personnelIds.length;
  const uniqueRefNums = [...new Set(lockedRefNums)].sort();
  const subject = `Unlock Request Status: ${status} for ${totalCount} Personnel Schedules (Ref# ${uniqueRefNums.join(', ')})`;
  
  const combinedRequests = personnelIds.map((id, index) => ({
    id: id, ref: lockedRefNums[index], name: personnelNames[index] 
  }));
  combinedRequests.sort((a, b) => a.ref.localeCompare(b.ref)); 

  const idList = combinedRequests.map(item => 
    `<li><b>${item.name}</b> (ID ${item.id}) (Ref #: ${item.ref})</li>` 
  ).join('');
  
  let body = '';
  if (status === 'APPROVED') {
    body = `
      Good news!<br>
      Your request to unlock the following ${totalCount} schedules has been **APPROVED** by the Admin.<br>
      <ul style="list-style-type: none; padding-left: 0; font-weight: bold;">${idList}</ul>
      <br>You may now return to the Attendance Plan Monitor app and refresh your browser to edit the schedules.<br>
      ---<br>This notification confirms the lock is removed.
    `;
  } else if (status === 'REJECTED') {
    body = `
      Your request to unlock the following ${totalCount} schedules has been **REJECTED** by the Admin.<br>
      <ul style="list-style-type: none; padding-left: 0; font-weight: bold;">${idList}</ul>
      <br>The print locks remain active, and the schedules cannot be edited at this time.<br>
      Please contact your Admin for details.<br>
      ---<br>This is an automated notification.
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
  } catch (e) {
    Logger.log(`[sendRequesterNotification] Failed to send status email: ${e.message}`);
  }
}

function unlockPersonnelIds(sfcRef, year, month, shift, personnelIdsToUnlock) {
    const userEmail = Session.getActiveUser().getEmail();
    if (!ADMIN_EMAILS.includes(userEmail)) {
      throw new Error("AUTHORIZATION ERROR: Only admin users can unlock printed schedules.");
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
      
      if (currentSfc !== sfcRef || currentPeriodDisplay !== targetPeriodDisplay) return; 
      
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
        targetRange.setNumberFormat('@').setValue(update.value);
    });
}

function requestUnlockEmailNotification(sfcRef, year, month, shift, lockedRefNums) { 
    if (!lockedRefNums || lockedRefNums.length === 0) {
        return { success: false, message: 'ERROR: No Reference Numbers selected for unlock.' };
    }

    const combinedRequests = getAllLockedIdsByRefs(sfcRef, year, month, shift, lockedRefNums);
    if (combinedRequests.length === 0) {
        return { success: false, message: 'ERROR: Could not find any locked schedules.' };
    }

    const requestingUserEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const unlockLogSheet = getOrCreateUnlockRequestLogSheet(ss);
    combinedRequests.sort((a, b) => a.ref.localeCompare(b.ref));
    const logEntries = [];
    combinedRequests.forEach(item => {
        logEntries.push([
            sfcRef, item.id, item.name, item.ref, requestingUserEmail, new Date(),
            '', '', 'PENDING', '', '' 
        ]);
    });
    if (logEntries.length > 0) {
        unlockLogSheet.getRange(unlockLogSheet.getLastRow() + 1, 1, logEntries.length, UNLOCK_LOG_HEADERS.length).setValues(logEntries);
    }

    const personnelIds = combinedRequests.map(item => item.id);
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
    const refsEncoded = encodeURIComponent(expandedRefNums.join(','));
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
        <ul style="list-style-type: none; padding-left: 0;">${requestDetails}</ul>
        <hr style="margin: 10px 0;">
        <h3 style="color: #1e40af;">Admin Action Required:</h3>
        <div style="margin-top: 15px;">
            <a href="${unlockUrl}" target="_blank" style="background-color: #10b981; color: white; padding: 10px 20px; text-decoration: none; display: inline-block; border-radius: 5px; font-weight: bold; margin-right: 10px;">
               ✅ APPROVE & UNLOCK ALL
            </a>
            <a href="${rejectUrl}" target="_blank" style="background-color: #f59e0b; color: white; padding: 10px 20px; text-decoration: none; display: inline-block; border-radius: 5px; font-weight: bold;">
               ❌ REJECT (Log Only)
            </a>
        </div>
        <p style="margin-top: 20px; font-size: 12px; color: #6b7280;">Login as Admin required.</p>
    `;
    
    try {
        MailApp.sendEmail({ to: adminEmails, subject: subject, htmlBody: htmlBody, name: 'Attendance Plan Monitor' });
        return { success: true, message: `Unlock request sent to Admin(s): ${adminEmails} for ${personnelIds.length} IDs.` }; 
    } catch (e) {
        return { success: false, message: `WARNING: Failed to send request email. Error: ${e.message}` };
    }
}

function logAdminUnlockAction(status, sfcRef, personnelIds, lockedRefNums, requesterEmail, adminEmail) {
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
     
        if (rowIndexToUpdate !== -1) {
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            targetRow.getCell(1, ADMIN_EMAIL_INDEX + 1).setValue(adminEmail);
            targetRow.getCell(1, ADMIN_ACTION_TIME_INDEX + 1).setValue(new Date());
            targetRow.getCell(1, STATUS_INDEX + 1).setValue(status);
        }
    });
}

function logUserActionAfterUnlock(sfcRef, employeeChanges, attendanceChanges, userEmail, year, month, shift) { 
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const logSheet = getOrCreateUnlockRequestLogSheet(ss);
    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) return;
    const targetMonthYear = new Date(year, month, 1).toLocaleString('en-US', { month: 'short' }) + year;
    const targetPeriodIdentifier = `${targetMonthYear}-${shift}`;
    const values = logSheet.getRange(2, 1, lastRow - 1, UNLOCK_LOG_HEADERS.length).getValues();
    
    const SFC_INDEX = 0;
    const ID_INDEX = 1;
    const REQUESTER_INDEX = 4;
    const STATUS_INDEX = 8;
    const USER_ACTION_TYPE_INDEX = 9;
    const USER_ACTION_TIME_INDEX = 10;
    const REF_INDEX = 3;
    
    const MIN_ROSTER_FOR_ENTIRE_EDIT = 5;
    let actionType = 'Edit Personal AP (Info Only)';
    if (attendanceChanges.length > 0) {
        const planEmployees = getEmployeeMasterData(sfcRef);
        const modifiedIDs = new Set(attendanceChanges.map(c => cleanPersonnelId(c.personnelId)));
        const allPlanIDs = new Set(planEmployees.map(e => e.id));
        if (employeeChanges.some(c => c.isNew)) {
             actionType = 'Create AP Plan For an OLD AP';
        } else if (allPlanIDs.size >= MIN_ROSTER_FOR_ENTIRE_EDIT && modifiedIDs.size / allPlanIDs.size > 0.6) {
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
        
        if (rowSfc === sfcRef && modifiedIdsInSave.has(rowId) && rowStatus === 'APPROVED' && rowActionType === '' && isTargetPeriod) {
            const rowKey = `${rowId}_${rowSfc}_${rowRequester}`;
            if (processedKeys.has(rowKey)) continue;

            const rowIndexToUpdate = i + 2;
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            targetRow.getCell(1, USER_ACTION_TYPE_INDEX + 1).setValue(actionType);
            targetRow.getCell(1, USER_ACTION_TIME_INDEX + 1).setValue(new Date());
            processedKeys.add(rowKey);
        }
    }
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
    const processedKeys = new Set(); 

    for (let i = values.length - 1; i >= 0; i--) {
        const row = values[i];
        const rowId = String(row[ID_INDEX]).trim();
        const rowSfc = String(row[SFC_INDEX]).trim();
        const rowStatus = String(row[STATUS_INDEX]).trim();
        const rowActionType = String(row[USER_ACTION_TYPE_INDEX]).trim();
        const rowRequester = String(row[REQUESTER_INDEX]).trim();

        if (rowSfc === sfcRef && printedPersonnelIds.includes(rowId) && rowStatus === 'APPROVED' && rowActionType === '') {
            const rowKey = `${rowId}_${rowSfc}_${rowRequester}`;
            if (processedKeys.has(rowKey)) continue;
            
            const rowIndexToUpdate = i + 2;
            const targetRow = logSheet.getRange(rowIndexToUpdate, 1, 1, UNLOCK_LOG_HEADERS.length);
            targetRow.getCell(1, USER_ACTION_TYPE_INDEX + 1).setValue('Re-Print Attendance Plan');
            targetRow.getCell(1, USER_ACTION_TIME_INDEX + 1).setValue(new Date());
            processedKeys.add(rowKey);
        }
    }
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
            incomingKeys.has(rowKey) && !latestStatusMap[rowKey]) {
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
     return HtmlService.createHtmlOutput('<h1>INVALID REQUEST</h1>');
  }

  const userEmail = Session.getActiveUser().getEmail();
  if (!ADMIN_EMAILS.includes(userEmail)) {
    return HtmlService.createHtmlOutput('<h1>AUTHORIZATION FAILED</h1>');
  }
  
  const currentBatchStatus = getBatchRequestStatus(sfcRef, personnelIds, lockedRefNums, requesterEmail);
  if (currentBatchStatus !== 'PENDING') {
      const template = HtmlService.createTemplateFromFile('UnlockStatus');
      template.status = 'INFO';
      template.message = `This unlock request has already been processed as **${currentBatchStatus}**.`;
      return template.evaluate().setTitle('Request Already Processed');
  }
  
  const summary = `${personnelIds.length} schedules (Ref# ${[...new Set(lockedRefNums)].join(', ')})`;
  if (params.action === 'reject_info') {
      sendRequesterNotification('REJECTED', personnelIds, lockedRefNums, personnelNames, requesterEmail);
      logAdminUnlockAction('REJECTED', sfcRef, personnelIds, lockedRefNums, requesterEmail, userEmail);
      const template = HtmlService.createTemplateFromFile('UnlockStatus');
      template.status = 'INFO';
      template.message = `Admin (${userEmail}) REJECTED the request. Locks remain active.`;
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
        template.message = `Successfully unlocked ${summary}.`;
        return template.evaluate().setTitle('Unlock Status');
      } catch (e) {
        const template = HtmlService.createTemplateFromFile('UnlockStatus');
        template.status = 'ERROR';
        template.message = `Failed to unlock. Error: ${e.message}`;
        return template.evaluate().setTitle('Unlock Status');
      }
  }
  return HtmlService.createHtmlOutput('<h1>Invalid Action</h1>');
}
