// --- CONFIGURATION ---
const SPREADSHEET_ID = '1rQnJGqcWcEBjoyAccjYYMOQj7EkIu1ykXTMLGFzzn2I';
const TARGET_SPREADSHEET_ID = '16HS0KIr3xV4iFvEUixWSBGWfAA9VPtTpn5XhoBeZdk4'; 
const CONTRACTS_SHEET_NAME = 'MASTER';

const MASTER_HEADER_ROW = 5;
const PLAN_HEADER_ROW = 6;

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
    return idString.replace(/\D/g, '');
}

function getAttendancePlan(sfcRef, year, month, shift) {
    if (!sfcRef) throw new Error("SFC Ref# is required.");
    const empSheetName = getDynamicSheetName(sfcRef, 'employees');
    const empData = getSheetData(TARGET_SPREADSHEET_ID, empSheetName);
    
    const planSheetName = getDynamicSheetName(sfcRef, 'plan', year, month, shift);
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const planSheet = ss.getSheetByName(planSheetName);
    
    if (!planSheet) return { employees: empData, planMap: {} };
    const HEADER_ROW = PLAN_HEADER_ROW;
    const lastRow = planSheet.getLastRow();
    const numRowsToRead = lastRow - HEADER_ROW;
    const numColumns = planSheet.getLastColumn();

    if (numRowsToRead <= 0 || numColumns < 33) { 
        return { employees: empData, planMap: {} };
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
            area: String(e['Area Posting'] || '').trim(),
        }
    }).filter(e => e.id);
    return { employees, planMap };
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
    saveContractInfo(sfcRef, contractInfo, year, month, shift);
    if (employeeChanges && employeeChanges.length > 0) {
        saveEmployeeInfoBulk(sfcRef, employeeChanges, year, month, shift);
    }
    
    if (attendanceChanges && attendanceChanges.length > 0) {
        saveAttendancePlanBulk(sfcRef, attendanceChanges, year, month, shift);
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
            targetLookupHeader = `Day${dayNumber}`;
        }
        
        const dayColIndex = sanitizedHeadersMap[sanitizeHeader(targetLookupHeader)];
        if (dayColIndex === undefined) {
            Logger.log(`[savePlanBulk] FATAL MISS: Header Lookup '${targetLookupHeader}' failed. Available Sanitized Keys: ${Object.keys(sanitizedHeadersMap).join(' | ')}`);
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
    
    const planSheetName1st = getDynamicSheetName(sfcRef, 'plan', year, month, '1stHalf');
    const planSheet1st = ss.getSheetByName(planSheetName1st);
    const planSheetName2nd = getDynamicSheetName(sfcRef, 'plan', year, month, '2ndHalf'); 
    const planSheet2nd = ss.getSheetByName(planSheetName2nd);
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
    
    const planRowsToAppend1st = []; 
    const planRowsToAppend2nd = [];
    const personnelIdMap = {};
    for (let i = 1; i < values.length; i++) { 
        personnelIdMap[String(values[i][personnelIdIndex] || '').trim()] = i + 1;
    }
    
    changes.forEach((data) => {
        const oldId = String(data.oldPersonnelId || '').trim();
        const newId = String(data.id || '').trim();
        
        if (!data.isNew && personnelIdMap[oldId]) { 
             const sheetRowNumber = personnelIdMap[oldId];
          
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
    if (planRowsToAppend1st.length > 0 && planSheet1st) {
        planSheet1st.getRange(planSheet1st.getLastRow() + 1, 1, planRowsToAppend1st.length, planRowsToAppend1st[0].length).setValues(planRowsToAppend1st);
    }
    
    // 4. Append new rows (Attendance Plan Sheet - 2nd Half)
    if (planRowsToAppend2nd.length > 0 && planSheet2nd) {
        planSheet2nd.getRange(planSheet2nd.getLastRow() + 1, 1, planRowsToAppend2nd.length, planRowsToAppend2nd[0].length).setValues(planRowsToAppend2nd);
    }
}
