/**
 * Syncs data from the source sheet to 'Sheet1' of the active spreadsheet.
 * 100% FORMATTING SAFE: Uses targeted cell updates. Only touches blank cells 
 * and appends new rows at the bottom. Never overwrites formulas or formatting.
 */
function syncDataFromSource() {
  const SOURCE_SS_ID = '1WDCRPlKBeKPCOlzD9EVlM-imLTYyXFp3oGwP99CVvV8';
  const SOURCE_SHEET_NAME = 'OS to RM Handover'; 
  const TARGET_SHEET_NAME = 'Sheet1';
  
  // Columns managed in the Master Sheet that the sync should completely ignore
  const EXCLUDED_COL_LETTERS = ['D', 'M', 'N', 'O', 'R', 'S', 'T']; 

  const RM_IDENTIFIERS = ['zaion.abrar', 'sanjida.ahmed', 'sadia.akter'];
  const COL_B_FILTER = 'nuhash.ahmed';

  const targetSs = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = targetSs.getSheetByName(TARGET_SHEET_NAME);
  const sourceSs = SpreadsheetApp.openById(SOURCE_SS_ID);
  const sourceSheet = sourceSs.getSheetByName(SOURCE_SHEET_NAME);
  
  if (!targetSheet || !sourceSheet) return;

  // Get Source Data
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];
  const sourceRows = sourceData.slice(1);
  const safeSourceHeaders = sourceHeaders.map(h => h.toString().trim().toLowerCase());
  
  const rmEmailIdx = safeSourceHeaders.indexOf('rm email');
  const partnerEmailIdx = safeSourceHeaders.indexOf('partner email');

  if (rmEmailIdx === -1 || partnerEmailIdx === -1) return;

  // Get Target Data
  const targetData = targetSheet.getDataRange().getValues();
  const targetHeaders = targetData[0];
  const safeTargetHeaders = targetHeaders.map(h => h.toString().trim().toLowerCase());
  const targetPartnerEmailIdx = safeTargetHeaders.indexOf('partner email');

  if (targetPartnerEmailIdx === -1) return;

  // Map columns (Source -> Target)
  function colLetterToIndex(l) { return l.split('').reduce((r, a) => r * 26 + parseInt(a, 36) - 9, 0) - 1; }
  const excludedIndices = EXCLUDED_COL_LETTERS.map(colLetterToIndex);
  
  const columnMapping = [];
  safeSourceHeaders.forEach((sH, sIdx) => {
    const tIdx = safeTargetHeaders.indexOf(sH);
    if (tIdx !== -1 && sH !== '' && !excludedIndices.includes(tIdx)) {
      columnMapping.push({ sIdx: sIdx, tIdx: tIdx });
    }
  });

  // Map existing emails to their row index
  const targetEmailMap = {};
  for (let i = 1; i < targetData.length; i++) {
    const email = (targetData[i][targetPartnerEmailIdx] || '').toString().trim().toLowerCase();
    if (email) targetEmailMap[email] = i; 
  }

  const newRowsToAppend = [];

  // Process Source Rows
  sourceRows.forEach(row => {
    const colB = (row[1] || '').toString().toLowerCase();
    const rmVal = (row[rmEmailIdx] || '').toString().toLowerCase();
    
    if (colB.includes(COL_B_FILTER) && RM_IDENTIFIERS.some(id => rmVal.includes(id))) {
      const email = (row[partnerEmailIdx] || '').toString().trim().toLowerCase();
      if (!email) return;

      const targetRowIdx = targetEmailMap[email];

      if (targetRowIdx !== undefined) {
        // TARGETED UPDATE: Row already exists.
        columnMapping.forEach(map => {
          const currentVal = targetData[targetRowIdx][map.tIdx];
          const sourceVal = row[map.sIdx];
          
          // Only take action if the cell in the destination sheet is completely blank
          if (currentVal === "" && sourceVal !== "") {
            
            // Get the exact single cell (Row and Col are 1-based in Apps Script)
            const cell = targetSheet.getRange(targetRowIdx + 1, map.tIdx + 1);
            
            // Double-check that it doesn't contain a formula before writing
            if (cell.getFormula() === "") {
              cell.setValue(sourceVal); // Updates just this one cell, preserving all formatting!
              targetData[targetRowIdx][map.tIdx] = sourceVal; // Update memory to prevent duplicates
            }
          }
        });
      } else {
        // NEW PARTNER: Prepare a new row to be added at the bottom
        const newRow = new Array(targetHeaders.length).fill('');
        columnMapping.forEach(map => { newRow[map.tIdx] = row[map.sIdx]; });
        newRowsToAppend.push(newRow);
        
        // Add to map so we don't append them twice if they appear twice in the source
        targetEmailMap[email] = targetData.length + newRowsToAppend.length - 1;
      }
    }
  });

  // Append all brand-new rows at the bottom of the sheet at once
  if (newRowsToAppend.length > 0) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRowsToAppend.length, targetHeaders.length).setValues(newRowsToAppend);
  }
}
