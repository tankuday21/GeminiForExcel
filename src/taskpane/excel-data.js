/**
 * Excel Data Module
 * Handles Excel data access and context building
 */

/* global Excel */

// ============================================================================
// Column Letter Helpers
// ============================================================================

/**
 * Converts a zero-based column index to Excel column letter(s)
 * Supports multi-letter columns (A, Z, AA, AB, etc.)
 * @param {number} index - Zero-based column index
 * @returns {string} Column letter(s)
 */
function colIndexToLetter(index) {
    let letter = "";
    while (index >= 0) {
        letter = String.fromCharCode((index % 26) + 65) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}

/**
 * Converts Excel column letter(s) to zero-based index
 * Supports multi-letter columns (A, Z, AA, AB, etc.)
 * @param {string} col - Column letter(s) like "A", "Z", "AA", "AB"
 * @returns {number} Zero-based column index
 */
function colLetterToIndex(col) {
    let index = 0;
    const upper = col.toUpperCase();
    for (let i = 0; i < upper.length; i++) {
        index = index * 26 + (upper.charCodeAt(i) - 64);
    }
    return index - 1; // Convert to zero-based
}

// ============================================================================
// Data Reading
// ============================================================================

/**
 * Reads Excel data from the workbook
 * @param {Object} state - Application state object
 * @param {Function} updateContextInfo - Callback to update UI context info
 * @param {Function} logDiagnostic - Optional callback for diagnostic logging
 * @returns {Promise<void>}
 */
async function readExcelData(state, updateContextInfo, logDiagnostic) {
    const log = logDiagnostic || (() => {});
    
    try {
        await Excel.run(async (ctx) => {
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            // Read sheets based on scope setting
            const allSheetsData = [];
            const shouldReadAllSheets = state.worksheetScope === "all";
            
            // Get active sheet first
            const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await ctx.sync();
            
            const activeSheetName = activeSheet.name;
            
            // Determine which sheets to read
            const sheetsToRead = shouldReadAllSheets 
                ? sheets.items.slice(0, 10) // All sheets (max 10)
                : [sheets.items.find(s => s.name === activeSheetName) || sheets.items[0]];
            
            for (const sheet of sheetsToRead) {
                try {
                    const usedRange = sheet.getUsedRange();
                    sheet.load("name");
                    usedRange.load(["address", "values", "rowCount", "columnCount", "columnIndex", "rowIndex"]);
                    await ctx.sync();
                    
                    const sheetName = sheet.name;
                    const values = usedRange.values;
                    const startCol = usedRange.columnIndex;
                    const startRow = usedRange.rowIndex;
                    const rowCount = usedRange.rowCount;
                    const colCount = usedRange.columnCount;
                    
                    // Handle empty sheets - skip if no data
                    if (rowCount === 0 || !values || values.length === 0) {
                        log(`Sheet "${sheetName}" has no data, skipping`);
                        continue;
                    }
                    
                    // Detect headers (first row) with guard
                    const headers = values.length > 0 ? (values[0] || []) : [];
                    
                    // Validate headers - check if first row looks like headers
                    const headerValidation = validateHeaders(headers);
                    
                    // Build column mapping
                    const columnMap = [];
                    for (let c = 0; c < colCount; c++) {
                        const colLetter = colIndexToLetter(startCol + c);
                        let headerName;
                        if (headerValidation.isValid && headers[c]) {
                            headerName = headers[c];
                        } else {
                            // Use generic column names if headers don't look valid
                            headerName = `Column ${colLetter}`;
                        }
                        columnMap.push({
                            letter: colLetter,
                            index: c,
                            header: headerName
                        });
                    }
                    
                    allSheetsData.push({
                        sheetName,
                        address: usedRange.address,
                        values,
                        headers: headerValidation.isValid ? headers : columnMap.map(c => c.header),
                        columnMap,
                        startRow: startRow + 1,
                        startCol: colIndexToLetter(startCol),
                        rowCount,
                        colCount,
                        dataStartRow: startRow + 2,
                        headerValidation
                    });
                    
                    log(`Read sheet "${sheetName}": ${rowCount} rows × ${colCount} cols`);
                } catch (e) {
                    // Sheet might be empty, log and skip it
                    const sheetName = sheet.name || "Unknown";
                    console.warn(`Skipping sheet ${sheetName}:`, e);
                    log(`Failed to read sheet "${sheetName}": ${e.message}`);
                }
            }
            
            // Handle case where no sheets have usable data
            if (allSheetsData.length === 0) {
                state.currentData = null;
                state.allSheetsData = [];
                updateContextInfo("No usable data found in any sheet");
                log("No usable data found in any sheet");
                return;
            }
            
            // Set current data to active sheet
            const activeSheetData = allSheetsData.find(s => s.sheetName === activeSheetName);
            state.currentData = activeSheetData || allSheetsData[0] || null;
            state.allSheetsData = shouldReadAllSheets ? allSheetsData : [];
            
            if (state.currentData) {
                const scopeText = shouldReadAllSheets ? ` (${allSheetsData.length} sheets)` : "";
                updateContextInfo(`${state.currentData.sheetName}: ${state.currentData.rowCount} rows × ${state.currentData.colCount} cols${scopeText}`);
            } else {
                updateContextInfo("No data");
            }
        });
    } catch (e) {
        // Log the actual error for debugging
        console.error("Failed to read Excel data:", e);
        const errorReason = e.message || "Unknown error";
        updateContextInfo(`Failed to read data: ${errorReason}`);
        state.currentData = null;
        state.allSheetsData = [];
        log(`readExcelData error: ${errorReason}`);
    }
}

/**
 * Validates if the first row looks like headers
 * @param {Array} headers - First row values
 * @returns {Object} Validation result with isValid flag and reason
 */
function validateHeaders(headers) {
    if (!headers || headers.length === 0) {
        return { isValid: false, reason: "Empty headers" };
    }
    
    // Count how many cells look like headers (strings, not numbers)
    let stringCount = 0;
    let numberCount = 0;
    let emptyCount = 0;
    
    for (const cell of headers) {
        if (cell === null || cell === undefined || cell === "") {
            emptyCount++;
        } else if (typeof cell === "string") {
            stringCount++;
        } else if (typeof cell === "number") {
            numberCount++;
        }
    }
    
    const total = headers.length;
    const stringRatio = stringCount / total;
    
    // Headers are valid if mostly strings (>50%)
    if (stringRatio >= 0.5) {
        return { isValid: true, reason: "Mostly string values" };
    }
    
    // If mostly numbers, probably not headers
    if (numberCount > stringCount) {
        return { isValid: false, reason: "First row appears to be data (mostly numbers)" };
    }
    
    return { isValid: true, reason: "Default assumption" };
}

/**
 * Builds data context string for AI prompts
 * @param {Object} state - Application state with currentData and allSheetsData
 * @returns {string} Formatted data context
 */
function buildDataContext(state) {
    if (!state.currentData) {
        return "ERROR: No Excel data available.";
    }
    
    const { sheetName, values, columnMap, rowCount, colCount, dataStartRow, address, headerValidation } = state.currentData;
    
    let context = `## EXCEL WORKBOOK DATA\n\n`;
    
    // List all sheets in workbook
    if (state.allSheetsData && state.allSheetsData.length > 1) {
        context += `**Available Sheets:** ${state.allSheetsData.map(s => s.sheetName).join(", ")}\n`;
        context += `**Active Sheet:** ${sheetName}\n\n`;
    } else {
        context += `**Sheet:** ${sheetName}\n`;
    }
    
    context += `**Data Range:** ${address}\n`;
    context += `**Total Rows:** ${rowCount} (including header)\n`;
    context += `**Total Columns:** ${colCount}\n`;
    
    // Add header validation note if headers look suspicious
    if (headerValidation && !headerValidation.isValid) {
        context += `**Note:** ${headerValidation.reason} - using generic column names\n`;
    }
    context += `\n`;
    
    // Column structure - CRITICAL for AI to understand
    context += `## COLUMN STRUCTURE\n`;
    context += `| Column Letter | Header Name | Sample Values |\n`;
    context += `|---------------|-------------|---------------|\n`;
    
    for (const col of columnMap) {
        // Get sample values from first few data rows
        const samples = [];
        for (let r = 1; r < Math.min(4, values.length); r++) {
            const val = values[r]?.[col.index];
            if (val !== null && val !== undefined && val !== "") {
                samples.push(String(val).substring(0, 20));
            }
        }
        context += `| ${col.letter} | ${col.header} | ${samples.join(", ")} |\n`;
    }
    
    context += `\n## DATA PREVIEW (First 30 rows)\n\n`;
    
    // Header row
    context += `| Row |`;
    for (const col of columnMap) {
        context += ` ${col.letter}: ${col.header} |`;
    }
    context += `\n|-----|`;
    for (let c = 0; c < colCount; c++) {
        context += `------------|`;
    }
    context += `\n`;
    
    // Data rows
    const maxRows = Math.min(30, values.length);
    for (let r = 0; r < maxRows; r++) {
        const rowNum = state.currentData.startRow + r;
        context += `| ${rowNum} |`;
        for (let c = 0; c < colCount; c++) {
            let val = values[r]?.[c];
            if (val === null || val === undefined) val = "";
            val = String(val).substring(0, 25);
            context += ` ${val} |`;
        }
        context += `\n`;
    }
    
    if (rowCount > 30) {
        context += `\n... and ${rowCount - 30} more rows\n`;
    }
    
    // Add unique values for key columns (for dropdowns)
    context += `\n## UNIQUE VALUES IN EACH COLUMN (for dropdowns)\n`;
    for (const col of columnMap) {
        const uniqueVals = new Set();
        for (let r = 1; r < values.length; r++) {
            const val = values[r]?.[col.index];
            if (val !== null && val !== undefined && val !== "") {
                uniqueVals.add(val);
            }
        }
        if (uniqueVals.size > 0 && uniqueVals.size <= 50) {
            context += `**${col.letter} (${col.header}):** ${Array.from(uniqueVals).slice(0, 20).join(", ")}`;
            if (uniqueVals.size > 20) context += ` ... (${uniqueVals.size} total)`;
            context += `\n`;
        }
    }
    
    // Add information about other sheets
    if (state.allSheetsData && state.allSheetsData.length > 1) {
        context += `\n## OTHER SHEETS IN WORKBOOK\n`;
        for (const sheet of state.allSheetsData) {
            if (sheet.sheetName === sheetName) continue; // Skip current sheet
            
            context += `\n### ${sheet.sheetName}\n`;
            context += `- Columns: ${sheet.headers.join(", ")}\n`;
            context += `- Rows: ${sheet.rowCount}\n`;
            
            // Show first few rows as sample
            if (sheet.values.length > 1) {
                context += `- Sample data (first 3 rows):\n`;
                for (let r = 0; r < Math.min(3, sheet.values.length); r++) {
                    const row = sheet.values[r];
                    context += `  ${r === 0 ? "Headers" : `Row ${r}`}: ${row.slice(0, 5).join(" | ")}\n`;
                }
            }
        }
        context += `\n**Note:** You can reference data from any sheet using sheet name (e.g., DeptManagers!A2:B10)\n`;
    }
    
    return context;
}

/**
 * Sets up selection change listener for auto-refresh
 * @param {Function} onSelectionChange - Callback when selection changes
 * @param {Function} logDiagnostic - Optional callback for diagnostic logging
 * @param {Function} showToast - Optional callback to show toast messages
 * @returns {Promise<Object|null>} Event handler reference or null
 */
async function setupSelectionListener(onSelectionChange, logDiagnostic, showToast) {
    const log = logDiagnostic || (() => {});
    const toast = showToast || (() => {});
    
    try {
        let handler = null;
        await Excel.run(async (ctx) => {
            const worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            handler = worksheet.onSelectionChanged.add(onSelectionChange);
            await ctx.sync();
            log("Selection listener attached successfully");
        });
        return handler;
    } catch (e) {
        console.warn("Could not attach selection listener:", e);
        log(`Selection listener failed: ${e.message}`);
        toast("Selection auto-refresh unavailable");
        return null;
    }
}

/**
 * Removes a selection change listener
 * @param {Object} handler - Event handler to remove
 * @returns {Promise<void>}
 */
async function removeSelectionListener(handler) {
    if (!handler) return;
    
    try {
        await Excel.run(async (ctx) => {
            handler.remove();
            await ctx.sync();
        });
    } catch (e) {
        console.warn("Could not remove selection listener:", e);
    }
}

// Export for ES modules
export {
    colIndexToLetter,
    colLetterToIndex,
    readExcelData,
    validateHeaders,
    buildDataContext,
    setupSelectionListener,
    removeSelectionListener
};
