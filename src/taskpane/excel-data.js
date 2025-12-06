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
                    
                    // Detect PivotTables on this sheet (optimized with batched loads)
                    const pivotTables = [];
                    try {
                        // First sync: load all PivotTables for this sheet
                        sheet.pivotTables.load("items");
                        await ctx.sync();
                        
                        if (sheet.pivotTables.items.length > 0) {
                            // Second sync: batch load all PivotTable properties and hierarchy collections
                            for (const pt of sheet.pivotTables.items) {
                                pt.load(["name"]);
                                pt.layout.load("layoutType");
                                pt.rowHierarchies.load("items");
                                pt.columnHierarchies.load("items");
                                pt.dataHierarchies.load("items");
                                pt.filterHierarchies.load("items");
                            }
                            await ctx.sync();
                            
                            // Third sync: batch load all hierarchy item properties
                            for (const pt of sheet.pivotTables.items) {
                                for (const h of pt.rowHierarchies.items) {
                                    h.load("name");
                                }
                                for (const h of pt.columnHierarchies.items) {
                                    h.load("name");
                                }
                                for (const h of pt.dataHierarchies.items) {
                                    h.load(["name", "summarizeBy"]);
                                }
                                for (const h of pt.filterHierarchies.items) {
                                    h.load("name");
                                }
                            }
                            await ctx.sync();
                            
                            // Helper functions for enum to string conversion
                            const getAggregationName = (summarizeBy) => {
                                const aggMap = {
                                    [Excel.AggregationFunction.sum]: "Sum",
                                    [Excel.AggregationFunction.count]: "Count",
                                    [Excel.AggregationFunction.average]: "Average",
                                    [Excel.AggregationFunction.max]: "Max",
                                    [Excel.AggregationFunction.min]: "Min",
                                    [Excel.AggregationFunction.countNumbers]: "CountNumbers",
                                    [Excel.AggregationFunction.standardDeviation]: "StdDev",
                                    [Excel.AggregationFunction.variance]: "Var"
                                };
                                return aggMap[summarizeBy] || "Sum";
                            };
                            
                            const getLayoutName = (layoutType) => {
                                if (layoutType === Excel.PivotLayoutType.compact) return "Compact";
                                if (layoutType === Excel.PivotLayoutType.outline) return "Outline";
                                if (layoutType === Excel.PivotLayoutType.tabular) return "Tabular";
                                return "Compact";
                            };
                            
                            // Now extract data from already-loaded properties
                            for (const pt of sheet.pivotTables.items) {
                                try {
                                    pivotTables.push({
                                        name: pt.name,
                                        layout: getLayoutName(pt.layout.layoutType),
                                        rowFields: pt.rowHierarchies.items.map(h => h.name),
                                        columnFields: pt.columnHierarchies.items.map(h => h.name),
                                        dataFields: pt.dataHierarchies.items.map(h => ({ 
                                            field: h.name, 
                                            function: getAggregationName(h.summarizeBy) 
                                        })),
                                        filterFields: pt.filterHierarchies.items.map(h => h.name)
                                    });
                                    log(`Found PivotTable "${pt.name}" on sheet "${sheetName}"`);
                                } catch (ptError) {
                                    log(`Error reading PivotTable "${pt.name}" details: ${ptError.message}`);
                                }
                            }
                        }
                    } catch (pivotError) {
                        // PivotTables not available or error reading them
                        log(`Could not read PivotTables for sheet "${sheetName}": ${pivotError.message}`);
                    }
                    
                    // Detect worksheet-scoped named ranges
                    const worksheetNamedRanges = [];
                    try {
                        sheet.names.load("items");
                        await ctx.sync();
                        
                        if (sheet.names.items.length > 0) {
                            for (const item of sheet.names.items) {
                                item.load(["name", "formula", "comment", "type", "visible"]);
                            }
                            await ctx.sync();
                            
                            for (const item of sheet.names.items) {
                                worksheetNamedRanges.push({
                                    name: item.name,
                                    formula: item.formula,
                                    comment: item.comment || "",
                                    type: item.type,
                                    visible: item.visible
                                });
                            }
                            log(`Found ${worksheetNamedRanges.length} worksheet-scoped named ranges on "${sheetName}"`);
                        }
                    } catch (namedRangeError) {
                        log(`Could not read named ranges for sheet "${sheetName}": ${namedRangeError.message}`);
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
                        headerValidation,
                        pivotTables,
                        namedRanges: worksheetNamedRanges
                    });
                    
                    log(`Read sheet "${sheetName}": ${rowCount} rows × ${colCount} cols, ${pivotTables.length} PivotTables, ${worksheetNamedRanges.length} named ranges`);
                } catch (e) {
                    // Sheet might be empty, log and skip it
                    const sheetName = sheet.name || "Unknown";
                    console.warn(`Skipping sheet ${sheetName}:`, e);
                    log(`Failed to read sheet "${sheetName}": ${e.message}`);
                }
            }
            
            // Detect workbook-scoped named ranges
            const workbookNamedRanges = [];
            try {
                ctx.workbook.names.load("items");
                await ctx.sync();
                
                if (ctx.workbook.names.items.length > 0) {
                    for (const item of ctx.workbook.names.items) {
                        item.load(["name", "formula", "comment", "type", "visible"]);
                    }
                    await ctx.sync();
                    
                    for (const item of ctx.workbook.names.items) {
                        workbookNamedRanges.push({
                            name: item.name,
                            formula: item.formula,
                            comment: item.comment || "",
                            type: item.type,
                            visible: item.visible
                        });
                    }
                    log(`Found ${workbookNamedRanges.length} workbook-scoped named ranges`);
                }
            } catch (namedRangeError) {
                log(`Could not read workbook named ranges: ${namedRangeError.message}`);
            }
            
            // Handle case where no sheets have usable data
            if (allSheetsData.length === 0) {
                state.currentData = null;
                state.allSheetsData = [];
                state.workbookNamedRanges = workbookNamedRanges;
                updateContextInfo("No usable data found in any sheet");
                log("No usable data found in any sheet");
                return;
            }
            
            // Set current data to active sheet
            const activeSheetData = allSheetsData.find(s => s.sheetName === activeSheetName);
            state.currentData = activeSheetData || allSheetsData[0] || null;
            state.allSheetsData = shouldReadAllSheets ? allSheetsData : [];
            state.workbookNamedRanges = workbookNamedRanges;
            
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
    
    // Add information about existing PivotTables
    const allPivotTables = [];
    
    // Collect PivotTables from current sheet
    if (state.currentData.pivotTables && state.currentData.pivotTables.length > 0) {
        for (const pt of state.currentData.pivotTables) {
            allPivotTables.push({ ...pt, sheetName: state.currentData.sheetName });
        }
    }
    
    // Collect PivotTables from other sheets
    if (state.allSheetsData && state.allSheetsData.length > 0) {
        for (const sheet of state.allSheetsData) {
            if (sheet.sheetName === sheetName) continue; // Already added
            if (sheet.pivotTables && sheet.pivotTables.length > 0) {
                for (const pt of sheet.pivotTables) {
                    allPivotTables.push({ ...pt, sheetName: sheet.sheetName });
                }
            }
        }
    }
    
    if (allPivotTables.length > 0) {
        context += `\n## EXISTING PIVOTTABLES IN WORKBOOK\n`;
        for (const pt of allPivotTables) {
            context += `\n### ${pt.name} (on sheet "${pt.sheetName}")\n`;
            context += `- Layout: ${pt.layout}\n`;
            if (pt.rowFields.length > 0) {
                context += `- Row Fields: ${pt.rowFields.join(", ")}\n`;
            }
            if (pt.columnFields.length > 0) {
                context += `- Column Fields: ${pt.columnFields.join(", ")}\n`;
            }
            if (pt.dataFields.length > 0) {
                context += `- Data Fields: ${pt.dataFields.map(d => `${d.function} of ${d.field}`).join(", ")}\n`;
            }
            if (pt.filterFields.length > 0) {
                context += `- Filter Fields: ${pt.filterFields.join(", ")}\n`;
            }
        }
        context += `\n**Note:** You can refresh existing PivotTables with refreshPivotTable action or create new ones with createPivotTable.\n`;
    }
    
    // Collect all named ranges (workbook + worksheet scoped)
    const allNamedRanges = [];
    
    // Add workbook-scoped names
    if (state.workbookNamedRanges && state.workbookNamedRanges.length > 0) {
        for (const nr of state.workbookNamedRanges) {
            allNamedRanges.push({ ...nr, scope: "workbook" });
        }
    }
    
    // Add worksheet-scoped names from current sheet
    if (state.currentData.namedRanges && state.currentData.namedRanges.length > 0) {
        for (const nr of state.currentData.namedRanges) {
            allNamedRanges.push({ ...nr, scope: "worksheet", sheetName: state.currentData.sheetName });
        }
    }
    
    // Add worksheet-scoped names from other sheets
    if (state.allSheetsData && state.allSheetsData.length > 0) {
        for (const sheet of state.allSheetsData) {
            if (sheet.sheetName === sheetName) continue; // Already added
            if (sheet.namedRanges && sheet.namedRanges.length > 0) {
                for (const nr of sheet.namedRanges) {
                    allNamedRanges.push({ ...nr, scope: "worksheet", sheetName: sheet.sheetName });
                }
            }
        }
    }
    
    if (allNamedRanges.length > 0) {
        context += `\n## EXISTING NAMED RANGES IN WORKBOOK\n`;
        context += `Named ranges improve formula readability and maintainability. You can reference these in formulas.\n\n`;
        
        // Group by scope
        const workbookNames = allNamedRanges.filter(nr => nr.scope === "workbook");
        const worksheetNames = allNamedRanges.filter(nr => nr.scope === "worksheet");
        
        if (workbookNames.length > 0) {
            context += `### Workbook-Scoped Names (accessible from any sheet)\n`;
            for (const nr of workbookNames) {
                context += `- **${nr.name}**: ${nr.formula}`;
                if (nr.comment) context += ` (${nr.comment})`;
                context += `\n`;
            }
            context += `\n`;
        }
        
        if (worksheetNames.length > 0) {
            context += `### Worksheet-Scoped Names (sheet-specific)\n`;
            for (const nr of worksheetNames) {
                context += `- **${nr.name}** (${nr.sheetName}): ${nr.formula}`;
                if (nr.comment) context += ` (${nr.comment})`;
                context += `\n`;
            }
            context += `\n`;
        }
        
        context += `**Usage in formulas:** Reference by name (e.g., =SUM(SalesData) or =TotalRevenue*0.1)\n`;
        context += `**Note:** You can create new named ranges with createNamedRange action for frequently used ranges.\n`;
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
