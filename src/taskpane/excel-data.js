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
                    
                    // Detect worksheet protection status
                    let worksheetProtection = null;
                    try {
                        sheet.protection.load(["protected", "options"]);
                        await ctx.sync();
                        
                        if (sheet.protection.protected) {
                            worksheetProtection = {
                                protected: true,
                                options: {
                                    allowAutoFilter: sheet.protection.options.allowAutoFilter,
                                    allowDeleteColumns: sheet.protection.options.allowDeleteColumns,
                                    allowDeleteRows: sheet.protection.options.allowDeleteRows,
                                    allowFormatCells: sheet.protection.options.allowFormatCells,
                                    allowFormatColumns: sheet.protection.options.allowFormatColumns,
                                    allowFormatRows: sheet.protection.options.allowFormatRows,
                                    allowInsertColumns: sheet.protection.options.allowInsertColumns,
                                    allowInsertRows: sheet.protection.options.allowInsertRows,
                                    allowInsertHyperlinks: sheet.protection.options.allowInsertHyperlinks,
                                    allowPivotTables: sheet.protection.options.allowPivotTables,
                                    allowSort: sheet.protection.options.allowSort,
                                    selectionMode: sheet.protection.options.selectionMode
                                }
                            };
                            log(`Sheet "${sheetName}" is protected`);
                        } else {
                            worksheetProtection = { protected: false };
                        }
                    } catch (protectionError) {
                        log(`Could not read worksheet protection status for "${sheetName}": ${protectionError.message}`);
                        worksheetProtection = { protected: false, error: protectionError.message };
                    }
                    
                    // Detect comments and notes on this sheet
                    const commentsAndNotes = { comments: [], notes: [] };
                    try {
                        // Load threaded comments
                        sheet.comments.load("items");
                        await ctx.sync();
                        
                        if (sheet.comments.items.length > 0) {
                            // Batch load comment properties
                            for (const comment of sheet.comments.items) {
                                comment.load(["id", "authorName", "content", "creationDate", "resolved"]);
                                comment.replies.load("items");
                            }
                            await ctx.sync();
                            
                            // Batch load reply properties
                            for (const comment of sheet.comments.items) {
                                for (const reply of comment.replies.items) {
                                    reply.load(["id", "authorName", "content", "creationDate"]);
                                }
                            }
                            await ctx.sync();
                            
                            // Extract comment data
                            for (const comment of sheet.comments.items) {
                                try {
                                    // Get the cell location of the comment
                                    const location = comment.getLocation();
                                    location.load("address");
                                    await ctx.sync();
                                    
                                    commentsAndNotes.comments.push({
                                        cell: location.address,
                                        author: comment.authorName,
                                        content: comment.content,
                                        resolved: comment.resolved,
                                        createdDate: comment.creationDate,
                                        replyCount: comment.replies.items.length,
                                        replies: comment.replies.items.map(r => ({
                                            author: r.authorName,
                                            content: r.content,
                                            createdDate: r.creationDate
                                        }))
                                    });
                                } catch (locError) {
                                    log(`Could not get location for comment: ${locError.message}`);
                                }
                            }
                            log(`Found ${commentsAndNotes.comments.length} comments on "${sheetName}"`);
                        }
                        
                        // Note: Legacy notes detection is limited in Office.js
                        // Notes are accessed via range.note property, but there's no efficient way
                        // to enumerate all notes without checking each cell individually
                        // For performance, we skip note enumeration and rely on comments API
                        
                    } catch (commentError) {
                        log(`Could not read comments for sheet "${sheetName}": ${commentError.message}`);
                    }
                    
                    // Detect sparkline groups on this sheet
                    const sparklineGroups = [];
                    try {
                        // Check if sparklineGroups API is available
                        if (sheet.sparklineGroups) {
                            sheet.sparklineGroups.load("items");
                            await ctx.sync();
                            
                            if (sheet.sparklineGroups.items.length > 0) {
                                // Batch load sparkline group properties
                                for (const group of sheet.sparklineGroups.items) {
                                    group.load(["type"]);
                                    group.load("sparklines/items/location");
                                }
                                await ctx.sync();
                                
                                // Batch load location addresses
                                for (const group of sheet.sparklineGroups.items) {
                                    for (const sparkline of group.sparklines.items) {
                                        sparkline.location.load("address");
                                    }
                                }
                                await ctx.sync();
                                
                                // Helper to convert sparkline type enum to string
                                const getSparklineTypeName = (type) => {
                                    if (type === Excel.SparklineType.line) return "Line";
                                    if (type === Excel.SparklineType.column) return "Column";
                                    if (type === Excel.SparklineType.winLoss) return "WinLoss";
                                    return "Unknown";
                                };
                                
                                // Extract sparkline group data
                                for (const group of sheet.sparklineGroups.items) {
                                    try {
                                        const locations = group.sparklines.items.map(s => s.location.address);
                                        sparklineGroups.push({
                                            type: getSparklineTypeName(group.type),
                                            locations: locations,
                                            count: locations.length
                                        });
                                    } catch (groupError) {
                                        log(`Error reading sparkline group details: ${groupError.message}`);
                                    }
                                }
                                log(`Found ${sparklineGroups.length} sparkline group(s) on "${sheetName}"`);
                            }
                        }
                    } catch (sparklineError) {
                        log(`Could not read sparklines for sheet "${sheetName}": ${sparklineError.message}`);
                    }
                    
                    // Detect data type cells (EntityCellValue, LinkedEntityCellValue)
                    const dataTypeCells = [];
                    try {
                        // Sample first 50x10 cells for performance (full scan expensive)
                        const sampleRows = Math.min(49, rowCount - 1);
                        const sampleCols = Math.min(9, colCount - 1);
                        if (sampleRows >= 0 && sampleCols >= 0) {
                            const sampleRange = usedRange.getCell(0, 0).getResizedRange(sampleRows, sampleCols);
                            sampleRange.load(["address", "valueTypes", "valuesAsJson"]);
                            await ctx.sync();
                            
                            for (let r = 0; r < sampleRange.valueTypes.length; r++) {
                                for (let c = 0; c < sampleRange.valueTypes[r].length; c++) {
                                    const cellType = sampleRange.valueTypes[r][c];
                                    if (cellType === "Entity" || cellType === "LinkedEntity") {
                                        const cellValue = sampleRange.valuesAsJson[r][c];
                                        const cellAddress = `${colIndexToLetter(startCol + c)}${startRow + r + 1}`;
                                        dataTypeCells.push({
                                            address: cellAddress,
                                            type: cellType,
                                            text: cellValue.text || "",
                                            basicValue: cellValue.basicValue || "",
                                            properties: Object.keys(cellValue.properties || {}).slice(0, 5),
                                            serviceId: cellType === "LinkedEntity" ? (cellValue.serviceId || "Unknown") : null
                                        });
                                    }
                                }
                            }
                            if (dataTypeCells.length > 0) {
                                log(`Found ${dataTypeCells.length} data type cells on "${sheetName}" (sampled first ${sampleRows + 1}x${sampleCols + 1} cells)`);
                            }
                        }
                    } catch (dataTypeError) {
                        log(`Could not read data types for sheet "${sheetName}": ${dataTypeError.message}`);
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
                        namedRanges: worksheetNamedRanges,
                        protection: worksheetProtection,
                        commentsAndNotes,
                        sparklineGroups,
                        dataTypeCells
                    });
                    
                    log(`Read sheet "${sheetName}": ${rowCount} rows × ${colCount} cols, ${pivotTables.length} PivotTables, ${worksheetNamedRanges.length} named ranges, ${commentsAndNotes.comments.length} comments`);
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
            
            // Detect workbook protection status
            let workbookProtection = null;
            try {
                ctx.workbook.protection.load("protected");
                await ctx.sync();
                workbookProtection = {
                    protected: ctx.workbook.protection.protected
                };
                if (workbookProtection.protected) {
                    log("Workbook structure is protected");
                }
            } catch (protectionError) {
                log(`Could not read workbook protection status: ${protectionError.message}`);
                workbookProtection = { protected: false, error: protectionError.message };
            }
            
            // Handle case where no sheets have usable data
            if (allSheetsData.length === 0) {
                state.currentData = null;
                state.allSheetsData = [];
                state.workbookNamedRanges = workbookNamedRanges;
                state.workbookProtection = workbookProtection;
                updateContextInfo("No usable data found in any sheet");
                log("No usable data found in any sheet");
                return;
            }
            
            // Set current data to active sheet
            const activeSheetData = allSheetsData.find(s => s.sheetName === activeSheetName);
            state.currentData = activeSheetData || allSheetsData[0] || null;
            state.allSheetsData = shouldReadAllSheets ? allSheetsData : [];
            state.workbookNamedRanges = workbookNamedRanges;
            state.workbookProtection = workbookProtection;
            
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
    
    // Add protection status information
    if (state.currentData && state.currentData.protection) {
        const prot = state.currentData.protection;
        context += `\n## WORKSHEET PROTECTION STATUS\n`;
        if (prot.protected) {
            context += `**Status:** Protected\n`;
            context += `**Allowed Actions:**\n`;
            if (prot.options) {
                const allowed = [];
                if (prot.options.allowFormatCells) allowed.push("Format cells");
                if (prot.options.allowSort) allowed.push("Sort");
                if (prot.options.allowAutoFilter) allowed.push("Filter");
                if (prot.options.allowInsertRows) allowed.push("Insert rows");
                if (prot.options.allowInsertColumns) allowed.push("Insert columns");
                if (prot.options.allowDeleteRows) allowed.push("Delete rows");
                if (prot.options.allowDeleteColumns) allowed.push("Delete columns");
                if (prot.options.allowPivotTables) allowed.push("PivotTables");
                if (allowed.length > 0) {
                    context += `- ${allowed.join(", ")}\n`;
                } else {
                    context += `- None (fully locked)\n`;
                }
                context += `**Selection Mode:** ${prot.options.selectionMode}\n`;
            }
            context += `\n**Note:** To modify protection, use unprotectWorksheet action (password may be required).\n`;
        } else {
            context += `**Status:** Not protected\n`;
            context += `**Note:** You can protect this worksheet with protectWorksheet action to prevent unauthorized changes.\n`;
        }
    }
    
    if (state.workbookProtection) {
        context += `\n## WORKBOOK PROTECTION STATUS\n`;
        if (state.workbookProtection.protected) {
            context += `**Status:** Protected (structure locked)\n`;
            context += `**Effect:** Cannot add, delete, rename, or move sheets\n`;
            context += `**Note:** To modify structure, use unprotectWorkbook action (password may be required).\n`;
        } else {
            context += `**Status:** Not protected\n`;
            context += `**Note:** You can protect workbook structure with protectWorkbook action.\n`;
        }
    }
    
    // Add comments and notes information (aggregate from all sheets when in multi-sheet mode)
    const allComments = [];
    const allNotes = [];
    
    // Collect comments/notes from current sheet
    if (state.currentData && state.currentData.commentsAndNotes) {
        const { comments, notes } = state.currentData.commentsAndNotes;
        for (const comment of comments) {
            allComments.push({ ...comment, sheetName: state.currentData.sheetName });
        }
        for (const note of notes) {
            allNotes.push({ ...note, sheetName: state.currentData.sheetName });
        }
    }
    
    // Collect comments/notes from other sheets when in multi-sheet mode
    if (state.allSheetsData && state.allSheetsData.length > 0) {
        for (const sheet of state.allSheetsData) {
            if (sheet.sheetName === sheetName) continue; // Already added from currentData
            if (sheet.commentsAndNotes) {
                const { comments, notes } = sheet.commentsAndNotes;
                for (const comment of comments) {
                    allComments.push({ ...comment, sheetName: sheet.sheetName });
                }
                for (const note of notes) {
                    allNotes.push({ ...note, sheetName: sheet.sheetName });
                }
            }
        }
    }
    
    if (allComments.length > 0 || allNotes.length > 0) {
        context += `\n## EXISTING COMMENTS AND NOTES\n`;
        
        if (allComments.length > 0) {
            context += `\n### Threaded Comments (${allComments.length} total)\n`;
            context += `Modern collaboration comments with replies and resolution tracking.\n\n`;
            
            // Limit to first 15 comments across all sheets
            for (const comment of allComments.slice(0, 15)) {
                const cellRef = comment.sheetName !== sheetName 
                    ? `${comment.sheetName}!${comment.cell}` 
                    : comment.cell;
                context += `**${cellRef}** by ${comment.author}:\n`;
                context += `- Content: "${comment.content.substring(0, 100)}${comment.content.length > 100 ? '...' : ''}"\n`;
                context += `- Status: ${comment.resolved ? 'Resolved' : 'Open'}\n`;
                if (comment.replyCount > 0) {
                    context += `- Replies: ${comment.replyCount}\n`;
                }
                context += `\n`;
            }
            
            if (allComments.length > 15) {
                context += `... and ${allComments.length - 15} more comments across sheets\n\n`;
            }
        }
        
        if (allNotes.length > 0) {
            context += `\n### Notes (${allNotes.length} total)\n`;
            context += `Legacy annotations for reminders and documentation.\n\n`;
            
            // Limit to first 15 notes across all sheets
            for (const note of allNotes.slice(0, 15)) {
                const cellRef = note.sheetName !== sheetName 
                    ? `${note.sheetName}!${note.cell}` 
                    : note.cell;
                context += `**${cellRef}**: "${note.text.substring(0, 80)}${note.text.length > 80 ? '...' : ''}"\n`;
            }
            
            if (allNotes.length > 15) {
                context += `... and ${allNotes.length - 15} more notes across sheets\n\n`;
            }
        }
        
        context += `\n**Actions Available:**\n`;
        context += `- Add comments for collaboration: addComment action\n`;
        context += `- Add notes for documentation: addNote action\n`;
        context += `- Reply to comments: replyToComment action\n`;
        context += `- Resolve discussions: resolveComment action\n`;
        context += `- Edit or delete: editComment, deleteComment, editNote, deleteNote actions\n`;
    }
    
    // Add sparkline information (aggregate from all sheets when in multi-sheet mode)
    const allSparklineGroups = [];
    
    // Collect sparklines from current sheet
    if (state.currentData && state.currentData.sparklineGroups && state.currentData.sparklineGroups.length > 0) {
        for (const group of state.currentData.sparklineGroups) {
            allSparklineGroups.push({ ...group, sheetName: state.currentData.sheetName });
        }
    }
    
    // Collect sparklines from other sheets when in multi-sheet mode
    if (state.allSheetsData && state.allSheetsData.length > 0) {
        for (const sheet of state.allSheetsData) {
            if (sheet.sheetName === sheetName) continue; // Already added from currentData
            if (sheet.sparklineGroups && sheet.sparklineGroups.length > 0) {
                for (const group of sheet.sparklineGroups) {
                    allSparklineGroups.push({ ...group, sheetName: sheet.sheetName });
                }
            }
        }
    }
    
    if (allSparklineGroups.length > 0) {
        context += `\n## EXISTING SPARKLINES IN WORKBOOK\n`;
        context += `Compact inline visualizations for trend analysis.\n\n`;
        
        // Group by sheet
        const sparklinesBySheet = {};
        for (const group of allSparklineGroups) {
            if (!sparklinesBySheet[group.sheetName]) {
                sparklinesBySheet[group.sheetName] = [];
            }
            sparklinesBySheet[group.sheetName].push(group);
        }
        
        for (const [sheet, groups] of Object.entries(sparklinesBySheet)) {
            const sheetLabel = sheet === sheetName ? `${sheet} (current)` : sheet;
            context += `### ${sheetLabel}\n`;
            for (const group of groups) {
                const locationSummary = group.locations.length <= 3 
                    ? group.locations.join(", ")
                    : `${group.locations.slice(0, 3).join(", ")} ... (${group.count} total)`;
                context += `- **${group.type}** sparkline(s) at: ${locationSummary}\n`;
            }
            context += `\n`;
        }
        
        context += `**Actions Available:**\n`;
        context += `- Create new sparklines: createSparkline action\n`;
        context += `- Configure existing: configureSparkline action (colors, markers, axes)\n`;
        context += `- Delete sparklines: deleteSparkline action\n`;
        context += `\n**Note:** Sparklines require ExcelApi 1.10+ (Excel 365, Excel 2019+, or Excel Online).\n`;
    }
    
    // Add data type cells information (aggregate from all sheets)
    const allDataTypeCells = [];
    
    // Current sheet
    if (state.currentData?.dataTypeCells) {
        allDataTypeCells.push(...state.currentData.dataTypeCells.map(c => ({...c, sheetName: state.currentData.sheetName})));
    }
    
    // All sheets
    state.allSheetsData?.forEach(sheet => {
        if (sheet.sheetName === sheetName) return; // Already added from currentData
        if (sheet.dataTypeCells) {
            allDataTypeCells.push(...sheet.dataTypeCells.map(c => ({...c, sheetName: sheet.sheetName})));
        }
    });
    
    // Limit to top 20 total for performance/token limit
    const topDataTypes = allDataTypeCells.slice(0, 20);
    
    if (topDataTypes.length > 0) {
        context += `\n## EXISTING DATA TYPE CELLS (${topDataTypes.length} sampled)\n`;
        context += `Entity cards with hover properties. Built-in Stocks/Geography shown as LinkedEntity.\n\n`;
        
        // Group by sheet
        const bySheet = {};
        topDataTypes.forEach(cell => {
            if (!bySheet[cell.sheetName]) bySheet[cell.sheetName] = [];
            bySheet[cell.sheetName].push(cell);
        });
        
        Object.entries(bySheet).forEach(([sheet, cells]) => {
            const label = sheet === sheetName ? `${sheet} (current)` : sheet;
            context += `### ${label} (${cells.length})\n`;
            cells.slice(0, 10).forEach(cell => {
                const props = cell.properties.length > 0 ? ` (${cell.properties.join(', ')})` : '';
                const service = cell.serviceId ? ` [${cell.serviceId}]` : '';
                context += `- **${cell.address}**: ${cell.type} - "${cell.text}"${props}${service}\n`;
            });
            if (cells.length > 10) context += `... ${cells.length - 10} more\n`;
            context += `\n`;
        });
        
        context += `**Note**: Custom entities fully supported. LinkedEntity (Stocks/Geography) require manual UI conversion. Use \`insertDataType\`/\`refreshDataType\` for custom entities, reference properties in formulas (e.g., \`=A2.Price\`).\n`;
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
