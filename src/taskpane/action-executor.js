/**
 * Action Executor Module
 * Handles execution of Excel actions (formulas, values, formatting, charts, etc.)
 */

/* global Excel */

import { colIndexToLetter, colLetterToIndex } from './excel-data.js';

// ============================================================================
// Diagnostics
// ============================================================================

let diagnosticLogger = null;

/**
 * Sets the diagnostic logger function
 * @param {Function} logger - Function to log diagnostic messages
 */
function setDiagnosticLogger(logger) {
    diagnosticLogger = logger;
}

/**
 * Logs a diagnostic message
 * @param {string} message - Message to log
 */
function logDiag(message) {
    if (diagnosticLogger) {
        diagnosticLogger(message);
    }
}

// ============================================================================
// Main Executor
// ============================================================================

/**
 * Executes a single action
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action to execute
 * @returns {Promise<void>}
 */
async function executeAction(ctx, sheet, action) {
    const { type, target, source, chartType, data } = action;
    
    logDiag(`Executing ${type} action on ${target || 'N/A'}`);

    // Sheet creation doesn't need a range
    if (type === "sheet") {
        await createSheet(ctx, target, data);
        return;
    }
    
    if (!target) {
        logDiag(`Skipped action: No target specified`);
        throw new Error("No target specified");
    }
    
    // Actions that use logical names (table names, PivotTable names) instead of range addresses
    // These should NOT pre-load a range as the target is not a valid range address
    const logicalNameActions = [
        "createPivotTable",      // target is source range, but handler resolves it
        "addPivotField",         // target is PivotTable name
        "configurePivotLayout",  // target is PivotTable name
        "refreshPivotTable",     // target is PivotTable name
        "deletePivotTable",      // target is PivotTable name
        "styleTable",            // target is table name
        "addTableRow",           // target is table name
        "addTableColumn",        // target is table name
        "resizeTable",           // target is table name
        "convertToRange",        // target is table name
        "toggleTableTotals",     // target is table name
        "insertRows",            // target is row number, not range
        "insertColumns",         // target is column letter, not range
        "deleteRows",            // target is row range like "10:15"
        "deleteColumns",         // target is column range like "D:F"
        "createSlicer",          // target is table/pivot name
        "configureSlicer",       // target is slicer name
        "connectSlicerToTable",  // target is slicer name
        "connectSlicerToPivot",  // target is slicer name
        "deleteSlicer"           // target is slicer name
    ];
    
    // Only pre-load range for actions that actually need it
    let range = null;
    if (!logicalNameActions.includes(type)) {
        range = sheet.getRange(target);
        range.load(["rowCount", "columnCount"]);
        await ctx.sync();
    }
    
    switch (type) {
        case "formula":
            await applyFormula(range, data);
            break;
            
        case "values":
            applyValues(range, data);
            break;
            
        case "format":
            await applyFormat(ctx, range, data);
            break;
            
        case "conditionalFormat":
            await applyConditionalFormat(ctx, range, data);
            break;
            
        case "clearFormat":
            await clearConditionalFormat(ctx, range);
            break;
            
        case "validation":
            await applyValidation(ctx, sheet, range, source);
            break;
            
        case "chart":
            await createChart(ctx, sheet, range, action);
            break;
            
        case "pivotChart":
            await createPivotChart(ctx, sheet, range, action);
            break;
            
        case "sort":
            applySort(range, data);
            break;
            
        case "autofill":
            if (source) {
                const sourceRange = sheet.getRange(source);
                sourceRange.autoFill(range, Excel.AutoFillType.fillDefault);
            }
            break;
            
        case "copy":
            await applyCopy(ctx, sheet, source, target);
            break;
            
        case "copyValues":
            await applyCopyValues(ctx, sheet, source, target);
            break;
            
        case "filter":
            await applyFilter(ctx, sheet, range, data);
            break;
            
        case "clearFilter":
            await clearFilter(ctx, sheet);
            break;
            
        case "removeDuplicates":
            await removeDuplicates(ctx, range, data);
            break;
            
        case "createTable":
            await createTable(ctx, sheet, range, action);
            break;
            
        case "styleTable":
            await styleTable(ctx, sheet, target, data);
            break;
            
        case "addTableRow":
            await addTableRow(ctx, sheet, action);
            break;
            
        case "addTableColumn":
            await addTableColumn(ctx, sheet, action);
            break;
            
        case "resizeTable":
            await resizeTable(ctx, sheet, action);
            break;
            
        case "convertToRange":
            await convertToRange(ctx, sheet, target);
            break;
            
        case "toggleTableTotals":
            await toggleTableTotals(ctx, sheet, action);
            break;
            
        case "insertRows":
            await insertRows(ctx, sheet, action);
            break;
            
        case "insertColumns":
            await insertColumns(ctx, sheet, action);
            break;
            
        case "deleteRows":
            await deleteRows(ctx, sheet, action);
            break;
            
        case "deleteColumns":
            await deleteColumns(ctx, sheet, action);
            break;
            
        case "mergeCells":
            await mergeCells(ctx, sheet, action);
            break;
            
        case "unmergeCells":
            await unmergeCells(ctx, sheet, action);
            break;
            
        case "findReplace":
            await findReplace(ctx, sheet, action);
            break;
            
        case "textToColumns":
            await textToColumns(ctx, sheet, action);
            break;
            
        case "createPivotTable":
            await createPivotTable(ctx, sheet, action);
            break;
            
        case "addPivotField":
            await addPivotField(ctx, sheet, action);
            break;
            
        case "configurePivotLayout":
            await configurePivotLayout(ctx, sheet, action);
            break;
            
        case "refreshPivotTable":
            await refreshPivotTable(ctx, sheet, action);
            break;
            
        case "deletePivotTable":
            await deletePivotTable(ctx, sheet, action);
            break;
            
        case "createSlicer":
            await createSlicer(ctx, sheet, action);
            break;
            
        case "configureSlicer":
            await configureSlicer(ctx, sheet, action);
            break;
            
        case "connectSlicerToTable":
            await connectSlicerToTable(ctx, sheet, action);
            break;
            
        case "connectSlicerToPivot":
            await connectSlicerToPivot(ctx, sheet, action);
            break;
            
        case "deleteSlicer":
            await deleteSlicer(ctx, sheet, action);
            break;
            
        default:
            // Guard for future action types not yet implemented
            // These action types are advertised in AI prompts but executor support is pending
            const futureActionTypes = [
                // Shapes and images
                "insertShape", "insertImage", "insertTextBox", "formatShape",
                "deleteShape", "groupShapes", "arrangeShapes",
                // Comments and notes
                "addComment", "addNote", "editComment", "editNote",
                "deleteComment", "deleteNote", "replyToComment", "resolveComment",
                // Protection
                "protectWorksheet", "unprotectWorksheet", "protectRange",
                "unprotectRange", "protectWorkbook", "unprotectWorkbook",
                // Page setup
                "setPageSetup", "setPageMargins", "setPageOrientation",
                "setPrintArea", "setHeaderFooter", "setPageBreaks"
            ];
            
            if (futureActionTypes.includes(type)) {
                logDiag(`Action type "${type}" is planned but not yet implemented - skipping gracefully`);
                console.warn(`Action type "${type}" is not yet supported. This feature is coming soon.`);
                // Don't throw - just log and continue
                return;
            }
            
            // For truly unknown types, try to apply data if present
            if (data) range.values = [[data]];
            logDiag(`Applied default action with data`);
    }
    
    logDiag(`Completed ${type} action on ${target}`);
}

// ============================================================================
// Formula Application
// ============================================================================

/**
 * Applies a formula to a range with proper row/column adjustment
 * @param {Excel.Range} range - Target range
 * @param {string} formula - Formula to apply
 */
async function applyFormula(range, formula) {
    const rows = range.rowCount;
    const cols = range.columnCount;
    
    // For single cell, just set the formula
    if (rows === 1 && cols === 1) {
        range.formulas = [[formula]];
        logDiag(`Applied formula to single cell: ${formula}`);
        return;
    }
    
    // For multi-row, single-column ranges, use autofill approach
    if (rows > 1 && cols === 1) {
        const firstCell = range.getCell(0, 0);
        firstCell.formulas = [[formula]];
        
        try {
            firstCell.autoFill(range, Excel.AutoFillType.fillDefault);
            logDiag(`Autofilled formula to ${rows} rows`);
            return;
        } catch (autofillError) {
            console.warn("Autofill failed, using formula array:", autofillError);
            logDiag(`Autofill failed, building formula array`);
            
            const formulas = [];
            for (let r = 0; r < rows; r++) {
                let f = formula;
                if (r > 0) {
                    f = adjustFormulaReferences(formula, r, 0);
                }
                formulas.push([f]);
            }
            range.formulas = formulas;
            return;
        }
    }
    
    // For single-row, multi-column ranges
    if (rows === 1 && cols > 1) {
        const formulas = [[]];
        for (let c = 0; c < cols; c++) {
            let f = formula;
            if (c > 0) {
                f = adjustFormulaReferences(formula, 0, c);
            }
            formulas[0].push(f);
        }
        range.formulas = formulas;
        logDiag(`Applied formula to ${cols} columns`);
        return;
    }
    
    // For multi-row, multi-column ranges
    if (rows > 1 && cols > 1) {
        const formulas = [];
        for (let r = 0; r < rows; r++) {
            const rowFormulas = [];
            for (let c = 0; c < cols; c++) {
                let f = formula;
                if (r > 0 || c > 0) {
                    f = adjustFormulaReferences(formula, r, c);
                }
                rowFormulas.push(f);
            }
            formulas.push(rowFormulas);
        }
        range.formulas = formulas;
        logDiag(`Applied formula to ${rows}x${cols} range`);
        return;
    }
}

/**
 * Adjusts cell references in a formula for row/column offset
 * Supports multi-letter columns (AA, AB, etc.)
 * @param {string} formula - Original formula
 * @param {number} rowOffset - Row offset to apply
 * @param {number} colOffset - Column offset to apply
 * @returns {string} Adjusted formula
 */
function adjustFormulaReferences(formula, rowOffset, colOffset) {
    return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colAbs, col, rowAbs, row) => {
        let newCol = col;
        let newRow = parseInt(row);
        
        // Adjust column if not absolute and offset > 0
        if (colAbs !== "$" && colOffset > 0) {
            // Use robust base-26 conversion for multi-letter columns
            const colIndex = colLetterToIndex(col);
            newCol = colIndexToLetter(colIndex + colOffset);
        }
        
        // Adjust row if not absolute and offset > 0
        if (rowAbs !== "$" && rowOffset > 0) {
            newRow = newRow + rowOffset;
        }
        
        return `${colAbs}${newCol}${rowAbs}${newRow}`;
    });
}

// ============================================================================
// Values Application
// ============================================================================

/**
 * Applies values to a range
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string of values
 */
function applyValues(range, data) {
    let values;
    try {
        values = JSON.parse(data);
        if (!Array.isArray(values)) values = [[values]];
        if (!Array.isArray(values[0])) values = [values];
    } catch {
        values = [[data]];
    }
    range.values = values;
    logDiag(`Applied values to range`);
}

// ============================================================================
// Formatting
// ============================================================================

/**
 * Applies formatting to a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string of format options
 */
async function applyFormat(ctx, range, data) {
    let fmt;
    try { fmt = JSON.parse(data); } catch { fmt = {}; }
    
    if (fmt.bold) range.format.font.bold = true;
    if (fmt.italic) range.format.font.italic = true;
    if (fmt.fill) range.format.fill.color = fmt.fill;
    if (fmt.fontColor) range.format.font.color = fmt.fontColor;
    if (fmt.fontSize) range.format.font.size = fmt.fontSize;
    if (fmt.numberFormat) range.numberFormat = [[fmt.numberFormat]];
    if (fmt.border) {
        range.format.borders.getItem("EdgeTop").style = "Continuous";
        range.format.borders.getItem("EdgeBottom").style = "Continuous";
        range.format.borders.getItem("EdgeLeft").style = "Continuous";
        range.format.borders.getItem("EdgeRight").style = "Continuous";
    }
    logDiag(`Applied formatting: ${Object.keys(fmt).join(", ")}`);
}

/**
 * Applies conditional formatting to a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string of conditional format rules
 */
async function applyConditionalFormat(ctx, range, data) {
    let rules;
    try { 
        const parsed = JSON.parse(data);
        rules = Array.isArray(parsed) ? parsed : [parsed];
    } catch { 
        rules = []; 
        logDiag(`Failed to parse conditional format rules`);
    }
    
    range.conditionalFormats.clearAll();
    await ctx.sync();
    
    for (const rule of rules) {
        if (rule.type === "cellValue" && rule.operator && rule.value) {
            const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cf.cellValue.format.fill.color = rule.fill || "#FFFF00";
            
            let operator = rule.operator;
            if (operator === "GreaterThan") operator = Excel.ConditionalCellValueOperator.greaterThan;
            else if (operator === "LessThan") operator = Excel.ConditionalCellValueOperator.lessThan;
            else if (operator === "EqualTo") operator = Excel.ConditionalCellValueOperator.equalTo;
            else if (operator === "GreaterThanOrEqual") operator = Excel.ConditionalCellValueOperator.greaterThanOrEqual;
            else if (operator === "LessThanOrEqual") operator = Excel.ConditionalCellValueOperator.lessThanOrEqual;
            else if (operator === "Between") operator = Excel.ConditionalCellValueOperator.between;
            
            cf.cellValue.rule = {
                formula1: String(rule.value),
                formula2: rule.value2 ? String(rule.value2) : undefined,
                operator: operator
            };
        }
    }
    logDiag(`Applied ${rules.length} conditional format rule(s)`);
}

/**
 * Clears conditional formatting from a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 */
async function clearConditionalFormat(ctx, range) {
    range.conditionalFormats.clearAll();
    await ctx.sync();
    logDiag(`Cleared conditional formatting`);
}

// ============================================================================
// Validation
// ============================================================================

/**
 * Applies data validation (dropdown) to a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} range - Target range
 * @param {string} source - Source range for dropdown values
 */
async function applyValidation(ctx, sheet, range, source) {
    if (source) {
        range.dataValidation.clear();
        await ctx.sync();
        
        const sourceRange = sheet.getRange(source);
        sourceRange.load("values");
        await ctx.sync();
        
        const uniqueValues = [];
        const seen = new Set();
        for (const row of sourceRange.values) {
            const val = row[0];
            if (val !== null && val !== undefined && val !== "" && !seen.has(val)) {
                seen.add(val);
                uniqueValues.push(String(val));
            }
        }
        
        const listSource = uniqueValues.join(",");
        
        range.dataValidation.rule = {
            list: {
                inCellDropDown: true,
                source: listSource
            }
        };
        logDiag(`Applied validation with ${uniqueValues.length} options`);
    }
}

// ============================================================================
// Charts
// ============================================================================

/**
 * Creates a chart from data
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} dataRange - Data range for chart
 * @param {Object} action - Chart action with options
 */
async function createChart(ctx, sheet, dataRange, action) {
    const { chartType, data } = action;
    const ct = (chartType || "column").toLowerCase();
    
    dataRange.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await ctx.sync();
    
    const values = dataRange.values;
    const headers = values[0];
    const rowCount = dataRange.rowCount;
    
    let title = action.title || "Chart";
    let position = action.position || "H2";
    
    // Smart aggregation detection
    let shouldAggregate = false;
    let categoryCol = -1;
    let valueCol = -1;
    
    if (rowCount > 10 && headers.length >= 2) {
        for (let c = 0; c < headers.length; c++) {
            const sample = values.slice(1, Math.min(6, values.length)).map(r => r[c]);
            const hasText = sample.some(v => typeof v === "string" && v.length > 0);
            const hasRepeats = new Set(sample).size < sample.length;
            if (hasText && hasRepeats) {
                categoryCol = c;
                break;
            }
        }
        
        for (let c = 0; c < headers.length; c++) {
            if (c === categoryCol) continue;
            
            const header = String(headers[c] || "").toLowerCase();
            const sample = values.slice(1, Math.min(10, values.length)).map(r => r[c]);
            const hasNumbers = sample.every(v => typeof v === "number" || !isNaN(parseFloat(v)));
            
            if (!hasNumbers) continue;
            
            const isID = header.includes("id") || header.includes("no") || header.includes("number");
            const numericSample = sample.map(v => parseFloat(v)).filter(v => !isNaN(v));
            const isSequential = numericSample.length > 3 && 
                numericSample.every((v, i) => i === 0 || v > numericSample[i-1]);
            const isUnique = new Set(numericSample).size === numericSample.length;
            
            if (isID || (isSequential && isUnique)) continue;
            
            valueCol = c;
            break;
        }
        
        shouldAggregate = categoryCol !== -1;
    }
    
    let chartDataRange = dataRange;
    
    if (shouldAggregate) {
        const aggregated = {};
        for (let r = 1; r < values.length; r++) {
            const key = String(values[r][categoryCol] || "").trim();
            if (!key) continue;
            if (!aggregated[key]) aggregated[key] = { count: 0, sum: 0 };
            aggregated[key].count++;
            if (valueCol !== -1) {
                const val = parseFloat(values[r][valueCol]);
                if (!isNaN(val)) aggregated[key].sum += val;
            }
        }
        
        const aggData = Object.entries(aggregated)
            .map(([key, data]) => [key, valueCol !== -1 ? data.sum : data.count])
            .sort((a, b) => b[1] - a[1]);
        
        const aggStartRow = dataRange.rowIndex + rowCount + 2;
        const aggValues = [[headers[categoryCol] || "Category", valueCol !== -1 ? headers[valueCol] : "Count"], ...aggData];
        const aggRange = sheet.getRangeByIndexes(aggStartRow, dataRange.columnIndex, aggValues.length, 2);
        aggRange.values = aggValues;
        await ctx.sync();
        
        chartDataRange = aggRange;
        logDiag(`Aggregated data for chart: ${aggData.length} categories`);
    }
    
    // Determine chart type
    let type = Excel.ChartType.columnClustered;
    
    if (ct.includes("line")) type = Excel.ChartType.line;
    else if (ct.includes("pie")) type = Excel.ChartType.pie;
    else if (ct.includes("doughnut") || ct.includes("donut")) type = Excel.ChartType.doughnut;
    else if (ct.includes("bar")) type = Excel.ChartType.barClustered;
    else if (ct.includes("area")) type = Excel.ChartType.area;
    else if (ct.includes("scatter") || ct.includes("xy")) type = Excel.ChartType.xyscatter;
    else if (ct.includes("radar") || ct.includes("spider")) type = Excel.ChartType.radar;
    else if (ct.includes("stacked")) {
        type = ct.includes("bar") ? Excel.ChartType.barStacked : Excel.ChartType.columnStacked;
    }
    
    // Handle non-contiguous ranges
    const targetAddress = action.target;
    if (targetAddress && targetAddress.includes(",")) {
        console.warn("Non-contiguous ranges not fully supported for charts, using first range");
        logDiag(`Warning: Non-contiguous range "${targetAddress}" - using first range only`);
        const ranges = targetAddress.split(",").map(r => r.trim());
        chartDataRange = sheet.getRange(ranges[0]);
    }
    
    const chart = sheet.charts.add(type, chartDataRange, Excel.ChartSeriesBy.auto);
    
    const startCol = position.match(/[A-Z]+/)?.[0] || "H";
    const startRow = parseInt(position.match(/\d+/)?.[0] || "2");
    const endCol = String.fromCharCode(startCol.charCodeAt(0) + 8);
    const endRow = startRow + 15;
    
    chart.setPosition(position, `${endCol}${endRow}`);
    chart.title.text = title;
    chart.title.visible = true;
    chart.legend.visible = true;
    chart.legend.position = (ct.includes("pie") || ct.includes("doughnut")) 
        ? Excel.ChartLegendPosition.right 
        : Excel.ChartLegendPosition.bottom;
    
    logDiag(`Created ${ct} chart at ${position}`);
}

/**
 * Creates a pivot chart with aggregation
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} range - Data range
 * @param {Object} action - Pivot chart action with options
 */
async function createPivotChart(ctx, sheet, range, action) {
    range.load(["values", "rowIndex", "columnIndex", "rowCount"]);
    await ctx.sync();
    
    const values = range.values;
    const headers = values[0];
    
    let options = { groupBy: null, aggregate: null, aggregateFunc: "sum", chartType: "column", title: "Pivot Chart", position: "H2" };
    if (action.data) {
        try { options = { ...options, ...JSON.parse(action.data) }; } catch (e) {}
    }
    if (action.chartType) options.chartType = action.chartType;
    if (action.title) options.title = action.title;
    if (action.position) options.position = action.position;
    
    let groupByIdx = -1;
    for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i]).toLowerCase().trim();
        const searchTerm = String(options.groupBy || "").toLowerCase().trim();
        if (header === searchTerm || header.includes(searchTerm) || searchTerm.includes(header)) {
            groupByIdx = i;
            break;
        }
    }
    
    if (groupByIdx === -1) {
        logDiag(`Pivot chart failed: Column "${options.groupBy}" not found`);
        throw new Error(`Column "${options.groupBy}" not found. Available: ${headers.join(", ")}`);
    }
    
    let aggregateIdx = -1;
    if (options.aggregate) {
        for (let i = 0; i < headers.length; i++) {
            const header = String(headers[i]).toLowerCase().trim();
            const searchTerm = String(options.aggregate).toLowerCase().trim();
            if (header === searchTerm || header.includes(searchTerm) || searchTerm.includes(header)) {
                aggregateIdx = i;
                break;
            }
        }
    }
    
    const aggregated = {};
    for (let r = 1; r < values.length; r++) {
        const groupValue = values[r][groupByIdx];
        const key = String(groupValue || "").trim();
        if (!key || key === "null" || key === "undefined") continue;
        
        if (!aggregated[key]) aggregated[key] = { count: 0, sum: 0, values: [] };
        aggregated[key].count++;
        
        if (aggregateIdx !== -1) {
            const val = parseFloat(values[r][aggregateIdx]);
            if (!isNaN(val)) {
                aggregated[key].sum += val;
                aggregated[key].values.push(val);
            }
        }
    }
    
    const chartData = [];
    for (const [key, data] of Object.entries(aggregated)) {
        let value;
        const func = (options.aggregateFunc || "count").toLowerCase();
        switch (func) {
            case "count": value = data.count; break;
            case "average": case "avg": value = data.values.length > 0 ? data.sum / data.values.length : data.count; break;
            case "max": value = data.values.length > 0 ? Math.max(...data.values) : data.count; break;
            case "min": value = data.values.length > 0 ? Math.min(...data.values) : data.count; break;
            case "sum": default: value = data.values.length > 0 ? data.sum : data.count; break;
        }
        chartData.push([key, value]);
    }
    chartData.sort((a, b) => b[1] - a[1]);
    
    const chartStartRow = range.rowIndex + range.rowCount + 2;
    const chartValues = [[options.groupBy || "Category", options.aggregate || "Value"], ...chartData];
    const chartDataRange = sheet.getRangeByIndexes(chartStartRow, range.columnIndex, chartValues.length, 2);
    chartDataRange.values = chartValues;
    await ctx.sync();
    
    let type = Excel.ChartType.columnClustered;
    const ct = options.chartType.toLowerCase();
    if (ct.includes("pie")) type = Excel.ChartType.pie;
    else if (ct.includes("bar")) type = Excel.ChartType.barClustered;
    else if (ct.includes("line")) type = Excel.ChartType.line;
    
    const chart = sheet.charts.add(type, chartDataRange, Excel.ChartSeriesBy.columns);
    const position = options.position || "H2";
    const startCol = position.match(/[A-Z]+/)?.[0] || "H";
    const startRow = parseInt(position.match(/\d+/)?.[0] || "2");
    chart.setPosition(position, `${String.fromCharCode(startCol.charCodeAt(0) + 8)}${startRow + 15}`);
    chart.title.text = options.title;
    chart.legend.visible = true;
    chart.legend.position = ct.includes("pie") ? Excel.ChartLegendPosition.right : Excel.ChartLegendPosition.bottom;
    await ctx.sync();
    
    logDiag(`Created pivot chart: ${options.title}`);
}

// ============================================================================
// Sorting and Filtering
// ============================================================================

/**
 * Applies sorting to a range
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string of sort options
 */
function applySort(range, data) {
    let opts = {};
    
    if (typeof data === "string") {
        try {
            opts = JSON.parse(data);
        } catch {
            const parts = data.split(",");
            for (const part of parts) {
                const [key, value] = part.split(":").map(s => s.trim());
                if (key === "column") opts.column = parseInt(value) || 0;
                if (key === "ascending") opts.ascending = value !== "false";
                if (key === "hasHeaders") opts.hasHeaders = value === "true";
            }
        }
    } else {
        opts = data || {};
    }
    
    const columnIndex = opts.column || 0;
    const ascending = opts.ascending !== false;
    const hasHeaders = opts.hasHeaders !== false;
    
    range.sort.apply(
        [{ key: columnIndex, ascending: ascending }],
        false,
        hasHeaders,
        Excel.SortOrientation.rows
    );
    logDiag(`Sorted by column ${columnIndex}, ascending: ${ascending}`);
}

/**
 * Applies AutoFilter to a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string of filter options
 */
async function applyFilter(ctx, sheet, range, data) {
    let filterOpts = {};
    
    if (typeof data === "string") {
        try {
            filterOpts = JSON.parse(data);
        } catch {
            logDiag(`Invalid filter data format`);
            throw new Error("Invalid filter data format");
        }
    } else {
        filterOpts = data || {};
    }
    
    try {
        sheet.autoFilter.clearCriteria();
        await ctx.sync();
    } catch (e) {
        // No existing filter
    }
    
    sheet.autoFilter.apply(range);
    await ctx.sync();
    
    if (filterOpts.column !== undefined && filterOpts.values) {
        const criteria = {
            filterOn: Excel.FilterOn.values,
            values: filterOpts.values
        };
        sheet.autoFilter.apply(range, filterOpts.column, criteria);
        await ctx.sync();
        logDiag(`Applied filter on column ${filterOpts.column}`);
    }
}

/**
 * Clears all filters from the worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 */
async function clearFilter(ctx, sheet) {
    try {
        sheet.autoFilter.clearCriteria();
        await ctx.sync();
        logDiag(`Cleared filters`);
    } catch (e) {
        // No filter to clear
    }
}

// ============================================================================
// Copy and Sheet Operations
// ============================================================================

/**
 * Copies data from source to target range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {string} source - Source range address
 * @param {string} target - Target range address
 */
async function applyCopy(ctx, sheet, source, target) {
    if (!source || !target) {
        throw new Error("Copy requires both source and target ranges");
    }
    
    const sourceRange = sheet.getRange(source);
    sourceRange.load(["values", "formulas", "rowCount", "columnCount"]);
    await ctx.sync();
    
    const rowCount = sourceRange.rowCount;
    const colCount = sourceRange.columnCount;
    
    const targetCell = sheet.getRange(target);
    const targetRange = targetCell.getResizedRange(rowCount - 1, colCount - 1);
    targetRange.formulas = sourceRange.formulas;
    logDiag(`Copied ${source} to ${target}`);
}

/**
 * Copies only values (not formulas) from source to target
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {string} source - Source range address
 * @param {string} target - Target range address
 */
async function applyCopyValues(ctx, sheet, source, target) {
    if (!source || !target) {
        throw new Error("Copy requires both source and target ranges");
    }
    
    const sourceRange = sheet.getRange(source);
    sourceRange.load(["values", "rowCount", "columnCount"]);
    await ctx.sync();
    
    const rowCount = sourceRange.rowCount;
    const colCount = sourceRange.columnCount;
    
    const targetAddress = target.includes(":") ? target.split(":")[0] : target;
    const targetCell = sheet.getRange(targetAddress);
    const targetRange = targetCell.getResizedRange(rowCount - 1, colCount - 1);
    targetRange.values = sourceRange.values;
    logDiag(`Copied values from ${source} to ${target}`);
}

/**
 * Creates a new worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {string} sheetName - Name for the new sheet
 * @param {string} data - Optional JSON data to populate
 */
async function createSheet(ctx, sheetName, data) {
    if (!sheetName) {
        throw new Error("Sheet name is required");
    }
    
    const sheets = ctx.workbook.worksheets;
    const newSheet = sheets.add();
    newSheet.name = sheetName;
    
    if (data) {
        try {
            const values = JSON.parse(data);
            if (Array.isArray(values) && values.length > 0) {
                const range = newSheet.getRange(`A1:${colIndexToLetter(values[0].length - 1)}${values.length}`);
                range.values = values;
            }
        } catch (e) {
            // Data parsing failed
        }
    }
    logDiag(`Created sheet: ${sheetName}`);
}

/**
 * Removes duplicate rows from a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string with columns array
 */
async function removeDuplicates(ctx, range, data) {
    range.load(["values", "rowCount", "columnCount", "address"]);
    await ctx.sync();
    
    const values = range.values;
    const rowCount = range.rowCount;
    const colCount = range.columnCount;
    const rangeAddress = range.address;
    
    let options = { columns: [] };
    if (data) {
        try {
            options = JSON.parse(data);
        } catch (e) {
            options.columns = Array.from({ length: colCount }, (_, i) => i);
        }
    }
    
    if (!options.columns || options.columns.length === 0) {
        options.columns = Array.from({ length: colCount }, (_, i) => i);
    }
    
    const seen = new Set();
    const uniqueRows = [];
    
    for (let r = 0; r < rowCount; r++) {
        const row = values[r];
        const key = options.columns.map(colIdx => {
            const val = row[colIdx];
            return val === null || val === undefined ? "" : String(val);
        }).join("|");
        
        if (!seen.has(key)) {
            seen.add(key);
            uniqueRows.push(row);
        }
    }
    
    const removedCount = rowCount - uniqueRows.length;
    logDiag(`Removing ${removedCount} duplicate rows`);
    
    range.clear(Excel.ClearApplyTo.contents);
    await ctx.sync();
    
    if (uniqueRows.length > 0) {
        const sheet = range.worksheet;
        const address = rangeAddress.split("!")[1] || rangeAddress;
        const startCell = address.split(":")[0];
        
        const targetCell = sheet.getRange(startCell);
        const newRange = targetCell.getResizedRange(uniqueRows.length - 1, colCount - 1);
        newRange.values = uniqueRows;
    }
}

// ============================================================================
// Table Operations
// ============================================================================

/**
 * Valid table styles for validation
 */
const VALID_TABLE_STYLES = [
    // Light styles (1-21)
    ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
    // Medium styles (1-28)
    ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
    // Dark styles (1-11)
    ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`)
];

/**
 * Creates an Excel Table from a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} range - Data range for the table
 * @param {Object} action - Action with table options
 */
async function createTable(ctx, sheet, range, action) {
    logDiag(`Starting createTable at range "${action.target}"`);
    
    let options = { tableName: null, style: "TableStyleMedium2" };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for createTable, using defaults`);
        }
    }
    
    // Validate style with clear error message
    if (options.style && !VALID_TABLE_STYLES.includes(options.style)) {
        logDiag(`Warning: Invalid table style "${options.style}". Valid styles: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11. Using TableStyleMedium2.`);
        options.style = "TableStyleMedium2";
    }
    
    try {
        // Create table with headers (true = first row is header)
        const table = sheet.tables.add(range, true);
        
        // Set table name if provided
        if (options.tableName) {
            table.name = options.tableName;
        }
        
        // Apply style
        table.style = options.style;
        
        // Enable default table features
        table.showBandedRows = true;
        table.showFilterButton = true;
        
        await ctx.sync();
        
        const tableName = options.tableName || table.name;
        logDiag(`Successfully created table "${tableName}" at ${action.target} with style ${options.style}`);
    } catch (e) {
        const errorMsg = e.message && e.message.includes("already") 
            ? `Failed to create table: Range ${action.target} already contains a table or overlaps with one.`
            : `Failed to create table at ${action.target}: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Applies styling to an existing table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {string} tableName - Name of the table to style
 * @param {string} data - JSON string with style options
 */
async function styleTable(ctx, sheet, tableName, data) {
    logDiag(`Starting styleTable for table "${tableName}"`);
    
    let options = { style: "TableStyleMedium2" };
    
    if (data) {
        try {
            options = { ...options, ...JSON.parse(data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse data for styleTable, using defaults`);
        }
    }
    
    // Validate style with clear error message
    if (options.style && !VALID_TABLE_STYLES.includes(options.style)) {
        logDiag(`Warning: Invalid table style "${options.style}". Valid styles: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11. Using TableStyleMedium2.`);
        options.style = "TableStyleMedium2";
    }
    
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Apply style
        table.style = options.style;
        
        // Apply additional style options if provided
        if (options.highlightFirstColumn !== undefined) {
            table.highlightFirstColumn = options.highlightFirstColumn;
        }
        if (options.highlightLastColumn !== undefined) {
            table.highlightLastColumn = options.highlightLastColumn;
        }
        if (options.showBandedRows !== undefined) {
            table.showBandedRows = options.showBandedRows;
        }
        if (options.showBandedColumns !== undefined) {
            table.showBandedColumns = options.showBandedColumns;
        }
        
        await ctx.sync();
        logDiag(`Successfully applied style "${options.style}" to table "${tableName}"`);
    } catch (e) {
        const errorMsg = `Failed to style table "${tableName}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Adds a row to an existing table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with row options
 */
async function addTableRow(ctx, sheet, action) {
    logDiag(`Starting addTableRow for target "${action.target}"`);
    
    let options = { tableName: action.target, position: "end", values: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for addTableRow, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    // Determine index: null for end, 0 for start, or specific number
    let index = null;
    if (options.position === "start" || options.position === 0) {
        index = 0;
    } else if (typeof options.position === "number" && options.position > 0) {
        index = options.position;
    }
    // null means append to end
    
    // Prepare values - should be array of arrays
    let rowValues = null;
    if (options.values) {
        rowValues = Array.isArray(options.values[0]) ? options.values : [options.values];
    }
    
    try {
        table.rows.add(index, rowValues);
        await ctx.sync();
        logDiag(`Successfully added row to table "${tableName}" at position ${options.position || "end"}`);
    } catch (e) {
        const errorMsg = `Failed to add row to table "${tableName}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Adds a column to an existing table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with column options
 */
async function addTableColumn(ctx, sheet, action) {
    logDiag(`Starting addTableColumn for target "${action.target}"`);
    
    let options = { tableName: action.target, columnName: "NewColumn", position: "end", values: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for addTableColumn, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    // Determine index: null for end, 0 for start, or specific number
    let index = null;
    if (options.position === "start" || options.position === 0) {
        index = 0;
    } else if (typeof options.position === "number" && options.position > 0) {
        index = options.position;
    }
    
    // Prepare values - should include header as first element
    let columnValues = null;
    if (options.values) {
        columnValues = Array.isArray(options.values[0]) ? options.values : options.values.map(v => [v]);
    } else if (options.columnName) {
        // Just add header if no values provided
        columnValues = [[options.columnName]];
    }
    
    try {
        table.columns.add(index, columnValues);
        await ctx.sync();
        logDiag(`Successfully added column "${options.columnName}" to table "${tableName}" at position ${options.position || "end"}`);
    } catch (e) {
        const errorMsg = `Failed to add column to table "${tableName}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Resizes an existing table to a new range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with resize options
 */
async function resizeTable(ctx, sheet, action) {
    logDiag(`Starting resizeTable for target "${action.target}"`);
    
    let options = { tableName: action.target, newRange: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for resizeTable, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    
    if (!options.newRange) {
        const errorMsg = `newRange is required for resizeTable operation on table "${tableName}".`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    // Get current range for logging
    const currentRange = table.getRange();
    currentRange.load("address");
    await ctx.sync();
    
    const oldAddress = currentRange.address;
    
    try {
        // Resize the table
        table.resize(options.newRange);
        await ctx.sync();
        logDiag(`Successfully resized table "${tableName}" from ${oldAddress} to ${options.newRange}`);
    } catch (e) {
        const errorMsg = `Failed to resize table "${tableName}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Converts a table back to a normal range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {string} tableName - Name of the table to convert
 */
async function convertToRange(ctx, sheet, tableName) {
    logDiag(`Starting convertToRange for table "${tableName}"`);
    
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Convert table to range - preserves data and formatting
        table.convertToRange();
        await ctx.sync();
        logDiag(`Successfully converted table "${tableName}" to normal range`);
    } catch (e) {
        const errorMsg = `Failed to convert table "${tableName}" to range: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Gets valid totals calculation functions map
 * Note: Must be called inside Excel.run context, not at module load time
 */
function getValidTotalsFunctions() {
    return {
        "sum": Excel.TotalsCalculation.sum,
        "average": Excel.TotalsCalculation.average,
        "avg": Excel.TotalsCalculation.average,
        "count": Excel.TotalsCalculation.count,
        "countnumbers": Excel.TotalsCalculation.countNumbers,
        "max": Excel.TotalsCalculation.max,
        "min": Excel.TotalsCalculation.min,
        "stddev": Excel.TotalsCalculation.stdDev,
        "var": Excel.TotalsCalculation.var,
        "none": Excel.TotalsCalculation.none
    };
}

/**
 * Toggles the total row for a table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with totals options
 */
async function toggleTableTotals(ctx, sheet, action) {
    logDiag(`Starting toggleTableTotals for target "${action.target}"`);
    
    let options = { tableName: action.target, show: true, totals: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for toggleTableTotals, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    // Toggle totals row visibility
    table.showTotals = options.show;
    await ctx.sync();
    
    logDiag(`Set showTotals=${options.show} for table "${tableName}"`);
    
    // If enabling totals and specific functions are requested, apply them
    const appliedFunctions = [];
    if (options.show && options.totals && Array.isArray(options.totals) && options.totals.length > 0) {
        table.columns.load("count");
        await ctx.sync();
        
        const columnCount = table.columns.count;
        
        for (const totalConfig of options.totals) {
            // Validate columnIndex
            if (totalConfig.columnIndex === undefined || totalConfig.columnIndex === null) {
                logDiag(`Warning: Skipping totals config - missing columnIndex`);
                continue;
            }
            
            if (typeof totalConfig.columnIndex !== "number" || totalConfig.columnIndex < 0) {
                logDiag(`Warning: Skipping totals config - invalid columnIndex "${totalConfig.columnIndex}"`);
                continue;
            }
            
            if (totalConfig.columnIndex >= columnCount) {
                logDiag(`Warning: Skipping totals config - columnIndex ${totalConfig.columnIndex} exceeds table column count ${columnCount}`);
                continue;
            }
            
            // Validate function name
            if (!totalConfig.function) {
                logDiag(`Warning: Skipping totals config for column ${totalConfig.columnIndex} - missing function`);
                continue;
            }
            
            const funcName = String(totalConfig.function).toLowerCase().replace(/\s/g, "");
            const validFunctions = getValidTotalsFunctions();
            const calcFunc = validFunctions[funcName];
            
            if (!calcFunc) {
                logDiag(`Warning: Invalid totals function "${totalConfig.function}" for column ${totalConfig.columnIndex}. Valid functions: Sum, Average, Count, Max, Min, StdDev, Var, None`);
                continue;
            }
            
            try {
                const column = table.columns.getItemAt(totalConfig.columnIndex);
                column.totalsCalculation = calcFunc;
                appliedFunctions.push(`column ${totalConfig.columnIndex}: ${totalConfig.function}`);
            } catch (e) {
                logDiag(`Warning: Failed to apply ${totalConfig.function} to column ${totalConfig.columnIndex}: ${e.message}`);
            }
        }
        
        await ctx.sync();
    }
    
    if (appliedFunctions.length > 0) {
        logDiag(`Applied totals functions for table "${tableName}": ${appliedFunctions.join(", ")}`);
    }
    
    logDiag(`Completed toggleTableTotals for table "${tableName}": show=${options.show}`);
}

// ============================================================================
// Data Manipulation Operations
// ============================================================================

/**
 * Inserts rows at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with row options
 */
async function insertRows(ctx, sheet, action) {
    logDiag(`Starting insertRows for target "${action.target}"`);
    
    let options = { count: 1 };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for insertRows, using defaults`);
        }
    }
    
    // Validate target is a row range (e.g., "5" or "5:7")
    const rowPattern = /^(\d+)(:\d+)?$/;
    if (!rowPattern.test(action.target)) {
        const errorMsg = `Invalid row range "${action.target}". Use format "5" or "5:7".`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    // Validate count
    if (typeof options.count !== "number" || options.count < 1) {
        logDiag(`Warning: Invalid count "${options.count}", using 1`);
        options.count = 1;
    }
    
    try {
        const range = sheet.getRange(action.target);
        const entireRow = range.getEntireRow();
        
        // Insert rows multiple times if count > 1
        for (let i = 0; i < options.count; i++) {
            entireRow.insert(Excel.InsertShiftDirection.down);
        }
        
        await ctx.sync();
        logDiag(`Successfully inserted ${options.count} row(s) at row ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to insert rows at "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Inserts columns at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with column options
 */
async function insertColumns(ctx, sheet, action) {
    logDiag(`Starting insertColumns for target "${action.target}"`);
    
    let options = { count: 1 };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for insertColumns, using defaults`);
        }
    }
    
    // Validate target is a column range (e.g., "C" or "C:E")
    const colPattern = /^([A-Z]+)(:[A-Z]+)?$/i;
    if (!colPattern.test(action.target)) {
        const errorMsg = `Invalid column range "${action.target}". Use format "C" or "C:E".`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    // Validate count
    if (typeof options.count !== "number" || options.count < 1) {
        logDiag(`Warning: Invalid count "${options.count}", using 1`);
        options.count = 1;
    }
    
    try {
        const range = sheet.getRange(action.target);
        const entireColumn = range.getEntireColumn();
        
        // Insert columns multiple times if count > 1
        for (let i = 0; i < options.count; i++) {
            entireColumn.insert(Excel.InsertShiftDirection.right);
        }
        
        await ctx.sync();
        logDiag(`Successfully inserted ${options.count} column(s) at column ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to insert columns at "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes rows at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with row options
 */
async function deleteRows(ctx, sheet, action) {
    logDiag(`Starting deleteRows for target "${action.target}"`);
    
    // Validate target is a row range (e.g., "10" or "10:15")
    const rowPattern = /^(\d+)(:\d+)?$/;
    if (!rowPattern.test(action.target)) {
        const errorMsg = `Invalid row range "${action.target}". Use format "10" or "10:15".`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        const range = sheet.getRange(action.target);
        const entireRow = range.getEntireRow();
        entireRow.delete(Excel.DeleteShiftDirection.up);
        
        await ctx.sync();
        logDiag(`Successfully deleted row(s) at ${action.target}. Warning: This may affect formula references.`);
    } catch (e) {
        const errorMsg = `Failed to delete rows at "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes columns at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with column options
 */
async function deleteColumns(ctx, sheet, action) {
    logDiag(`Starting deleteColumns for target "${action.target}"`);
    
    // Validate target is a column range (e.g., "D" or "D:F")
    const colPattern = /^([A-Z]+)(:[A-Z]+)?$/i;
    if (!colPattern.test(action.target)) {
        const errorMsg = `Invalid column range "${action.target}". Use format "D" or "D:F".`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        const range = sheet.getRange(action.target);
        const entireColumn = range.getEntireColumn();
        entireColumn.delete(Excel.DeleteShiftDirection.left);
        
        await ctx.sync();
        logDiag(`Successfully deleted column(s) at ${action.target}. Warning: This may affect formula references.`);
    } catch (e) {
        const errorMsg = `Failed to delete columns at "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Merges cells in the specified range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with merge options
 */
async function mergeCells(ctx, sheet, action) {
    logDiag(`Starting mergeCells for target "${action.target}"`);
    
    try {
        const range = sheet.getRange(action.target);
        range.load(["address", "rowCount", "columnCount"]);
        await ctx.sync();
        
        // Validate range is at least 2 cells
        if (range.rowCount === 1 && range.columnCount === 1) {
            const errorMsg = `Cannot merge a single cell. Range must contain at least 2 cells.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        range.merge(false);
        await ctx.sync();
        
        logDiag(`Successfully merged cells at ${action.target}. Note: Only the top-left cell value is retained.`);
    } catch (e) {
        if (e.message && e.message.includes("merge")) {
            const errorMsg = `Failed to merge cells at "${action.target}": Range may already contain merged cells.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        const errorMsg = `Failed to merge cells at "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Unmerges cells in the specified range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with unmerge options
 */
async function unmergeCells(ctx, sheet, action) {
    logDiag(`Starting unmergeCells for target "${action.target}"`);
    
    try {
        const range = sheet.getRange(action.target);
        range.unmerge();
        await ctx.sync();
        
        logDiag(`Successfully unmerged cells at ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to unmerge cells at "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Finds and replaces text in the specified range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with find/replace options
 */
async function findReplace(ctx, sheet, action) {
    logDiag(`Starting findReplace for target "${action.target}"`);
    
    let options = { find: "", replace: "", matchCase: false, matchEntireCell: false };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for findReplace`);
        }
    }
    
    // Validate find string
    if (!options.find || options.find.length === 0) {
        const errorMsg = `Find string cannot be empty.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        const range = sheet.getRange(action.target);
        
        const searchCriteria = {
            completeMatch: options.matchEntireCell,
            matchCase: options.matchCase
        };
        
        range.replaceAll(options.find, options.replace || "", searchCriteria);
        await ctx.sync();
        
        logDiag(`Successfully replaced "${options.find}" with "${options.replace}" in ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to find/replace in "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Splits text in a column into multiple columns based on delimiter
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with text-to-columns options
 */
async function textToColumns(ctx, sheet, action) {
    logDiag(`Starting textToColumns for target "${action.target}"`);
    
    let options = { delimiter: ",", destination: null };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for textToColumns, using defaults`);
        }
    }
    
    try {
        const sourceRange = sheet.getRange(action.target);
        sourceRange.load(["values", "rowCount", "columnCount", "columnIndex", "rowIndex"]);
        await ctx.sync();
        
        // Validate source is single column
        if (sourceRange.columnCount !== 1) {
            const errorMsg = `Text to columns requires a single-column range. Got ${sourceRange.columnCount} columns.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        const values = sourceRange.values;
        const delimiter = options.delimiter || ",";
        
        // Split each cell value
        const splitData = [];
        let maxColumns = 0;
        
        for (const row of values) {
            const cellValue = row[0];
            const parts = cellValue !== null && cellValue !== undefined 
                ? String(cellValue).split(delimiter) 
                : [""];
            splitData.push(parts);
            maxColumns = Math.max(maxColumns, parts.length);
        }
        
        // Pad arrays to same length
        for (const row of splitData) {
            while (row.length < maxColumns) {
                row.push("");
            }
        }
        
        // Determine destination
        let destRange;
        if (options.destination) {
            destRange = sheet.getRange(options.destination);
            destRange = destRange.getResizedRange(splitData.length - 1, maxColumns - 1);
        } else {
            // Use columns immediately to the right of source
            const destStartCol = sourceRange.columnIndex + 1;
            destRange = sheet.getRangeByIndexes(
                sourceRange.rowIndex,
                destStartCol,
                splitData.length,
                maxColumns
            );
        }
        
        // Check for existing data in destination range (Comment 4 safeguard)
        destRange.load("values");
        await ctx.sync();
        
        const existingValues = destRange.values;
        let nonEmptyCellCount = 0;
        for (const row of existingValues) {
            for (const cell of row) {
                if (cell !== null && cell !== undefined && cell !== "") {
                    nonEmptyCellCount++;
                }
            }
        }
        
        if (nonEmptyCellCount > 0 && !options.forceOverwrite) {
            const errorMsg = `Destination range contains ${nonEmptyCellCount} non-empty cell(s). Set "forceOverwrite": true in data to overwrite existing data, or choose a different destination.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        if (nonEmptyCellCount > 0 && options.forceOverwrite) {
            logDiag(`Warning: Overwriting ${nonEmptyCellCount} non-empty cell(s) in destination range (forceOverwrite=true)`);
        }
        
        destRange.values = splitData;
        await ctx.sync();
        
        logDiag(`Successfully split ${values.length} cells into ${maxColumns} columns.`);
    } catch (e) {
        const errorMsg = `Failed to split text to columns for "${action.target}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// PivotTable Operations
// ============================================================================

/**
 * Creates a PivotTable from a data range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with PivotTable options
 */
async function createPivotTable(ctx, sheet, action) {
    logDiag(`Starting createPivotTable for target "${action.target}"`);
    
    let options = { name: null, destination: null, layout: null };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for createPivotTable`);
        }
    }
    
    // Validate required fields
    if (!options.name) {
        const errorMsg = `PivotTable name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.destination) {
        const errorMsg = `Destination is required (e.g., "PivotSheet!A1").`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Parse destination into sheet name and cell
        let destSheetName, destCell;
        if (options.destination.includes("!")) {
            const parts = options.destination.split("!");
            destSheetName = parts[0];
            destCell = parts[1];
        } else {
            destSheetName = sheet.name;
            destCell = options.destination;
        }
        
        // Get or create destination sheet
        let destSheet = ctx.workbook.worksheets.getItemOrNullObject(destSheetName);
        await ctx.sync();
        
        if (destSheet.isNullObject) {
            logDiag(`Creating new sheet "${destSheetName}" for PivotTable`);
            destSheet = ctx.workbook.worksheets.add(destSheetName);
            await ctx.sync();
        }
        
        // Get source range - could be a range address or table name
        let sourceRange;
        const source = action.target;
        
        // Check if source is a table name (no colon in address)
        if (source && !source.includes(":") && !source.match(/^[A-Z]+\d+$/i)) {
            // Try to get as table
            const table = sheet.tables.getItemOrNullObject(source);
            table.load("isNullObject");
            await ctx.sync();
            
            if (!table.isNullObject) {
                sourceRange = table.getRange();
                logDiag(`Using table "${source}" as PivotTable source`);
            } else {
                sourceRange = sheet.getRange(source);
            }
        } else {
            sourceRange = sheet.getRange(source);
        }
        
        // Get destination range
        const destRange = destSheet.getRange(destCell);
        
        // Create PivotTable
        const pivotTable = destSheet.pivotTables.add(options.name, sourceRange, destRange);
        
        // Set layout if specified
        if (options.layout) {
            const layoutType = options.layout.toLowerCase();
            if (layoutType === "compact") {
                pivotTable.layout.layoutType = Excel.PivotLayoutType.compact;
            } else if (layoutType === "outline") {
                pivotTable.layout.layoutType = Excel.PivotLayoutType.outline;
            } else if (layoutType === "tabular") {
                pivotTable.layout.layoutType = Excel.PivotLayoutType.tabular;
            }
        }
        
        await ctx.sync();
        logDiag(`Successfully created PivotTable "${options.name}" from ${source} to ${options.destination}`);
    } catch (e) {
        const errorMsg = `Failed to create PivotTable: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Adds a field to a PivotTable
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with field options
 */
async function addPivotField(ctx, sheet, action) {
    logDiag(`Starting addPivotField for target "${action.target}"`);
    
    let options = { pivotName: action.target, field: null, area: null, function: "Sum" };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for addPivotField`);
        }
    }
    
    const pivotName = options.pivotName || action.target;
    
    // Validate required fields
    if (!options.field) {
        const errorMsg = `Field name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.area) {
        const errorMsg = `Area is required (row, column, data, or filter).`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for PivotTable in all sheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let pivotTable = null;
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${pivotName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Get the hierarchy for the field
        const hierarchy = pivotTable.hierarchies.getItem(options.field);
        
        // Add to appropriate area
        const area = options.area.toLowerCase();
        if (area === "row") {
            pivotTable.rowHierarchies.add(hierarchy);
            logDiag(`Added field "${options.field}" to row area of PivotTable "${pivotName}"`);
        } else if (area === "column") {
            pivotTable.columnHierarchies.add(hierarchy);
            logDiag(`Added field "${options.field}" to column area of PivotTable "${pivotName}"`);
        } else if (area === "data" || area === "value" || area === "values") {
            const dataHierarchy = pivotTable.dataHierarchies.add(hierarchy);
            
            // Set aggregation function with validation
            const rawFuncName = options.function || "Sum";
            const funcName = rawFuncName.toLowerCase().replace(/_/g, ""); // Normalize aliases like "count_numbers"
            
            const funcMap = {
                "sum": Excel.AggregationFunction.sum,
                "count": Excel.AggregationFunction.count,
                "average": Excel.AggregationFunction.average,
                "avg": Excel.AggregationFunction.average,  // Common alias
                "max": Excel.AggregationFunction.max,
                "min": Excel.AggregationFunction.min,
                "countnumbers": Excel.AggregationFunction.countNumbers,
                "stddev": Excel.AggregationFunction.standardDeviation,
                "stdev": Excel.AggregationFunction.standardDeviation,  // Common alias
                "standarddeviation": Excel.AggregationFunction.standardDeviation,
                "var": Excel.AggregationFunction.variance,
                "variance": Excel.AggregationFunction.variance
            };
            
            const supportedFunctions = "Sum, Count, Average, Max, Min, CountNumbers, StdDev, Var";
            
            if (funcMap[funcName]) {
                dataHierarchy.summarizeBy = funcMap[funcName];
                logDiag(`Added field "${options.field}" to data area with ${funcName} aggregation`);
            } else {
                // Invalid function - warn and fall back to Sum
                logDiag(`Warning: Unknown aggregation function "${rawFuncName}". Supported: ${supportedFunctions}. Falling back to Sum.`);
                dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
                logDiag(`Added field "${options.field}" to data area with Sum aggregation (fallback)`);
            }
        } else if (area === "filter") {
            pivotTable.filterHierarchies.add(hierarchy);
            logDiag(`Added field "${options.field}" to filter area of PivotTable "${pivotName}"`);
        } else {
            const errorMsg = `Invalid area "${options.area}". Use row, column, data, or filter.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        await ctx.sync();
    } catch (e) {
        const errorMsg = `Failed to add field "${options.field}": ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Configures the layout of a PivotTable
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with layout options
 */
async function configurePivotLayout(ctx, sheet, action) {
    logDiag(`Starting configurePivotLayout for target "${action.target}"`);
    
    let options = { pivotName: action.target, layout: null, showRowHeaders: null, showColumnHeaders: null };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for configurePivotLayout`);
        }
    }
    
    const pivotName = options.pivotName || action.target;
    
    if (!options.layout) {
        const errorMsg = `Layout type is required (Compact, Outline, or Tabular).`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for PivotTable in all sheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let pivotTable = null;
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${pivotName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Set layout type
        const layoutType = options.layout.toLowerCase();
        if (layoutType === "compact") {
            pivotTable.layout.layoutType = Excel.PivotLayoutType.compact;
        } else if (layoutType === "outline") {
            pivotTable.layout.layoutType = Excel.PivotLayoutType.outline;
        } else if (layoutType === "tabular") {
            pivotTable.layout.layoutType = Excel.PivotLayoutType.tabular;
        } else {
            const errorMsg = `Invalid layout type "${options.layout}". Use Compact, Outline, or Tabular.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Set header visibility if specified
        if (options.showRowHeaders !== null && options.showRowHeaders !== undefined) {
            pivotTable.layout.showRowHeaders = options.showRowHeaders;
        }
        if (options.showColumnHeaders !== null && options.showColumnHeaders !== undefined) {
            pivotTable.layout.showColumnHeaders = options.showColumnHeaders;
        }
        
        await ctx.sync();
        logDiag(`Successfully configured layout for PivotTable "${pivotName}" to ${options.layout}`);
    } catch (e) {
        const errorMsg = `Failed to configure layout: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Refreshes a PivotTable or all PivotTables
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with refresh options
 */
async function refreshPivotTable(ctx, sheet, action) {
    logDiag(`Starting refreshPivotTable for target "${action.target}"`);
    
    let options = { pivotName: action.target, refreshAll: false };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for refreshPivotTable`);
        }
    }
    
    try {
        if (options.refreshAll) {
            // Refresh all PivotTables in workbook
            ctx.workbook.pivotTables.refreshAll();
            await ctx.sync();
            logDiag(`Successfully refreshed all PivotTables in workbook`);
        } else {
            const pivotName = options.pivotName || action.target;
            
            // Search for PivotTable in all sheets
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            let pivotTable = null;
            for (const ws of sheets.items) {
                const pt = ws.pivotTables.getItemOrNullObject(pivotName);
                pt.load("isNullObject");
                await ctx.sync();
                
                if (!pt.isNullObject) {
                    pivotTable = pt;
                    break;
                }
            }
            
            if (!pivotTable) {
                const errorMsg = `PivotTable "${pivotName}" not found.`;
                logDiag(`Error: ${errorMsg}`);
                throw new Error(errorMsg);
            }
            
            pivotTable.refresh();
            await ctx.sync();
            logDiag(`Successfully refreshed PivotTable "${pivotName}"`);
        }
    } catch (e) {
        const errorMsg = `Failed to refresh PivotTable: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes a PivotTable
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with delete options
 */
async function deletePivotTable(ctx, sheet, action) {
    logDiag(`Starting deletePivotTable for target "${action.target}"`);
    
    let options = { pivotName: action.target };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for deletePivotTable`);
        }
    }
    
    const pivotName = options.pivotName || action.target;
    
    try {
        // Search for PivotTable in all sheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let pivotTable = null;
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${pivotName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        pivotTable.delete();
        await ctx.sync();
        logDiag(`Successfully deleted PivotTable "${pivotName}"`);
    } catch (e) {
        const errorMsg = `Failed to delete PivotTable: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// Slicer Operations
// ============================================================================

/**
 * Valid slicer styles for validation
 */
const VALID_SLICER_STYLES = [
    ...Array.from({ length: 6 }, (_, i) => `SlicerStyleLight${i + 1}`),
    ...Array.from({ length: 6 }, (_, i) => `SlicerStyleDark${i + 1}`)
];

/**
 * Creates a slicer for a Table or PivotTable
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with slicer options
 */
async function createSlicer(ctx, sheet, action) {
    logDiag(`Starting createSlicer for target "${action.target}"`);
    
    let options = {
        slicerName: null,
        sourceType: null,
        sourceName: action.target,
        field: null,
        position: { left: 100, top: 100, width: 200, height: 200 },
        style: "SlicerStyleLight1",
        selectedItems: null,  // Array of items to select
        multiSelect: true     // Whether multiple items can be selected
    };
    
    if (action.data) {
        try {
            const parsed = JSON.parse(action.data);
            options = { ...options, ...parsed };
            if (parsed.position) {
                options.position = { ...options.position, ...parsed.position };
            }
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for createSlicer`);
        }
    }
    
    // Validate required fields
    if (!options.sourceName) {
        const errorMsg = `Source name (table or pivot name) is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.field) {
        const errorMsg = `Field name is required for slicer.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.sourceType || !["table", "pivot"].includes(options.sourceType.toLowerCase())) {
        const errorMsg = `Source type must be "table" or "pivot".`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        let slicerSource = null;
        let targetWorksheet = sheet;
        const sourceType = options.sourceType.toLowerCase();
        
        if (sourceType === "table") {
            // Search for table in all worksheets (Comment 4: align with pivot search behavior)
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            let table = null;
            for (const ws of sheets.items) {
                const tbl = ws.tables.getItemOrNullObject(options.sourceName);
                tbl.load("isNullObject");
                await ctx.sync();
                
                if (!tbl.isNullObject) {
                    table = tbl;
                    targetWorksheet = ws;
                    break;
                }
            }
            
            if (!table) {
                const errorMsg = `Table "${options.sourceName}" not found in any worksheet.`;
                logDiag(`Error: ${errorMsg}`);
                throw new Error(errorMsg);
            }
            
            // Comment 2: Validate that the field exists in the table
            table.columns.load("items");
            await ctx.sync();
            
            const columnNames = table.columns.items.map(col => {
                col.load("name");
                return col;
            });
            await ctx.sync();
            
            const fieldExists = columnNames.some(col => col.name === options.field);
            if (!fieldExists) {
                const availableColumns = columnNames.map(col => col.name).join(", ");
                const errorMsg = `Field "${options.field}" not found in table "${options.sourceName}". Available columns: ${availableColumns}`;
                logDiag(`Error: ${errorMsg}`);
                throw new Error(errorMsg);
            }
            
            slicerSource = table;
            logDiag(`Found table "${options.sourceName}" for slicer with valid field "${options.field}"`);
        } else if (sourceType === "pivot") {
            // Search for PivotTable in all sheets
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            let pivotTable = null;
            for (const ws of sheets.items) {
                const pt = ws.pivotTables.getItemOrNullObject(options.sourceName);
                pt.load("isNullObject");
                await ctx.sync();
                
                if (!pt.isNullObject) {
                    pivotTable = pt;
                    ws.load("name");
                    await ctx.sync();
                    targetWorksheet = ws;
                    break;
                }
            }
            
            if (!pivotTable) {
                const errorMsg = `PivotTable "${options.sourceName}" not found.`;
                logDiag(`Error: ${errorMsg}`);
                throw new Error(errorMsg);
            }
            
            // Comment 2: Validate that the field exists in the pivot table hierarchies
            pivotTable.hierarchies.load("items");
            await ctx.sync();
            
            const hierarchyNames = pivotTable.hierarchies.items.map(h => {
                h.load("name");
                return h;
            });
            await ctx.sync();
            
            const fieldExists = hierarchyNames.some(h => h.name === options.field);
            if (!fieldExists) {
                const availableFields = hierarchyNames.map(h => h.name).join(", ");
                const errorMsg = `Field "${options.field}" not found in PivotTable "${options.sourceName}". Available fields: ${availableFields}`;
                logDiag(`Error: ${errorMsg}`);
                throw new Error(errorMsg);
            }
            
            slicerSource = pivotTable;
            logDiag(`Found PivotTable "${options.sourceName}" for slicer with valid field "${options.field}"`);
        }
        
        // Create the slicer
        const slicer = targetWorksheet.slicers.add(slicerSource, options.field, targetWorksheet);
        
        // Set slicer name if provided
        if (options.slicerName) {
            slicer.name = options.slicerName;
        }
        
        // Set position and size
        slicer.left = options.position.left || 100;
        slicer.top = options.position.top || 100;
        slicer.width = options.position.width || 200;
        slicer.height = options.position.height || 200;
        
        // Set style with validation
        if (options.style) {
            if (VALID_SLICER_STYLES.includes(options.style)) {
                slicer.style = options.style;
            } else {
                logDiag(`Warning: Invalid slicer style "${options.style}". Using default SlicerStyleLight1.`);
                slicer.style = "SlicerStyleLight1";
            }
        }
        
        await ctx.sync();
        
        // Comment 1: Configure slicer item selections if specified
        if (options.selectedItems && Array.isArray(options.selectedItems) && options.selectedItems.length > 0) {
            slicer.slicerItems.load("items");
            await ctx.sync();
            
            const slicerItems = slicer.slicerItems.items;
            for (const item of slicerItems) {
                item.load("name");
            }
            await ctx.sync();
            
            // If multiSelect is false, only select the first item from selectedItems
            const itemsToSelect = options.multiSelect === false 
                ? [options.selectedItems[0]] 
                : options.selectedItems;
            
            for (const item of slicerItems) {
                const shouldBeSelected = itemsToSelect.includes(item.name);
                item.isSelected = shouldBeSelected;
            }
            
            await ctx.sync();
            logDiag(`Configured slicer selections: ${itemsToSelect.join(", ")}`);
        }
        
        const slicerDisplayName = options.slicerName || options.field;
        logDiag(`Successfully created slicer "${slicerDisplayName}" for ${sourceType} "${options.sourceName}" on field "${options.field}"`);
    } catch (e) {
        const errorMsg = `Failed to create slicer: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Configures an existing slicer's properties
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with configuration options
 */
async function configureSlicer(ctx, sheet, action) {
    logDiag(`Starting configureSlicer for target "${action.target}"`);
    
    let options = {
        slicerName: action.target,
        caption: null,
        style: null,
        sortBy: null,
        width: null,
        height: null,
        left: null,
        top: null,
        selectedItems: null,  // Comment 1: Array of items to select
        multiSelect: true     // Comment 1: Whether multiple items can be selected
    };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for configureSlicer`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for slicer in all worksheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let slicer = null;
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load("isNullObject");
            await ctx.sync();
            
            if (!sl.isNullObject) {
                slicer = sl;
                break;
            }
        }
        
        if (!slicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        const updatedProps = [];
        
        // Apply properties conditionally
        if (options.caption !== null && options.caption !== undefined) {
            slicer.caption = options.caption;
            updatedProps.push("caption");
        }
        
        if (options.style) {
            if (VALID_SLICER_STYLES.includes(options.style)) {
                slicer.style = options.style;
                updatedProps.push("style");
            } else {
                logDiag(`Warning: Invalid slicer style "${options.style}". Skipping style update.`);
            }
        }
        
        if (options.sortBy) {
            const sortMap = {
                "datasourceorder": Excel.SlicerSortType.dataSourceOrder,
                "ascending": Excel.SlicerSortType.ascending,
                "descending": Excel.SlicerSortType.descending
            };
            const sortKey = options.sortBy.toLowerCase().replace(/\s/g, "");
            if (sortMap[sortKey]) {
                slicer.sortBy = sortMap[sortKey];
                updatedProps.push("sortBy");
            } else {
                logDiag(`Warning: Invalid sortBy value "${options.sortBy}". Use DataSourceOrder, Ascending, or Descending.`);
            }
        }
        
        if (options.width !== null && options.width !== undefined) {
            slicer.width = options.width;
            updatedProps.push("width");
        }
        
        if (options.height !== null && options.height !== undefined) {
            slicer.height = options.height;
            updatedProps.push("height");
        }
        
        if (options.left !== null && options.left !== undefined) {
            slicer.left = options.left;
            updatedProps.push("left");
        }
        
        if (options.top !== null && options.top !== undefined) {
            slicer.top = options.top;
            updatedProps.push("top");
        }
        
        await ctx.sync();
        
        // Comment 1: Configure slicer item selections if specified
        if (options.selectedItems && Array.isArray(options.selectedItems) && options.selectedItems.length > 0) {
            slicer.slicerItems.load("items");
            await ctx.sync();
            
            const slicerItems = slicer.slicerItems.items;
            for (const item of slicerItems) {
                item.load("name");
            }
            await ctx.sync();
            
            // If multiSelect is false, only select the first item from selectedItems
            const itemsToSelect = options.multiSelect === false 
                ? [options.selectedItems[0]] 
                : options.selectedItems;
            
            for (const item of slicerItems) {
                const shouldBeSelected = itemsToSelect.includes(item.name);
                item.isSelected = shouldBeSelected;
            }
            
            await ctx.sync();
            updatedProps.push(`selectedItems(${itemsToSelect.length})`);
            logDiag(`Configured slicer selections: ${itemsToSelect.join(", ")}`);
        }
        
        logDiag(`Successfully configured slicer "${slicerName}". Updated: ${updatedProps.join(", ") || "none"}`);
    } catch (e) {
        const errorMsg = `Failed to configure slicer: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Connects a slicer to a different table (recreates slicer)
 * Note: Office.js doesn't support rebinding slicers; this deletes and recreates
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with connection options
 */
async function connectSlicerToTable(ctx, sheet, action) {
    logDiag(`Starting connectSlicerToTable for target "${action.target}"`);
    
    let options = {
        slicerName: action.target,
        tableName: null,
        field: null
    };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for connectSlicerToTable`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.tableName) {
        const errorMsg = `Table name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.field) {
        const errorMsg = `Field name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Find existing slicer to get its properties
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let existingSlicer = null;
        let slicerWorksheet = null;
        
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load(["isNullObject", "left", "top", "width", "height", "style", "caption"]);
            await ctx.sync();
            
            if (!sl.isNullObject) {
                existingSlicer = sl;
                slicerWorksheet = ws;
                break;
            }
        }
        
        if (!existingSlicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Store slicer properties before deletion
        const slicerProps = {
            left: existingSlicer.left,
            top: existingSlicer.top,
            width: existingSlicer.width,
            height: existingSlicer.height,
            style: existingSlicer.style,
            caption: existingSlicer.caption
        };
        
        // Delete existing slicer
        existingSlicer.delete();
        await ctx.sync();
        logDiag(`Deleted existing slicer "${slicerName}" for reconnection`);
        
        // Comment 3: Search for table in all worksheets instead of just active sheet
        let table = null;
        let tableWorksheet = slicerWorksheet; // Default to original slicer's worksheet
        
        for (const ws of sheets.items) {
            const tbl = ws.tables.getItemOrNullObject(options.tableName);
            tbl.load("isNullObject");
            await ctx.sync();
            
            if (!tbl.isNullObject) {
                table = tbl;
                tableWorksheet = ws;
                break;
            }
        }
        
        if (!table) {
            const errorMsg = `Table "${options.tableName}" not found in any worksheet.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Create new slicer on the table's worksheet (Comment 3: use correct worksheet)
        const newSlicer = tableWorksheet.slicers.add(table, options.field, tableWorksheet);
        newSlicer.name = slicerName;
        newSlicer.left = slicerProps.left;
        newSlicer.top = slicerProps.top;
        newSlicer.width = slicerProps.width;
        newSlicer.height = slicerProps.height;
        if (slicerProps.style) newSlicer.style = slicerProps.style;
        if (slicerProps.caption) newSlicer.caption = slicerProps.caption;
        
        await ctx.sync();
        logDiag(`Successfully reconnected slicer "${slicerName}" to table "${options.tableName}" on field "${options.field}"`);
    } catch (e) {
        const errorMsg = `Failed to connect slicer to table: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Connects a slicer to a different PivotTable (recreates slicer)
 * Note: Office.js doesn't support rebinding slicers; this deletes and recreates
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with connection options
 */
async function connectSlicerToPivot(ctx, sheet, action) {
    logDiag(`Starting connectSlicerToPivot for target "${action.target}"`);
    
    let options = {
        slicerName: action.target,
        pivotName: null,
        field: null
    };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for connectSlicerToPivot`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.pivotName) {
        const errorMsg = `PivotTable name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (!options.field) {
        const errorMsg = `Field name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Find existing slicer to get its properties
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let existingSlicer = null;
        
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load(["isNullObject", "left", "top", "width", "height", "style", "caption"]);
            await ctx.sync();
            
            if (!sl.isNullObject) {
                existingSlicer = sl;
                break;
            }
        }
        
        if (!existingSlicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Store slicer properties before deletion
        const slicerProps = {
            left: existingSlicer.left,
            top: existingSlicer.top,
            width: existingSlicer.width,
            height: existingSlicer.height,
            style: existingSlicer.style,
            caption: existingSlicer.caption
        };
        
        // Delete existing slicer
        existingSlicer.delete();
        await ctx.sync();
        logDiag(`Deleted existing slicer "${slicerName}" for reconnection`);
        
        // Find the PivotTable
        let pivotTable = null;
        let pivotWorksheet = null;
        
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(options.pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                pivotWorksheet = ws;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${options.pivotName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Create new slicer with same properties
        const newSlicer = pivotWorksheet.slicers.add(pivotTable, options.field, pivotWorksheet);
        newSlicer.name = slicerName;
        newSlicer.left = slicerProps.left;
        newSlicer.top = slicerProps.top;
        newSlicer.width = slicerProps.width;
        newSlicer.height = slicerProps.height;
        if (slicerProps.style) newSlicer.style = slicerProps.style;
        if (slicerProps.caption) newSlicer.caption = slicerProps.caption;
        
        await ctx.sync();
        logDiag(`Successfully reconnected slicer "${slicerName}" to PivotTable "${options.pivotName}" on field "${options.field}"`);
    } catch (e) {
        const errorMsg = `Failed to connect slicer to PivotTable: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes a slicer
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with slicer name
 */
async function deleteSlicer(ctx, sheet, action) {
    logDiag(`Starting deleteSlicer for target "${action.target}"`);
    
    let options = { slicerName: action.target };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for deleteSlicer`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for slicer in all worksheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let slicer = null;
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load("isNullObject");
            await ctx.sync();
            
            if (!sl.isNullObject) {
                slicer = sl;
                break;
            }
        }
        
        if (!slicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        slicer.delete();
        await ctx.sync();
        logDiag(`Successfully deleted slicer "${slicerName}"`);
    } catch (e) {
        const errorMsg = `Failed to delete slicer: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// Exports
// ============================================================================

export {
    setDiagnosticLogger,
    executeAction,
    applyFormula,
    adjustFormulaReferences,
    applyValues,
    applyFormat,
    applyConditionalFormat,
    clearConditionalFormat,
    applyValidation,
    createChart,
    createPivotChart,
    applySort,
    applyFilter,
    clearFilter,
    applyCopy,
    applyCopyValues,
    createSheet,
    removeDuplicates,
    createTable,
    styleTable,
    addTableRow,
    addTableColumn,
    resizeTable,
    convertToRange,
    toggleTableTotals,
    insertRows,
    insertColumns,
    deleteRows,
    deleteColumns,
    mergeCells,
    unmergeCells,
    findReplace,
    textToColumns,
    createPivotTable,
    addPivotField,
    configurePivotLayout,
    refreshPivotTable,
    deletePivotTable,
    createSlicer,
    configureSlicer,
    connectSlicerToTable,
    connectSlicerToPivot,
    deleteSlicer
};
