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
    
    const range = sheet.getRange(target);
    range.load(["rowCount", "columnCount"]);
    await ctx.sync();
    
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
            
        default:
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
    removeDuplicates
};
