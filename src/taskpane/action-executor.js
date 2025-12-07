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
        "deleteSlicer",          // target is slicer name
        "deleteNamedRange",      // target is named range name
        "updateNamedRange",      // target is named range name
        "listNamedRanges",       // target is scope option
        "formatShape",           // target is shape name
        "deleteShape",           // target is shape name
        "groupShapes",           // target is shape names (comma-separated)
        "arrangeShapes",         // target is shape name
        "ungroupShapes",         // target is group name
        "addComment",            // target is cell address (comment API handles it)
        "addNote",               // target is cell address (note API handles it)
        "editComment",           // target is cell with comment
        "editNote",              // target is cell with note
        "deleteComment",         // target is cell with comment
        "deleteNote",            // target is cell with note
        "replyToComment",        // target is cell with parent comment
        "resolveComment",        // target is cell with comment
        "createSparkline",       // target is location cell/range
        "configureSparkline",    // target is sparkline location
        "deleteSparkline",       // target is sparkline location
        "renameSheet",           // target is sheet name
        "moveSheet",             // target is sheet name
        "hideSheet",             // target is sheet name
        "unhideSheet",           // target is sheet name
        "unfreezePane",          // target is "current" or sheet name
        "setZoom",               // target is "current" or sheet name
        "createView",            // target is view name
        "setPageSetup",          // target is sheet name
        "setPageMargins",        // target is sheet name
        "setPageOrientation",    // target is sheet name
        "setHeaderFooter",       // target is sheet name
        "setPageBreaks"          // target is sheet name
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
            
        case "createNamedRange":
            await createNamedRange(ctx, sheet, action);
            break;
            
        case "deleteNamedRange":
            await deleteNamedRange(ctx, sheet, action);
            break;
            
        case "updateNamedRange":
            await updateNamedRange(ctx, sheet, action);
            break;
            
        case "listNamedRanges":
            await listNamedRanges(ctx, sheet, action);
            break;
            
        case "protectWorksheet":
            await protectWorksheet(ctx, sheet, action);
            break;
            
        case "unprotectWorksheet":
            await unprotectWorksheet(ctx, sheet, action);
            break;
            
        case "protectRange":
            await protectRange(ctx, sheet, action);
            break;
            
        case "unprotectRange":
            await unprotectRange(ctx, sheet, action);
            break;
            
        case "protectWorkbook":
            await protectWorkbook(ctx, sheet, action);
            break;
            
        case "unprotectWorkbook":
            await unprotectWorkbook(ctx, sheet, action);
            break;
            
        case "insertShape":
            await insertShape(ctx, sheet, action);
            break;
            
        case "insertImage":
            await insertImage(ctx, sheet, action);
            break;
            
        case "insertTextBox":
            await insertTextBox(ctx, sheet, action);
            break;
            
        case "formatShape":
            await formatShape(ctx, sheet, target, data);
            break;
            
        case "deleteShape":
            await deleteShape(ctx, sheet, target);
            break;
            
        case "groupShapes":
            await groupShapes(ctx, sheet, action);
            break;
            
        case "arrangeShapes":
            await arrangeShapes(ctx, sheet, target, data);
            break;
            
        case "ungroupShapes":
            await ungroupShapes(ctx, sheet, target);
            break;
            
        case "addComment":
            await addComment(ctx, sheet, action);
            break;
            
        case "addNote":
            await addNote(ctx, sheet, action);
            break;
            
        case "editComment":
            await editComment(ctx, sheet, action);
            break;
            
        case "editNote":
            await editNote(ctx, sheet, action);
            break;
            
        case "deleteComment":
            await deleteComment(ctx, sheet, action);
            break;
            
        case "deleteNote":
            await deleteNote(ctx, sheet, action);
            break;
            
        case "replyToComment":
            await replyToComment(ctx, sheet, action);
            break;
            
        case "resolveComment":
            await resolveComment(ctx, sheet, action);
            break;
            
        case "createSparkline":
            await createSparkline(ctx, sheet, action);
            break;
            
        case "configureSparkline":
            await configureSparkline(ctx, sheet, action);
            break;
            
        case "deleteSparkline":
            await deleteSparkline(ctx, sheet, action);
            break;
            
        case "renameSheet":
            await renameSheet(ctx, sheet, action);
            break;
            
        case "moveSheet":
            await moveSheet(ctx, sheet, action);
            break;
            
        case "hideSheet":
            await hideSheet(ctx, sheet, action);
            break;
            
        case "unhideSheet":
            await unhideSheet(ctx, sheet, action);
            break;
            
        case "freezePanes":
            await freezePanes(ctx, sheet, action);
            break;
            
        case "unfreezePane":
            await unfreezePane(ctx, sheet, action);
            break;
            
        case "setZoom":
            await setZoom(ctx, sheet, action);
            break;
            
        case "splitPane":
            await splitPane(ctx, sheet, action);
            break;
            
        case "createView":
            await createView(ctx, sheet, action);
            break;
            
        case "setPageSetup":
            await setPageSetup(ctx, sheet, action);
            break;
            
        case "setPageMargins":
            await setPageMargins(ctx, sheet, action);
            break;
            
        case "setPageOrientation":
            await setPageOrientation(ctx, sheet, action);
            break;
            
        case "setPrintArea":
            await setPrintArea(ctx, sheet, action);
            break;
            
        case "setHeaderFooter":
            await setHeaderFooter(ctx, sheet, action);
            break;
            
        case "setPageBreaks":
            await setPageBreaks(ctx, sheet, action);
            break;
            
        case "insertDataType":
            await insertDataType(ctx, sheet, range, action);
            break;
            
        case "refreshDataType":
            await refreshDataType(ctx, sheet, range, action);
            break;
            
        case "addHyperlink":
            await addHyperlink(ctx, range, data);
            break;
            
        case "removeHyperlink":
            await removeHyperlink(ctx, range);
            break;
            
        case "editHyperlink":
            await editHyperlink(ctx, range, data);
            break;
            
        default:
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
 * Applies formatting to a range with comprehensive Office.js support
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string of format options
 * 
 * Supported format options:
 * - Font: bold, italic, fontColor, fontSize
 * - Fill: fill (hex color)
 * - Alignment: horizontalAlignment, verticalAlignment
 * - Text Control: wrapText, textOrientation, indentLevel, shrinkToFit, readingOrder
 * - Number Format: numberFormat (custom code), numberFormatPreset (shortcut)
 * - Cell Style: style (predefined style name)
 * - Borders: border (boolean for all edges), borders (object for individual sides)
 */
async function applyFormat(ctx, range, data) {
    let fmt;
    try { fmt = JSON.parse(data); } catch { fmt = {}; }
    
    const appliedProps = [];
    
    // ========== Font Properties ==========
    if (fmt.bold !== undefined) {
        range.format.font.bold = fmt.bold;
        appliedProps.push("bold");
    }
    if (fmt.italic !== undefined) {
        range.format.font.italic = fmt.italic;
        appliedProps.push("italic");
    }
    if (fmt.fontColor) {
        range.format.font.color = fmt.fontColor;
        appliedProps.push("fontColor");
    }
    if (fmt.fontSize) {
        range.format.font.size = fmt.fontSize;
        appliedProps.push("fontSize");
    }
    
    // ========== Fill Properties ==========
    if (fmt.fill) {
        range.format.fill.color = fmt.fill;
        appliedProps.push("fill");
    }
    
    // ========== Alignment Properties ==========
    const validHorizontalAlignments = ["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"];
    const validVerticalAlignments = ["Top", "Center", "Bottom", "Justify", "Distributed"];
    
    if (fmt.horizontalAlignment) {
        if (validHorizontalAlignments.includes(fmt.horizontalAlignment)) {
            range.format.horizontalAlignment = fmt.horizontalAlignment;
            appliedProps.push("horizontalAlignment");
        } else {
            logDiag(`Invalid horizontalAlignment: ${fmt.horizontalAlignment}`);
        }
    }
    if (fmt.verticalAlignment) {
        if (validVerticalAlignments.includes(fmt.verticalAlignment)) {
            range.format.verticalAlignment = fmt.verticalAlignment;
            appliedProps.push("verticalAlignment");
        } else {
            logDiag(`Invalid verticalAlignment: ${fmt.verticalAlignment}`);
        }
    }
    
    // ========== Text Control Properties ==========
    if (fmt.wrapText !== undefined) {
        range.format.wrapText = fmt.wrapText;
        appliedProps.push("wrapText");
    }
    if (fmt.textOrientation !== undefined) {
        // Valid range: -90 to 90, or 255 for vertical stacked text
        const orientation = parseInt(fmt.textOrientation);
        if ((orientation >= -90 && orientation <= 90) || orientation === 255) {
            range.format.textOrientation = orientation;
            appliedProps.push("textOrientation");
        } else {
            logDiag(`Invalid textOrientation: ${fmt.textOrientation} (must be -90 to 90, or 255)`);
        }
    }
    if (fmt.indentLevel !== undefined) {
        const indent = parseInt(fmt.indentLevel);
        if (indent >= 0 && indent <= 250) {
            range.format.indentLevel = indent;
            appliedProps.push("indentLevel");
        } else {
            logDiag(`Invalid indentLevel: ${fmt.indentLevel} (must be 0-250)`);
        }
    }
    if (fmt.shrinkToFit !== undefined) {
        range.format.shrinkToFit = fmt.shrinkToFit;
        appliedProps.push("shrinkToFit");
    }
    if (fmt.readingOrder) {
        const validReadingOrders = ["Context", "LeftToRight", "RightToLeft"];
        if (validReadingOrders.includes(fmt.readingOrder)) {
            range.format.readingOrder = fmt.readingOrder;
            appliedProps.push("readingOrder");
        } else {
            logDiag(`Invalid readingOrder: ${fmt.readingOrder}`);
        }
    }
    
    // ========== Number Format ==========
    // Number format presets mapping
    const numberFormatPresets = {
        "currency": "$#,##0.00",
        "accounting": "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        "percentage": "0.00%",
        "date": "m/d/yyyy",
        "dateShort": "mm/dd/yy",
        "dateLong": "dddd, mmmm dd, yyyy",
        "time": "h:mm:ss AM/PM",
        "timeShort": "h:mm AM/PM",
        "time24": "hh:mm:ss",
        "fraction": "# ?/?",
        "scientific": "0.00E+00",
        "text": "@",
        "number": "#,##0.00",
        "integer": "#,##0"
    };
    
    if (fmt.numberFormatPreset && numberFormatPresets[fmt.numberFormatPreset]) {
        range.numberFormat = [[numberFormatPresets[fmt.numberFormatPreset]]];
        appliedProps.push(`numberFormatPreset:${fmt.numberFormatPreset}`);
    } else if (fmt.numberFormat) {
        range.numberFormat = [[fmt.numberFormat]];
        appliedProps.push("numberFormat");
    }
    
    // ========== Cell Style ==========
    // Predefined Excel cell styles
    const validStyles = [
        "Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Title", "Total",
        "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
        "Good", "Bad", "Neutral", "Warning Text",
        "Input", "Output", "Calculation", "Check Cell", "Explanatory Text", "Linked Cell", "Note"
    ];
    
    if (fmt.style) {
        if (validStyles.includes(fmt.style)) {
            try {
                range.format.style = fmt.style;
                appliedProps.push(`style:${fmt.style}`);
            } catch (styleError) {
                logDiag(`Failed to apply style "${fmt.style}": ${styleError.message}`);
            }
        } else {
            logDiag(`Invalid style: ${fmt.style}. Valid styles: ${validStyles.join(", ")}`);
        }
    }
    
    // ========== Border Properties ==========
    const validBorderStyles = ["Continuous", "Dash", "DashDot", "DashDotDot", "Dot", "Double", "None"];
    const validBorderWeights = ["Hairline", "Thin", "Medium", "Thick"];
    const borderSides = {
        "top": "EdgeTop",
        "bottom": "EdgeBottom",
        "left": "EdgeLeft",
        "right": "EdgeRight",
        "insideHorizontal": "InsideHorizontal",
        "insideVertical": "InsideVertical",
        "diagonalDown": "DiagonalDown",
        "diagonalUp": "DiagonalUp"
    };
    
    // Simple border (backward compatible) - applies continuous black thin borders to all edges
    if (fmt.border === true) {
        const edgeTop = range.format.borders.getItem("EdgeTop");
        edgeTop.style = "Continuous";
        edgeTop.color = "#000000";
        edgeTop.weight = "Thin";
        
        const edgeBottom = range.format.borders.getItem("EdgeBottom");
        edgeBottom.style = "Continuous";
        edgeBottom.color = "#000000";
        edgeBottom.weight = "Thin";
        
        const edgeLeft = range.format.borders.getItem("EdgeLeft");
        edgeLeft.style = "Continuous";
        edgeLeft.color = "#000000";
        edgeLeft.weight = "Thin";
        
        const edgeRight = range.format.borders.getItem("EdgeRight");
        edgeRight.style = "Continuous";
        edgeRight.color = "#000000";
        edgeRight.weight = "Thin";
        
        appliedProps.push("border:all");
    }
    
    // Advanced borders (individual sides with style/color/weight)
    if (fmt.borders && typeof fmt.borders === "object") {
        for (const [side, borderConfig] of Object.entries(fmt.borders)) {
            const excelSide = borderSides[side];
            if (!excelSide) {
                logDiag(`Invalid border side: ${side}`);
                continue;
            }
            
            try {
                const border = range.format.borders.getItem(excelSide);
                
                // Apply border style
                if (borderConfig.style) {
                    if (validBorderStyles.includes(borderConfig.style)) {
                        border.style = borderConfig.style;
                    } else {
                        logDiag(`Invalid border style: ${borderConfig.style}`);
                    }
                } else {
                    // Default to Continuous if not specified
                    border.style = "Continuous";
                }
                
                // Apply border color
                if (borderConfig.color) {
                    border.color = borderConfig.color;
                }
                
                // Apply border weight
                if (borderConfig.weight) {
                    if (validBorderWeights.includes(borderConfig.weight)) {
                        border.weight = borderConfig.weight;
                    } else {
                        logDiag(`Invalid border weight: ${borderConfig.weight}`);
                    }
                }
                
                appliedProps.push(`border:${side}`);
            } catch (borderError) {
                logDiag(`Failed to apply border for ${side}: ${borderError.message}`);
            }
        }
    }
    
    logDiag(`Applied formatting: ${appliedProps.join(", ")}`);
}

/**
 * Applies conditional formatting to a range with comprehensive Office.js support
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string of conditional format rules (single rule or array)
 * 
 * Supported conditional format types:
 * - cellValue: Comparison operators (GreaterThan, LessThan, EqualTo, Between, etc.)
 * - colorScale: 2-color or 3-color gradient based on values (min, mid, max)
 * - dataBar: In-cell bar charts with direction and color customization
 * - iconSet: 3/4/5 icon indicators (arrows, traffic lights, ratings, etc.)
 * - topBottom: Top/bottom N items or percent
 * - preset: Duplicates, unique values, above/below average, date-based rules
 * - textComparison: Contains, begins with, ends with text matching
 * - custom: Formula-based conditional formatting for complex logic
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
    
    // Validation helper for hex colors
    const isValidHexColor = (color) => /^#[0-9A-Fa-f]{6}$/.test(color);
    
    // Valid icon set styles
    const validIconSets = [
        "threeArrows", "threeArrowsGray", "threeTriangles", "threeFlags",
        "threeTrafficLights1", "threeTrafficLights2", "threeSigns",
        "threeSymbols", "threeSymbols2", "threeStars",
        "fourArrows", "fourArrowsGray", "fourRedToBlack", "fourRating", "fourTrafficLights",
        "fiveArrows", "fiveArrowsGray", "fiveRating", "fiveQuarters", "fiveBoxes"
    ];
    
    // Valid preset criteria
    const validPresetCriteria = [
        "duplicateValues", "uniqueValues", "aboveAverage", "belowAverage",
        "equalOrAboveAverage", "equalOrBelowAverage",
        "oneStdDevAboveAverage", "oneStdDevBelowAverage",
        "twoStdDevAboveAverage", "twoStdDevBelowAverage",
        "threeStdDevAboveAverage", "threeStdDevBelowAverage",
        "yesterday", "today", "tomorrow", "lastSevenDays",
        "lastWeek", "thisWeek", "nextWeek",
        "lastMonth", "thisMonth", "nextMonth"
    ];
    
    range.conditionalFormats.clearAll();
    await ctx.sync();
    
    let appliedCount = 0;
    
    for (const rule of rules) {
        try {
            const ruleType = rule.type || "cellValue";
            
            // ========== Cell Value (existing) ==========
            if (ruleType === "cellValue" && rule.operator && rule.value !== undefined) {
                // Map operator string to Excel enum
                const operatorMap = {
                    "GreaterThan": Excel.ConditionalCellValueOperator.greaterThan,
                    "LessThan": Excel.ConditionalCellValueOperator.lessThan,
                    "EqualTo": Excel.ConditionalCellValueOperator.equalTo,
                    "NotEqualTo": Excel.ConditionalCellValueOperator.notEqual,
                    "GreaterThanOrEqual": Excel.ConditionalCellValueOperator.greaterThanOrEqual,
                    "LessThanOrEqual": Excel.ConditionalCellValueOperator.lessThanOrEqual,
                    "Between": Excel.ConditionalCellValueOperator.between
                };
                
                const operator = operatorMap[rule.operator];
                if (!operator) {
                    logDiag(`Invalid cellValue operator: ${rule.operator}. Valid operators: ${Object.keys(operatorMap).join(", ")}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                
                // Apply fill color with validation
                const fillColor = rule.fill || "#FFFF00";
                if (isValidHexColor(fillColor)) {
                    cf.cellValue.format.fill.color = fillColor;
                } else {
                    logDiag(`Invalid fill color: ${fillColor}, using default #FFFF00`);
                    cf.cellValue.format.fill.color = "#FFFF00";
                }
                
                // Apply font color with validation
                if (rule.fontColor) {
                    if (isValidHexColor(rule.fontColor)) {
                        cf.cellValue.format.font.color = rule.fontColor;
                    } else {
                        logDiag(`Invalid fontColor: ${rule.fontColor}, skipping`);
                    }
                }
                if (rule.bold) cf.cellValue.format.font.bold = rule.bold;
                
                cf.cellValue.rule = {
                    formula1: String(rule.value),
                    formula2: rule.value2 ? String(rule.value2) : undefined,
                    operator: operator
                };
                appliedCount++;
                logDiag(`Applied cellValue: ${rule.operator} ${rule.value}`);
            }
            
            // ========== Color Scale ==========
            else if (ruleType === "colorScale") {
                if (!rule.minimum || !rule.maximum) {
                    logDiag(`colorScale requires minimum and maximum properties`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                
                // Helper to map criterion type
                const mapCriterionType = (type) => {
                    const typeMap = {
                        "lowestValue": Excel.ConditionalFormatColorCriterionType.lowestValue,
                        "highestValue": Excel.ConditionalFormatColorCriterionType.highestValue,
                        "number": Excel.ConditionalFormatColorCriterionType.number,
                        "percent": Excel.ConditionalFormatColorCriterionType.percent,
                        "percentile": Excel.ConditionalFormatColorCriterionType.percentile,
                        "formula": Excel.ConditionalFormatColorCriterionType.formula
                    };
                    return typeMap[type] || Excel.ConditionalFormatColorCriterionType.lowestValue;
                };
                
                // Validate and get colors with defaults
                const minColor = rule.minimum.color && isValidHexColor(rule.minimum.color) ? rule.minimum.color : "#63BE7B";
                const maxColor = rule.maximum.color && isValidHexColor(rule.maximum.color) ? rule.maximum.color : "#F8696B";
                
                if (rule.minimum.color && !isValidHexColor(rule.minimum.color)) {
                    logDiag(`Invalid minimum color: ${rule.minimum.color}, using default #63BE7B`);
                }
                if (rule.maximum.color && !isValidHexColor(rule.maximum.color)) {
                    logDiag(`Invalid maximum color: ${rule.maximum.color}, using default #F8696B`);
                }
                
                const criteria = {
                    minimum: {
                        type: mapCriterionType(rule.minimum.type),
                        color: minColor,
                        formula: rule.minimum.formula || null
                    },
                    maximum: {
                        type: mapCriterionType(rule.maximum.type),
                        color: maxColor,
                        formula: rule.maximum.formula || null
                    }
                };
                
                // Add midpoint for 3-color scale
                if (rule.midpoint) {
                    const midColor = rule.midpoint.color && isValidHexColor(rule.midpoint.color) ? rule.midpoint.color : "#FFEB84";
                    if (rule.midpoint.color && !isValidHexColor(rule.midpoint.color)) {
                        logDiag(`Invalid midpoint color: ${rule.midpoint.color}, using default #FFEB84`);
                    }
                    criteria.midpoint = {
                        type: mapCriterionType(rule.midpoint.type),
                        color: midColor,
                        formula: rule.midpoint.formula || null
                    };
                }
                
                cf.colorScale.criteria = criteria;
                appliedCount++;
                const scaleType = rule.midpoint ? "3-color" : "2-color";
                logDiag(`Applied colorScale (${scaleType}: ${criteria.minimum.color} -> ${rule.midpoint ? criteria.midpoint.color + " -> " : ""}${criteria.maximum.color})`);
            }
            
            // ========== Data Bar ==========
            else if (ruleType === "dataBar") {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                
                // Set bar direction
                if (rule.barDirection) {
                    const directionMap = {
                        "Context": Excel.ConditionalDataBarDirection.context,
                        "LeftToRight": Excel.ConditionalDataBarDirection.leftToRight,
                        "RightToLeft": Excel.ConditionalDataBarDirection.rightToLeft
                    };
                    cf.dataBar.barDirection = directionMap[rule.barDirection] || Excel.ConditionalDataBarDirection.context;
                }
                
                // Show data bar only (hide values)
                if (rule.showDataBarOnly !== undefined) {
                    cf.dataBar.showDataBarOnly = rule.showDataBarOnly;
                }
                
                // Positive format with color validation
                if (rule.positiveFormat) {
                    if (rule.positiveFormat.fillColor) {
                        if (isValidHexColor(rule.positiveFormat.fillColor)) {
                            cf.dataBar.positiveFormat.fillColor = rule.positiveFormat.fillColor;
                        } else {
                            logDiag(`Invalid positiveFormat.fillColor: ${rule.positiveFormat.fillColor}, skipping`);
                        }
                    }
                    if (rule.positiveFormat.borderColor) {
                        if (isValidHexColor(rule.positiveFormat.borderColor)) {
                            cf.dataBar.positiveFormat.borderColor = rule.positiveFormat.borderColor;
                        } else {
                            logDiag(`Invalid positiveFormat.borderColor: ${rule.positiveFormat.borderColor}, skipping`);
                        }
                    }
                    if (rule.positiveFormat.gradientFill !== undefined) {
                        cf.dataBar.positiveFormat.gradientFill = rule.positiveFormat.gradientFill;
                    }
                }
                
                // Negative format with color validation
                if (rule.negativeFormat) {
                    if (rule.negativeFormat.fillColor) {
                        if (isValidHexColor(rule.negativeFormat.fillColor)) {
                            cf.dataBar.negativeFormat.fillColor = rule.negativeFormat.fillColor;
                        } else {
                            logDiag(`Invalid negativeFormat.fillColor: ${rule.negativeFormat.fillColor}, skipping`);
                        }
                    }
                    if (rule.negativeFormat.borderColor) {
                        if (isValidHexColor(rule.negativeFormat.borderColor)) {
                            cf.dataBar.negativeFormat.borderColor = rule.negativeFormat.borderColor;
                        } else {
                            logDiag(`Invalid negativeFormat.borderColor: ${rule.negativeFormat.borderColor}, skipping`);
                        }
                    }
                }
                
                // Axis color with validation
                if (rule.axisColor) {
                    if (isValidHexColor(rule.axisColor)) {
                        cf.dataBar.axisColor = rule.axisColor;
                    } else {
                        logDiag(`Invalid axisColor: ${rule.axisColor}, skipping`);
                    }
                }
                
                appliedCount++;
                logDiag(`Applied dataBar (direction: ${rule.barDirection || "Context"}, showDataBarOnly: ${rule.showDataBarOnly || false})`);
            }
            
            // ========== Icon Set ==========
            else if (ruleType === "iconSet") {
                if (!rule.style) {
                    logDiag(`iconSet requires style property`);
                    continue;
                }
                
                if (!validIconSets.includes(rule.style)) {
                    logDiag(`Invalid iconSet style: ${rule.style}. Valid styles: ${validIconSets.join(", ")}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
                
                // Map style to Excel enum
                const styleMap = {
                    "threeArrows": Excel.IconSet.threeArrows,
                    "threeArrowsGray": Excel.IconSet.threeArrowsGray,
                    "threeTriangles": Excel.IconSet.threeTriangles,
                    "threeFlags": Excel.IconSet.threeFlags,
                    "threeTrafficLights1": Excel.IconSet.threeTrafficLights1,
                    "threeTrafficLights2": Excel.IconSet.threeTrafficLights2,
                    "threeSigns": Excel.IconSet.threeSigns,
                    "threeSymbols": Excel.IconSet.threeSymbols,
                    "threeSymbols2": Excel.IconSet.threeSymbols2,
                    "threeStars": Excel.IconSet.threeStars,
                    "fourArrows": Excel.IconSet.fourArrows,
                    "fourArrowsGray": Excel.IconSet.fourArrowsGray,
                    "fourRedToBlack": Excel.IconSet.fourRedToBlack,
                    "fourRating": Excel.IconSet.fourRating,
                    "fourTrafficLights": Excel.IconSet.fourTrafficLights,
                    "fiveArrows": Excel.IconSet.fiveArrows,
                    "fiveArrowsGray": Excel.IconSet.fiveArrowsGray,
                    "fiveRating": Excel.IconSet.fiveRating,
                    "fiveQuarters": Excel.IconSet.fiveQuarters,
                    "fiveBoxes": Excel.IconSet.fiveBoxes
                };
                
                cf.iconSet.style = styleMap[rule.style];
                
                // Determine expected criteria count based on icon set style
                const threeIconSets = ["threeArrows", "threeArrowsGray", "threeTriangles", "threeFlags", "threeTrafficLights1", "threeTrafficLights2", "threeSigns", "threeSymbols", "threeSymbols2", "threeStars"];
                const fourIconSets = ["fourArrows", "fourArrowsGray", "fourRedToBlack", "fourRating", "fourTrafficLights"];
                const fiveIconSets = ["fiveArrows", "fiveArrowsGray", "fiveRating", "fiveQuarters", "fiveBoxes"];
                
                let expectedCriteriaCount = 3;
                if (fourIconSets.includes(rule.style)) expectedCriteriaCount = 4;
                else if (fiveIconSets.includes(rule.style)) expectedCriteriaCount = 5;
                
                // Set criteria if provided
                if (rule.criteria && Array.isArray(rule.criteria)) {
                    // Validate criteria count matches icon count
                    if (rule.criteria.length !== expectedCriteriaCount) {
                        logDiag(`iconSet criteria count mismatch: expected ${expectedCriteriaCount} for ${rule.style}, got ${rule.criteria.length}. Skipping rule.`);
                        continue;
                    }
                    
                    const criteriaArray = rule.criteria.map(c => {
                        if (!c || Object.keys(c).length === 0) return {};
                        
                        const criterionTypeMap = {
                            "number": Excel.ConditionalFormatIconRuleType.number,
                            "percent": Excel.ConditionalFormatIconRuleType.percent,
                            "percentile": Excel.ConditionalFormatIconRuleType.percentile,
                            "formula": Excel.ConditionalFormatIconRuleType.formula
                        };
                        
                        const operatorMap = {
                            "greaterThan": Excel.ConditionalIconCriterionOperator.greaterThan,
                            "greaterThanOrEqual": Excel.ConditionalIconCriterionOperator.greaterThanOrEqual
                        };
                        
                        return {
                            type: criterionTypeMap[c.type] || Excel.ConditionalFormatIconRuleType.percent,
                            operator: operatorMap[c.operator] || Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
                            formula: c.formula || "0"
                        };
                    });
                    cf.iconSet.criteria = criteriaArray;
                }
                
                // Show icon only (hide values)
                if (rule.showIconOnly !== undefined) {
                    cf.iconSet.showIconOnly = rule.showIconOnly;
                }
                
                // Reverse icon order
                if (rule.reverseIconOrder !== undefined) {
                    cf.iconSet.reverseIconOrder = rule.reverseIconOrder;
                }
                
                appliedCount++;
                logDiag(`Applied iconSet (${rule.style}, ${rule.criteria ? rule.criteria.length : 0} thresholds)`);
            }
            
            // ========== Top/Bottom ==========
            else if (ruleType === "topBottom") {
                if (!rule.rule || rule.rank === undefined) {
                    logDiag(`topBottom requires rule and rank properties`);
                    continue;
                }
                
                // Validate rank is a positive integer
                const rank = parseInt(rule.rank);
                if (!Number.isInteger(rank) || rank <= 0) {
                    logDiag(`Invalid topBottom rank: ${rule.rank}. Rank must be a positive integer. Skipping rule.`);
                    continue;
                }
                
                // Map rule type and validate
                const ruleTypeMap = {
                    "TopItems": Excel.ConditionalTopBottomCriterionType.topItems,
                    "BottomItems": Excel.ConditionalTopBottomCriterionType.bottomItems,
                    "TopPercent": Excel.ConditionalTopBottomCriterionType.topPercent,
                    "BottomPercent": Excel.ConditionalTopBottomCriterionType.bottomPercent
                };
                
                const mappedRuleType = ruleTypeMap[rule.rule];
                if (!mappedRuleType) {
                    logDiag(`Invalid topBottom rule type: ${rule.rule}. Valid types: ${Object.keys(ruleTypeMap).join(", ")}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
                
                cf.topBottom.rule = {
                    type: mappedRuleType,
                    rank: rank
                };
                
                // Apply formatting with color validation
                if (rule.fill) {
                    if (isValidHexColor(rule.fill)) {
                        cf.topBottom.format.fill.color = rule.fill;
                    } else {
                        logDiag(`Invalid topBottom fill color: ${rule.fill}, skipping`);
                    }
                }
                if (rule.fontColor) {
                    if (isValidHexColor(rule.fontColor)) {
                        cf.topBottom.format.font.color = rule.fontColor;
                    } else {
                        logDiag(`Invalid topBottom fontColor: ${rule.fontColor}, skipping`);
                    }
                }
                if (rule.bold) cf.topBottom.format.font.bold = rule.bold;
                
                appliedCount++;
                logDiag(`Applied topBottom (${rule.rule}, rank: ${rank})`);
            }
            
            // ========== Preset ==========
            else if (ruleType === "preset") {
                if (!rule.criterion) {
                    logDiag(`preset requires criterion property`);
                    continue;
                }
                
                if (!validPresetCriteria.includes(rule.criterion)) {
                    logDiag(`Invalid preset criterion: ${rule.criterion}. Valid criteria: ${validPresetCriteria.join(", ")}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
                
                // Map criterion
                const criterionMap = {
                    "duplicateValues": Excel.ConditionalFormatPresetCriterion.duplicateValues,
                    "uniqueValues": Excel.ConditionalFormatPresetCriterion.uniqueValues,
                    "aboveAverage": Excel.ConditionalFormatPresetCriterion.aboveAverage,
                    "belowAverage": Excel.ConditionalFormatPresetCriterion.belowAverage,
                    "equalOrAboveAverage": Excel.ConditionalFormatPresetCriterion.equalOrAboveAverage,
                    "equalOrBelowAverage": Excel.ConditionalFormatPresetCriterion.equalOrBelowAverage,
                    "oneStdDevAboveAverage": Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage,
                    "oneStdDevBelowAverage": Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage,
                    "twoStdDevAboveAverage": Excel.ConditionalFormatPresetCriterion.twoStdDevAboveAverage,
                    "twoStdDevBelowAverage": Excel.ConditionalFormatPresetCriterion.twoStdDevBelowAverage,
                    "threeStdDevAboveAverage": Excel.ConditionalFormatPresetCriterion.threeStdDevAboveAverage,
                    "threeStdDevBelowAverage": Excel.ConditionalFormatPresetCriterion.threeStdDevBelowAverage,
                    "yesterday": Excel.ConditionalFormatPresetCriterion.yesterday,
                    "today": Excel.ConditionalFormatPresetCriterion.today,
                    "tomorrow": Excel.ConditionalFormatPresetCriterion.tomorrow,
                    "lastSevenDays": Excel.ConditionalFormatPresetCriterion.lastSevenDays,
                    "lastWeek": Excel.ConditionalFormatPresetCriterion.lastWeek,
                    "thisWeek": Excel.ConditionalFormatPresetCriterion.thisWeek,
                    "nextWeek": Excel.ConditionalFormatPresetCriterion.nextWeek,
                    "lastMonth": Excel.ConditionalFormatPresetCriterion.lastMonth,
                    "thisMonth": Excel.ConditionalFormatPresetCriterion.thisMonth,
                    "nextMonth": Excel.ConditionalFormatPresetCriterion.nextMonth
                };
                
                cf.preset.rule = {
                    criterion: criterionMap[rule.criterion]
                };
                
                // Apply formatting with color validation
                if (rule.fill) {
                    if (isValidHexColor(rule.fill)) {
                        cf.preset.format.fill.color = rule.fill;
                    } else {
                        logDiag(`Invalid preset fill color: ${rule.fill}, skipping`);
                    }
                }
                if (rule.fontColor) {
                    if (isValidHexColor(rule.fontColor)) {
                        cf.preset.format.font.color = rule.fontColor;
                    } else {
                        logDiag(`Invalid preset fontColor: ${rule.fontColor}, skipping`);
                    }
                }
                if (rule.bold) cf.preset.format.font.bold = rule.bold;
                
                appliedCount++;
                logDiag(`Applied preset (${rule.criterion})`);
            }
            
            // ========== Text Comparison ==========
            else if (ruleType === "textComparison") {
                if (!rule.operator || !rule.text) {
                    logDiag(`textComparison requires operator and text properties`);
                    continue;
                }
                
                const validOperators = ["contains", "notContains", "beginsWith", "endsWith"];
                if (!validOperators.includes(rule.operator)) {
                    logDiag(`Invalid textComparison operator: ${rule.operator}. Valid: ${validOperators.join(", ")}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
                
                // Map operator
                const operatorMap = {
                    "contains": Excel.ConditionalTextOperator.contains,
                    "notContains": Excel.ConditionalTextOperator.notContains,
                    "beginsWith": Excel.ConditionalTextOperator.beginsWith,
                    "endsWith": Excel.ConditionalTextOperator.endsWith
                };
                
                cf.textComparison.rule = {
                    operator: operatorMap[rule.operator],
                    text: rule.text
                };
                
                // Apply formatting with color validation
                if (rule.fill) {
                    if (isValidHexColor(rule.fill)) {
                        cf.textComparison.format.fill.color = rule.fill;
                    } else {
                        logDiag(`Invalid textComparison fill color: ${rule.fill}, skipping`);
                    }
                }
                if (rule.fontColor) {
                    if (isValidHexColor(rule.fontColor)) {
                        cf.textComparison.format.font.color = rule.fontColor;
                    } else {
                        logDiag(`Invalid textComparison fontColor: ${rule.fontColor}, skipping`);
                    }
                }
                if (rule.bold) cf.textComparison.format.font.bold = rule.bold;
                
                appliedCount++;
                logDiag(`Applied textComparison (${rule.operator}: "${rule.text}")`);
            }
            
            // ========== Custom Formula ==========
            else if (ruleType === "custom") {
                if (!rule.formula) {
                    logDiag(`custom requires formula property`);
                    continue;
                }
                
                if (!rule.formula.startsWith("=")) {
                    logDiag(`custom formula must start with "=": ${rule.formula}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                
                cf.custom.rule = {
                    formula: rule.formula
                };
                
                // Apply formatting with color validation
                if (rule.fill) {
                    if (isValidHexColor(rule.fill)) {
                        cf.custom.format.fill.color = rule.fill;
                    } else {
                        logDiag(`Invalid custom fill color: ${rule.fill}, skipping`);
                    }
                }
                if (rule.fontColor) {
                    if (isValidHexColor(rule.fontColor)) {
                        cf.custom.format.font.color = rule.fontColor;
                    } else {
                        logDiag(`Invalid custom fontColor: ${rule.fontColor}, skipping`);
                    }
                }
                if (rule.bold) cf.custom.format.font.bold = rule.bold;
                if (rule.italic) cf.custom.format.font.italic = rule.italic;
                
                appliedCount++;
                const formulaPreview = rule.formula.length > 50 ? rule.formula.substring(0, 50) + "..." : rule.formula;
                logDiag(`Applied custom formula (${formulaPreview})`);
            }
            
        } catch (ruleError) {
            logDiag(`Error applying conditional format rule: ${ruleError.message}`);
        }
    }
    
    logDiag(`Applied ${appliedCount} conditional format rule(s)`);
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
    
    // Parse advanced chart options from action.data
    // Supports both JSON string (from AI-generated ACTION tags) and plain objects (programmatic calls)
    let advancedOptions = {};
    if (data) {
        if (typeof data === "string") {
            try {
                advancedOptions = JSON.parse(data);
            } catch (e) {
                logDiag(`Warning: Could not parse advanced chart options: ${e.message}`);
            }
        } else if (typeof data === "object") {
            advancedOptions = data;
        }
    }
    
    // Check if any series-level operations are needed (trendlines, dataLabels, comboSeries)
    const needsSeriesLoad = (advancedOptions.trendlines && Array.isArray(advancedOptions.trendlines) && advancedOptions.trendlines.length > 0) ||
                           advancedOptions.dataLabels ||
                           (advancedOptions.comboSeries && Array.isArray(advancedOptions.comboSeries) && advancedOptions.comboSeries.length > 0);
    
    // Load series once if needed for any series-level operations
    if (needsSeriesLoad) {
        chart.series.load("items");
        await ctx.sync();
    }
    
    // ========== Trendline Support ==========
    if (advancedOptions.trendlines && Array.isArray(advancedOptions.trendlines) && advancedOptions.trendlines.length > 0) {
        try {
            const trendlineTypeMap = {
                "Linear": Excel.ChartTrendlineType.linear,
                "Exponential": Excel.ChartTrendlineType.exponential,
                "Polynomial": Excel.ChartTrendlineType.polynomial,
                "MovingAverage": Excel.ChartTrendlineType.movingAverage
            };
            
            for (const trendlineConfig of advancedOptions.trendlines) {
                const seriesIndex = trendlineConfig.seriesIndex || 0;
                const trendlineType = trendlineConfig.type || "Linear";
                
                if (seriesIndex >= 0 && seriesIndex < chart.series.items.length) {
                    const series = chart.series.items[seriesIndex];
                    const trendline = series.trendlines.add(trendlineTypeMap[trendlineType] || Excel.ChartTrendlineType.linear);
                    
                    if (trendlineType === "MovingAverage" && trendlineConfig.period) {
                        trendline.movingAveragePeriod = trendlineConfig.period;
                    }
                    if (trendlineType === "Polynomial" && trendlineConfig.order) {
                        trendline.polynomialOrder = trendlineConfig.order;
                    }
                    
                    logDiag(`Added ${trendlineType} trendline to series ${seriesIndex}`);
                } else {
                    logDiag(`Warning: Invalid seriesIndex ${seriesIndex} for trendline, skipping`);
                }
            }
        } catch (e) {
            logDiag(`Warning: Trendline customization error: ${e.message}`);
        }
    }
    
    // ========== Data Label Customization ==========
    if (advancedOptions.dataLabels) {
        try {
            const dataLabelPositionMap = {
                "Center": Excel.ChartDataLabelPosition.center,
                "InsideEnd": Excel.ChartDataLabelPosition.insideEnd,
                "OutsideEnd": Excel.ChartDataLabelPosition.outsideEnd,
                "InsideBase": Excel.ChartDataLabelPosition.insideBase,
                "BestFit": Excel.ChartDataLabelPosition.bestFit,
                "Left": Excel.ChartDataLabelPosition.left,
                "Right": Excel.ChartDataLabelPosition.right,
                "Top": Excel.ChartDataLabelPosition.top,
                "Bottom": Excel.ChartDataLabelPosition.bottom
            };
            
            for (const series of chart.series.items) {
                series.hasDataLabels = true;
                const labels = series.dataLabels;
                
                if (advancedOptions.dataLabels.position && dataLabelPositionMap[advancedOptions.dataLabels.position]) {
                    labels.position = dataLabelPositionMap[advancedOptions.dataLabels.position];
                }
                if (advancedOptions.dataLabels.showValue !== undefined) {
                    labels.showValue = advancedOptions.dataLabels.showValue;
                }
                if (advancedOptions.dataLabels.showSeriesName !== undefined) {
                    labels.showSeriesName = advancedOptions.dataLabels.showSeriesName;
                }
                if (advancedOptions.dataLabels.showCategoryName !== undefined) {
                    labels.showCategoryName = advancedOptions.dataLabels.showCategoryName;
                }
                if (advancedOptions.dataLabels.showLegendKey !== undefined) {
                    labels.showLegendKey = advancedOptions.dataLabels.showLegendKey;
                }
                if (advancedOptions.dataLabels.showPercentage !== undefined) {
                    labels.showPercentage = advancedOptions.dataLabels.showPercentage;
                }
                if (advancedOptions.dataLabels.numberFormat) {
                    labels.numberFormat = advancedOptions.dataLabels.numberFormat;
                }
                
                // Font formatting for data labels
                if (advancedOptions.dataLabels.format && advancedOptions.dataLabels.format.font) {
                    const font = advancedOptions.dataLabels.format.font;
                    if (font.bold !== undefined) labels.format.font.bold = font.bold;
                    if (font.color) labels.format.font.color = font.color;
                    if (font.size) labels.format.font.size = font.size;
                }
            }
            
            logDiag(`Applied data labels: position=${advancedOptions.dataLabels.position || 'default'}`);
        } catch (e) {
            logDiag(`Warning: Data label customization error: ${e.message}`);
        }
    }
    
    // ========== Axis Formatting ==========
    if (advancedOptions.axes) {
        try {
            const displayUnitMap = {
                "Hundreds": Excel.ChartAxisDisplayUnit.hundreds,
                "Thousands": Excel.ChartAxisDisplayUnit.thousands,
                "TenThousands": Excel.ChartAxisDisplayUnit.tenThousands,
                "HundredThousands": Excel.ChartAxisDisplayUnit.hundredThousands,
                "Millions": Excel.ChartAxisDisplayUnit.millions,
                "TenMillions": Excel.ChartAxisDisplayUnit.tenMillions,
                "HundredMillions": Excel.ChartAxisDisplayUnit.hundredMillions,
                "Billions": Excel.ChartAxisDisplayUnit.billions
            };
            
            // Category axis (X-axis)
            if (advancedOptions.axes.category) {
                const catAxis = chart.axes.categoryAxis;
                const catConfig = advancedOptions.axes.category;
                
                if (catConfig.title) {
                    catAxis.title.text = catConfig.title;
                    catAxis.title.visible = true;
                }
                if (catConfig.gridlines !== undefined) {
                    catAxis.majorGridlines.visible = catConfig.gridlines;
                }
                if (catConfig.format && catConfig.format.font) {
                    const font = catConfig.format.font;
                    if (font.bold !== undefined) catAxis.format.font.bold = font.bold;
                    if (font.color) catAxis.format.font.color = font.color;
                    if (font.size) catAxis.format.font.size = font.size;
                }
                
                logDiag(`Applied category axis formatting: title="${catConfig.title || 'none'}"`);
            }
            
            // Value axis (Y-axis)
            if (advancedOptions.axes.value) {
                const valAxis = chart.axes.valueAxis;
                const valConfig = advancedOptions.axes.value;
                
                if (valConfig.title) {
                    valAxis.title.text = valConfig.title;
                    valAxis.title.visible = true;
                }
                if (valConfig.displayUnit && displayUnitMap[valConfig.displayUnit]) {
                    valAxis.displayUnit = displayUnitMap[valConfig.displayUnit];
                }
                if (valConfig.gridlines !== undefined) {
                    valAxis.majorGridlines.visible = valConfig.gridlines;
                }
                if (valConfig.minimum !== undefined) {
                    valAxis.minimum = valConfig.minimum;
                }
                if (valConfig.maximum !== undefined) {
                    valAxis.maximum = valConfig.maximum;
                }
                if (valConfig.format && valConfig.format.font) {
                    const font = valConfig.format.font;
                    if (font.bold !== undefined) valAxis.format.font.bold = font.bold;
                    if (font.color) valAxis.format.font.color = font.color;
                    if (font.size) valAxis.format.font.size = font.size;
                }
                
                logDiag(`Applied value axis formatting: title="${valConfig.title || 'none'}", displayUnit="${valConfig.displayUnit || 'none'}"`);
            }
        } catch (e) {
            logDiag(`Warning: Axis formatting error: ${e.message}`);
        }
    }
    
    // ========== Chart Element Formatting ==========
    if (advancedOptions.formatting) {
        try {
            // Title formatting
            if (advancedOptions.formatting.title && advancedOptions.formatting.title.font) {
                const font = advancedOptions.formatting.title.font;
                if (font.bold !== undefined) chart.title.format.font.bold = font.bold;
                if (font.color) chart.title.format.font.color = font.color;
                if (font.size) chart.title.format.font.size = font.size;
                if (font.italic !== undefined) chart.title.format.font.italic = font.italic;
                
                logDiag(`Applied title formatting: bold=${font.bold}, color=${font.color}, size=${font.size}`);
            }
            
            // Legend formatting
            if (advancedOptions.formatting.legend) {
                const legendConfig = advancedOptions.formatting.legend;
                
                const legendPositionMap = {
                    "Top": Excel.ChartLegendPosition.top,
                    "Bottom": Excel.ChartLegendPosition.bottom,
                    "Left": Excel.ChartLegendPosition.left,
                    "Right": Excel.ChartLegendPosition.right,
                    "Corner": Excel.ChartLegendPosition.corner,
                    "Custom": Excel.ChartLegendPosition.custom
                };
                
                if (legendConfig.position && legendPositionMap[legendConfig.position]) {
                    chart.legend.position = legendPositionMap[legendConfig.position];
                }
                if (legendConfig.font) {
                    if (legendConfig.font.bold !== undefined) chart.legend.format.font.bold = legendConfig.font.bold;
                    if (legendConfig.font.color) chart.legend.format.font.color = legendConfig.font.color;
                    if (legendConfig.font.size) chart.legend.format.font.size = legendConfig.font.size;
                }
                
                logDiag(`Applied legend formatting: position=${legendConfig.position || 'default'}`);
            }
            
            // Chart area formatting (fill and border)
            if (advancedOptions.formatting.chartArea) {
                if (advancedOptions.formatting.chartArea.fill) {
                    chart.format.fill.setSolidColor(advancedOptions.formatting.chartArea.fill);
                    logDiag(`Applied chart area fill: ${advancedOptions.formatting.chartArea.fill}`);
                }
                
                // Chart area border customization
                if (advancedOptions.formatting.chartArea.border) {
                    const borderConfig = advancedOptions.formatting.chartArea.border;
                    const chartLine = chart.format.border;
                    
                    if (borderConfig.color) {
                        chartLine.color = borderConfig.color;
                    }
                    if (borderConfig.weight !== undefined) {
                        chartLine.weight = borderConfig.weight;
                    }
                    if (borderConfig.lineStyle) {
                        const lineStyleMap = {
                            "Continuous": Excel.ChartLineStyle.continuous,
                            "Dash": Excel.ChartLineStyle.dash,
                            "DashDot": Excel.ChartLineStyle.dashDot,
                            "DashDotDot": Excel.ChartLineStyle.dashDotDot,
                            "Dot": Excel.ChartLineStyle.dot,
                            "Grey25": Excel.ChartLineStyle.grey25,
                            "Grey50": Excel.ChartLineStyle.grey50,
                            "Grey75": Excel.ChartLineStyle.grey75,
                            "Automatic": Excel.ChartLineStyle.automatic,
                            "None": Excel.ChartLineStyle.none
                        };
                        if (lineStyleMap[borderConfig.lineStyle]) {
                            chartLine.lineStyle = lineStyleMap[borderConfig.lineStyle];
                        }
                    }
                    
                    logDiag(`Applied chart area border: color=${borderConfig.color || 'default'}, weight=${borderConfig.weight || 'default'}, style=${borderConfig.lineStyle || 'default'}`);
                }
            }
            
            // Plot area formatting (fill and border)
            if (advancedOptions.formatting.plotArea) {
                const plotArea = chart.plotArea;
                
                if (advancedOptions.formatting.plotArea.fill) {
                    plotArea.format.fill.setSolidColor(advancedOptions.formatting.plotArea.fill);
                    logDiag(`Applied plot area fill: ${advancedOptions.formatting.plotArea.fill}`);
                }
                
                if (advancedOptions.formatting.plotArea.border) {
                    const borderConfig = advancedOptions.formatting.plotArea.border;
                    const plotLine = plotArea.format.border;
                    
                    if (borderConfig.color) {
                        plotLine.color = borderConfig.color;
                    }
                    if (borderConfig.weight !== undefined) {
                        plotLine.weight = borderConfig.weight;
                    }
                    if (borderConfig.lineStyle) {
                        const lineStyleMap = {
                            "Continuous": Excel.ChartLineStyle.continuous,
                            "Dash": Excel.ChartLineStyle.dash,
                            "DashDot": Excel.ChartLineStyle.dashDot,
                            "DashDotDot": Excel.ChartLineStyle.dashDotDot,
                            "Dot": Excel.ChartLineStyle.dot,
                            "None": Excel.ChartLineStyle.none
                        };
                        if (lineStyleMap[borderConfig.lineStyle]) {
                            plotLine.lineStyle = lineStyleMap[borderConfig.lineStyle];
                        }
                    }
                    
                    logDiag(`Applied plot area border: color=${borderConfig.color || 'default'}, weight=${borderConfig.weight || 'default'}`);
                }
            }
        } catch (e) {
            logDiag(`Warning: Chart element formatting error: ${e.message}`);
        }
    }
    
    // ========== Combo Chart / Secondary Axis Support ==========
    // Note: Series already loaded above if needsSeriesLoad was true
    if (advancedOptions.comboSeries && Array.isArray(advancedOptions.comboSeries) && advancedOptions.comboSeries.length > 0) {
        try {
            const comboChartTypeMap = {
                "Line": Excel.ChartType.line,
                "ColumnClustered": Excel.ChartType.columnClustered,
                "ColumnStacked": Excel.ChartType.columnStacked,
                "Area": Excel.ChartType.area,
                "AreaStacked": Excel.ChartType.areaStacked,
                "Scatter": Excel.ChartType.xyscatter
            };
            
            const axisGroupMap = {
                "Primary": Excel.ChartAxisGroup.primary,
                "Secondary": Excel.ChartAxisGroup.secondary
            };
            
            for (const comboConfig of advancedOptions.comboSeries) {
                const seriesIndex = comboConfig.seriesIndex;
                
                if (seriesIndex >= 0 && seriesIndex < chart.series.items.length) {
                    const series = chart.series.items[seriesIndex];
                    
                    if (comboConfig.chartType && comboChartTypeMap[comboConfig.chartType]) {
                        series.chartType = comboChartTypeMap[comboConfig.chartType];
                    }
                    if (comboConfig.axisGroup && axisGroupMap[comboConfig.axisGroup]) {
                        series.axisGroup = axisGroupMap[comboConfig.axisGroup];
                    }
                    
                    logDiag(`Set series ${seriesIndex} to ${comboConfig.chartType || 'default'} on ${comboConfig.axisGroup || 'Primary'} axis`);
                } else {
                    logDiag(`Warning: Invalid seriesIndex ${seriesIndex} for combo series, skipping`);
                }
            }
            
            // Configure secondary value axis if any series uses it
            if (advancedOptions.axes && advancedOptions.axes.value2) {
                try {
                    const secValAxis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
                    const val2Config = advancedOptions.axes.value2;
                    
                    if (val2Config.title) {
                        secValAxis.title.text = val2Config.title;
                        secValAxis.title.visible = true;
                    }
                    
                    logDiag(`Applied secondary value axis title: "${val2Config.title || 'none'}"`);
                } catch (secAxisError) {
                    logDiag(`Warning: Secondary axis configuration error: ${secAxisError.message}`);
                }
            }
        } catch (e) {
            logDiag(`Warning: Combo chart customization error: ${e.message}`);
        }
    }
    
    // Final sync to apply all chart customizations
    await ctx.sync();
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
// Named Range Operations
// ============================================================================

/**
 * Validates a named range name
 * @param {string} name - Name to validate
 * @returns {Object} Validation result with isValid and error message
 */
function validateNamedRangeName(name) {
    if (!name || typeof name !== "string") {
        return { isValid: false, error: "Named range name is required." };
    }
    
    // Must start with a letter or underscore
    if (!/^[A-Za-z_]/.test(name)) {
        return { isValid: false, error: "Named range name must start with a letter or underscore." };
    }
    
    // Can only contain letters, numbers, underscores, and periods
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) {
        return { isValid: false, error: "Named range name can only contain letters, numbers, underscores, and periods. Spaces are not allowed." };
    }
    
    // Cannot be a cell reference (e.g., A1, XFD1048576)
    if (/^[A-Za-z]{1,3}\d+$/.test(name)) {
        return { isValid: false, error: "Named range name cannot look like a cell reference (e.g., A1, B2)." };
    }
    
    // Max 255 characters
    if (name.length > 255) {
        return { isValid: false, error: "Named range name cannot exceed 255 characters." };
    }
    
    return { isValid: true };
}

/**
 * Creates a named range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with name, scope, formula, comment
 * 
 * For workbook-scoped named ranges referencing other sheets, use one of:
 * 1. Sheet-qualified target: "Sheet2!A1:B5" - will resolve to the correct sheet
 * 2. Formula option: {"formula":"=Sheet2!A1:B5"} - for explicit formula-based references
 * 
 * For worksheet-scoped names, target is always relative to the active sheet.
 */
async function createNamedRange(ctx, sheet, action) {
    logDiag(`Starting createNamedRange for target "${action.target}"`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for createNamedRange`);
        }
    }
    
    const name = options.name;
    const scope = options.scope || "workbook";
    const formula = options.formula;
    const comment = options.comment || "";
    
    // Validate name
    const validation = validateNamedRangeName(name);
    if (!validation.isValid) {
        logDiag(`Error: ${validation.error}`);
        throw new Error(validation.error);
    }
    
    try {
        // Check for existing name
        let existingName;
        if (scope === "worksheet") {
            existingName = sheet.names.getItemOrNullObject(name);
        } else {
            existingName = ctx.workbook.names.getItemOrNullObject(name);
        }
        existingName.load("isNullObject");
        await ctx.sync();
        
        if (!existingName.isNullObject) {
            const errorMsg = `A named range called '${name}' already exists in ${scope} scope. Choose a different name or delete the existing one first.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        // Determine what to add - formula or range reference
        let namedItem;
        if (formula) {
            // Named formula or constant
            const formulaValue = formula.startsWith("=") ? formula : `=${formula}`;
            if (scope === "worksheet") {
                namedItem = sheet.names.add(name, formulaValue, comment);
            } else {
                namedItem = ctx.workbook.names.add(name, formulaValue, comment);
            }
            logDiag(`Creating named formula '${name}' with formula '${formulaValue}'`);
        } else {
            // Named range reference
            if (!action.target) {
                const errorMsg = "Target range is required for named range (e.g., 'A1:E100' or 'Sheet2!A1:B5').";
                logDiag(`Error: ${errorMsg}`);
                throw new Error(errorMsg);
            }
            
            // Check if target contains sheet reference (e.g., "Sheet2!A1:B5")
            let targetRange;
            if (action.target.includes("!")) {
                // Sheet-qualified reference - parse and resolve
                const parts = action.target.split("!");
                const sheetName = parts[0].replace(/^'|'$/g, ""); // Remove quotes if present
                const rangeAddress = parts.slice(1).join("!"); // Handle edge case of ! in range
                
                const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
                targetSheet.load("isNullObject");
                await ctx.sync();
                
                if (targetSheet.isNullObject) {
                    const errorMsg = `Sheet "${sheetName}" not found. Check the sheet name in target "${action.target}".`;
                    logDiag(`Error: ${errorMsg}`);
                    throw new Error(errorMsg);
                }
                
                targetRange = targetSheet.getRange(rangeAddress);
                logDiag(`Resolved cross-sheet reference: ${sheetName}!${rangeAddress}`);
            } else {
                // Local range on active sheet
                targetRange = sheet.getRange(action.target);
            }
            
            if (scope === "worksheet") {
                namedItem = sheet.names.add(name, targetRange, comment);
            } else {
                namedItem = ctx.workbook.names.add(name, targetRange, comment);
            }
            logDiag(`Creating named range '${name}' for range '${action.target}'`);
        }
        
        await ctx.sync();
        logDiag(`Successfully created named range '${name}' with ${scope} scope`);
    } catch (e) {
        const errorMsg = `Failed to create named range: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes a named range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with named range name as target
 */
async function deleteNamedRange(ctx, sheet, action) {
    logDiag(`Starting deleteNamedRange for target "${action.target}"`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for deleteNamedRange`);
        }
    }
    
    const name = action.target;
    const scope = options.scope || "workbook";
    
    if (!name) {
        const errorMsg = "Named range name is required.";
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        let namedItem;
        if (scope === "worksheet") {
            namedItem = sheet.names.getItemOrNullObject(name);
        } else {
            namedItem = ctx.workbook.names.getItemOrNullObject(name);
        }
        namedItem.load("isNullObject");
        await ctx.sync();
        
        if (namedItem.isNullObject) {
            logDiag(`Warning: Named range '${name}' not found in ${scope} scope. Nothing to delete.`);
            return;
        }
        
        namedItem.delete();
        await ctx.sync();
        logDiag(`Successfully deleted named range '${name}' from ${scope} scope`);
    } catch (e) {
        const errorMsg = `Failed to delete named range: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Updates a named range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with named range name as target, newFormula, newComment
 */
async function updateNamedRange(ctx, sheet, action) {
    logDiag(`Starting updateNamedRange for target "${action.target}"`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for updateNamedRange`);
        }
    }
    
    const name = action.target;
    const scope = options.scope || "workbook";
    const newFormula = options.newFormula;
    const newComment = options.newComment;
    
    if (!name) {
        const errorMsg = "Named range name is required.";
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    if (newFormula === undefined && newComment === undefined) {
        const errorMsg = "At least one of newFormula or newComment must be provided.";
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
    
    try {
        let namedItem;
        if (scope === "worksheet") {
            namedItem = sheet.names.getItemOrNullObject(name);
        } else {
            namedItem = ctx.workbook.names.getItemOrNullObject(name);
        }
        namedItem.load(["isNullObject", "formula", "comment"]);
        await ctx.sync();
        
        if (namedItem.isNullObject) {
            const errorMsg = `Named range '${name}' not found in ${scope} scope. Use listNamedRanges to see available names.`;
            logDiag(`Error: ${errorMsg}`);
            throw new Error(errorMsg);
        }
        
        const updates = [];
        if (newFormula !== undefined) {
            const formulaValue = newFormula.startsWith("=") ? newFormula : `=${newFormula}`;
            namedItem.formula = formulaValue;
            updates.push(`formula=${formulaValue}`);
        }
        if (newComment !== undefined) {
            namedItem.comment = newComment;
            updates.push(`comment=${newComment}`);
        }
        
        await ctx.sync();
        logDiag(`Successfully updated named range '${name}': ${updates.join(", ")}`);
    } catch (e) {
        const errorMsg = `Failed to update named range: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

/**
 * Lists named ranges (diagnostics-only)
 * 
 * NOTE: This action is primarily for diagnostics and debugging purposes.
 * Results are logged to the diagnostics panel but are NOT returned to the AI
 * or surfaced in the UI, as the executeAction architecture does not currently
 * support action return values. The AI can reference named ranges through the
 * data context built by excel-data.js which includes existing named ranges.
 * 
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with scope option
 * @returns {Promise<Array>} Array of named range objects (for internal use only)
 */
async function listNamedRanges(ctx, sheet, action) {
    logDiag(`Starting listNamedRanges (diagnostics-only)`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logDiag(`Warning: Failed to parse action.data for listNamedRanges`);
        }
    }
    
    const scope = options.scope || "all";
    
    try {
        const results = [];
        
        // Load workbook-scoped names
        if (scope === "all" || scope === "workbook") {
            ctx.workbook.names.load("items");
            await ctx.sync();
            
            for (const item of ctx.workbook.names.items) {
                item.load(["name", "formula", "comment", "type", "visible"]);
            }
            await ctx.sync();
            
            for (const item of ctx.workbook.names.items) {
                results.push({
                    name: item.name,
                    scope: "workbook",
                    formula: item.formula,
                    comment: item.comment || "",
                    type: item.type,
                    visible: item.visible
                });
            }
            logDiag(`Found ${ctx.workbook.names.items.length} workbook-scoped named ranges`);
        }
        
        // Load worksheet-scoped names
        if (scope === "all" || scope === "worksheet") {
            sheet.names.load("items");
            await ctx.sync();
            
            for (const item of sheet.names.items) {
                item.load(["name", "formula", "comment", "type", "visible"]);
            }
            await ctx.sync();
            
            for (const item of sheet.names.items) {
                results.push({
                    name: item.name,
                    scope: "worksheet",
                    sheetName: sheet.name,
                    formula: item.formula,
                    comment: item.comment || "",
                    type: item.type,
                    visible: item.visible
                });
            }
            logDiag(`Found ${sheet.names.items.length} worksheet-scoped named ranges`);
        }
        
        // Log results
        if (results.length === 0) {
            logDiag("No named ranges found.");
        } else {
            logDiag(`=== Named Ranges (${results.length} total) ===`);
            for (const nr of results) {
                const scopeInfo = nr.scope === "worksheet" ? `worksheet:${nr.sheetName}` : "workbook";
                logDiag(`  ${nr.name} [${scopeInfo}]: ${nr.formula}${nr.comment ? ` (${nr.comment})` : ""}`);
            }
        }
        
        return results;
    } catch (e) {
        const errorMsg = `Failed to list named ranges: ${e.message}`;
        logDiag(`Error: ${errorMsg}`);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// Protection Operations
// ============================================================================

/**
 * Protects a worksheet with optional password and permissions
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (sheet name) and data (JSON with password, options)
 */
async function protectWorksheet(ctx, sheet, action) {
    // Guard: Ensure sheet context is available
    if (!sheet) {
        logDiag(`[protectWorksheet] Error: Sheet context is not available`);
        throw new Error("protectWorksheet requires a valid sheet context");
    }
    
    logDiag(`[protectWorksheet] Protecting worksheet: ${action.target || sheet.name}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        const targetSheetName = action.target || sheet.name;
        
        // Guard: Ensure target sheet name is valid
        if (!targetSheetName || typeof targetSheetName !== "string" || targetSheetName.trim() === "") {
            logDiag(`[protectWorksheet] Error: Invalid or empty sheet name`);
            throw new Error("protectWorksheet requires a valid sheet name in action.target or current sheet context");
        }
        
        // Get target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(targetSheetName);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${targetSheetName}" not found`);
        }
        
        // Check if already protected
        const protection = targetSheet.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (protection.protected) {
            throw new Error(`Sheet "${targetSheetName}" is already protected. Unprotect it first.`);
        }
        
        // Build protection options
        // Note: All options default to false (most restrictive) except allowAutoFilter which defaults to true
        // for usability. This is documented in the AI prompts.
        const options = {
            allowAutoFilter: data.allowAutoFilter !== false, // Default true for usability
            allowDeleteColumns: data.allowDeleteColumns === true,
            allowDeleteRows: data.allowDeleteRows === true,
            allowFormatCells: data.allowFormatCells === true,
            allowFormatColumns: data.allowFormatColumns === true,
            allowFormatRows: data.allowFormatRows === true,
            allowInsertColumns: data.allowInsertColumns === true,
            allowInsertRows: data.allowInsertRows === true,
            allowInsertHyperlinks: data.allowInsertHyperlinks === true,
            allowPivotTables: data.allowPivotTables === true,
            allowSort: data.allowSort === true,
            selectionMode: data.selectionMode || "Normal" // "Normal", "Unlocked", "None"
        };
        
        // Apply protection
        const password = data.password || undefined;
        protection.protect(options, password);
        await ctx.sync();
        
        logDiag(`[protectWorksheet] Successfully protected "${targetSheetName}"`);
    } catch (error) {
        logDiag(`[protectWorksheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unprotects a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (sheet name) and data (JSON with password)
 */
async function unprotectWorksheet(ctx, sheet, action) {
    // Guard: Ensure sheet context is available
    if (!sheet) {
        logDiag(`[unprotectWorksheet] Error: Sheet context is not available`);
        throw new Error("unprotectWorksheet requires a valid sheet context");
    }
    
    logDiag(`[unprotectWorksheet] Unprotecting worksheet: ${action.target || sheet.name}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        const targetSheetName = action.target || sheet.name;
        
        // Guard: Ensure target sheet name is valid
        if (!targetSheetName || typeof targetSheetName !== "string" || targetSheetName.trim() === "") {
            logDiag(`[unprotectWorksheet] Error: Invalid or empty sheet name`);
            throw new Error("unprotectWorksheet requires a valid sheet name in action.target or current sheet context");
        }
        
        // Get target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(targetSheetName);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${targetSheetName}" not found`);
        }
        
        // Check if protected
        const protection = targetSheet.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (!protection.protected) {
            logDiag(`[unprotectWorksheet] Sheet "${targetSheetName}" is not protected, skipping`);
            return;
        }
        
        // Unprotect with password if provided
        const password = data.password || undefined;
        protection.unprotect(password);
        await ctx.sync();
        
        logDiag(`[unprotectWorksheet] Successfully unprotected "${targetSheetName}"`);
    } catch (error) {
        logDiag(`[unprotectWorksheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Protects a range by locking cells (requires worksheet protection to take effect)
 * 
 * NOTE: This executor only supports cell-level locking and formula hiding for ranges.
 * Office.js does not support per-user range permissions (allowedUsers, allowedEditors, etc.).
 * Such fields will be ignored if provided.
 * 
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (range) and data (JSON with locked, formulaHidden)
 */
async function protectRange(ctx, sheet, action) {
    // Validate action.target before attempting to get range
    if (!action.target || typeof action.target !== "string" || action.target.trim() === "") {
        logDiag(`[protectRange] Error: Missing or invalid range address in action.target`);
        throw new Error("protectRange requires a valid range address in action.target");
    }
    
    logDiag(`[protectRange] Protecting range: ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Warn about unsupported user-permission fields (Office.js limitation)
        const unsupportedFields = ["allowedUsers", "allowedEditors", "users", "editors", "permissions"];
        const foundUnsupported = unsupportedFields.filter(field => data[field] !== undefined);
        if (foundUnsupported.length > 0) {
            logDiag(`[protectRange] Warning: Office.js does not support per-user range permissions. The following fields will be ignored: ${foundUnsupported.join(", ")}`);
        }
        
        // Get range with error handling for invalid addresses
        let range;
        try {
            range = sheet.getRange(action.target);
        } catch (rangeError) {
            logDiag(`[protectRange] Error: Invalid range address "${action.target}": ${rangeError.message}`);
            throw new Error(`Invalid range address "${action.target}". Please provide a valid Excel range (e.g., "A1:B10", "A:Z", "1:100").`);
        }
        
        // Set protection properties
        range.format.protection.locked = data.locked !== false; // Default true
        range.format.protection.formulaHidden = data.formulaHidden === true; // Default false
        await ctx.sync();
        
        logDiag(`[protectRange] Successfully set protection for "${action.target}" (locked: ${data.locked !== false}, formulaHidden: ${data.formulaHidden === true})`);
        logDiag(`[protectRange] Note: Protection takes effect only when worksheet is protected`);
    } catch (error) {
        logDiag(`[protectRange] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unprotects a range by unlocking cells
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (range)
 */
async function unprotectRange(ctx, sheet, action) {
    // Validate action.target before attempting to get range
    if (!action.target || typeof action.target !== "string" || action.target.trim() === "") {
        logDiag(`[unprotectRange] Error: Missing or invalid range address in action.target`);
        throw new Error("unprotectRange requires a valid range address in action.target");
    }
    
    logDiag(`[unprotectRange] Unprotecting range: ${action.target}`);
    
    try {
        // Get range with error handling for invalid addresses
        let range;
        try {
            range = sheet.getRange(action.target);
        } catch (rangeError) {
            logDiag(`[unprotectRange] Error: Invalid range address "${action.target}": ${rangeError.message}`);
            throw new Error(`Invalid range address "${action.target}". Please provide a valid Excel range (e.g., "A1:B10", "A:Z", "1:100").`);
        }
        
        // Unlock cells and unhide formulas
        range.format.protection.locked = false;
        range.format.protection.formulaHidden = false;
        await ctx.sync();
        
        logDiag(`[unprotectRange] Successfully unlocked "${action.target}"`);
    } catch (error) {
        logDiag(`[unprotectRange] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Protects workbook structure (prevents sheet add/delete/rename/move)
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object (unused, but required by signature)
 * @param {Object} action - Action object with data (JSON with password)
 */
async function protectWorkbook(ctx, sheet, action) {
    logDiag(`[protectWorkbook] Protecting workbook structure`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Check if already protected
        const protection = ctx.workbook.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (protection.protected) {
            throw new Error("Workbook is already protected. Unprotect it first.");
        }
        
        // Apply protection
        const password = data.password || undefined;
        protection.protect(password);
        await ctx.sync();
        
        logDiag(`[protectWorkbook] Successfully protected workbook structure`);
    } catch (error) {
        logDiag(`[protectWorkbook] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unprotects workbook structure
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object (unused, but required by signature)
 * @param {Object} action - Action object with data (JSON with password)
 */
async function unprotectWorkbook(ctx, sheet, action) {
    logDiag(`[unprotectWorkbook] Unprotecting workbook structure`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Check if protected
        const protection = ctx.workbook.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (!protection.protected) {
            logDiag(`[unprotectWorkbook] Workbook is not protected, skipping`);
            return;
        }
        
        // Unprotect with password if provided
        const password = data.password || undefined;
        protection.unprotect(password);
        await ctx.sync();
        
        logDiag(`[unprotectWorkbook] Successfully unprotected workbook structure`);
    } catch (error) {
        logDiag(`[unprotectWorkbook] Error: ${error.message}`);
        throw error;
    }
}

// ============================================================================
// Shape and Image Operations
// ============================================================================

/**
 * Valid geometric shape types supported by Office.js (all lowercase for validation)
 * Original casing is preserved in SHAPE_TYPE_MAP for Office.js API calls
 */
const VALID_SHAPE_TYPES = [
    "rectangle", "oval", "triangle", "righttriangle", "parallelogram", "trapezoid",
    "hexagon", "octagon", "pentagon", "plus", "star4", "star5", "star6",
    "arrow", "chevron", "homeplate", "cube", "bevel", "foldedcorner",
    "smileyface", "donut", "nosmoking", "blockarc", "heart", "lightningbolt",
    "sun", "moon", "cloud", "arc", "bracepair", "bracketpair", "can",
    "flowchartprocess", "flowchartdecision", "flowchartdata", "flowchartterminator",
    "line", "lineinverse", "straightconnector1", "bentconnector2", "bentconnector3"
];

/**
 * Maps lowercase shape types to proper Office.js enum casing
 */
const SHAPE_TYPE_MAP = {
    "rectangle": "Rectangle", "oval": "Oval", "triangle": "Triangle",
    "righttriangle": "RightTriangle", "parallelogram": "Parallelogram", "trapezoid": "Trapezoid",
    "hexagon": "Hexagon", "octagon": "Octagon", "pentagon": "Pentagon",
    "plus": "Plus", "star4": "Star4", "star5": "Star5", "star6": "Star6",
    "arrow": "Arrow", "chevron": "Chevron", "homeplate": "HomePlate",
    "cube": "Cube", "bevel": "Bevel", "foldedcorner": "FoldedCorner",
    "smileyface": "SmileyFace", "donut": "Donut", "nosmoking": "NoSmoking",
    "blockarc": "BlockArc", "heart": "Heart", "lightningbolt": "LightningBolt",
    "sun": "Sun", "moon": "Moon", "cloud": "Cloud", "arc": "Arc",
    "bracepair": "BracePair", "bracketpair": "BracketPair", "can": "Can",
    "flowchartprocess": "FlowchartProcess", "flowchartdecision": "FlowchartDecision",
    "flowchartdata": "FlowchartData", "flowchartterminator": "FlowchartTerminator",
    "line": "Line", "lineinverse": "LineInverse",
    "straightconnector1": "StraightConnector1", "bentconnector2": "BentConnector2", "bentconnector3": "BentConnector3"
};

/**
 * Inserts a geometric shape at a specified cell position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (cell position) and data (JSON with shape options)
 */
async function insertShape(ctx, sheet, action) {
    logDiag(`[insertShape] Starting shape insertion at ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        const shapeType = data.shapeType || "rectangle";
        const normalizedType = shapeType.toLowerCase();
        
        // Validate shape type using normalized lowercase comparison
        if (!VALID_SHAPE_TYPES.includes(normalizedType)) {
            logDiag(`[insertShape] Error: Invalid shape type "${shapeType}"`);
            throw new Error(`Invalid shape type "${shapeType}". Valid types: rectangle, oval, triangle, rightTriangle, arrow, star5, hexagon, line, etc.`);
        }
        
        // Get position from target cell
        let left = 100, top = 100;
        if (action.target) {
            try {
                const posRange = sheet.getRange(action.target);
                posRange.load(["left", "top"]);
                await ctx.sync();
                left = posRange.left;
                top = posRange.top;
            } catch (posError) {
                logDiag(`[insertShape] Warning: Could not parse position "${action.target}", using default`);
            }
        }
        
        // Map normalized shape type to proper Office.js enum casing
        const excelShapeType = SHAPE_TYPE_MAP[normalizedType] || (shapeType.charAt(0).toUpperCase() + shapeType.slice(1));
        
        // Create shape
        const shape = sheet.shapes.addGeometricShape(excelShapeType);
        
        // Set position
        shape.left = left;
        shape.top = top;
        
        // Set dimensions
        const width = data.width || 150;
        const height = data.height || 100;
        if (width <= 0 || height <= 0) {
            throw new Error("Shape dimensions must be positive numbers");
        }
        shape.width = width;
        shape.height = height;
        
        // Set rotation
        if (data.rotation !== undefined) {
            shape.rotation = data.rotation;
        }
        
        // Apply fill color
        if (data.fill && data.fill !== "none") {
            shape.fill.setSolidColor(data.fill);
        } else if (data.fill === "none") {
            shape.fill.clear();
        }
        
        // Apply line/border formatting
        if (data.lineColor && data.lineColor !== "none") {
            shape.lineFormat.color = data.lineColor;
        }
        if (data.lineWeight) {
            shape.lineFormat.weight = data.lineWeight;
        }
        
        // Add text if provided
        if (data.text) {
            shape.textFrame.textRange.text = data.text;
        }
        
        // Set custom name if provided
        if (data.name) {
            shape.name = data.name;
        }
        
        await ctx.sync();
        
        logDiag(`[insertShape] Successfully created ${shapeType} shape at position (${left}, ${top})`);
    } catch (error) {
        logDiag(`[insertShape] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Inserts an image from Base64-encoded data
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (cell position) and data (JSON with image options)
 */
async function insertImage(ctx, sheet, action) {
    logDiag(`[insertImage] Starting image insertion at ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Validate source
        if (!data.source) {
            throw new Error("insertImage requires a Base64-encoded image string in data.source");
        }
        
        // Extract Base64 data (handle data URI format)
        let base64Data = data.source;
        let isSvg = false;
        
        if (base64Data.startsWith("data:image/svg")) {
            isSvg = true;
            // For SVG, we need the XML content, not Base64
            if (base64Data.includes(";base64,")) {
                base64Data = atob(base64Data.split(";base64,")[1]);
            }
        } else if (base64Data.startsWith("data:image/")) {
            // Extract just the Base64 part
            base64Data = base64Data.split(",")[1] || base64Data;
        }
        
        // Get position from target cell
        let left = 100, top = 100;
        if (action.target) {
            try {
                const posRange = sheet.getRange(action.target);
                posRange.load(["left", "top"]);
                await ctx.sync();
                left = posRange.left;
                top = posRange.top;
            } catch (posError) {
                logDiag(`[insertImage] Warning: Could not parse position "${action.target}", using default`);
            }
        }
        
        // Insert image
        let image;
        if (isSvg) {
            image = sheet.shapes.addSvg(base64Data);
        } else {
            image = sheet.shapes.addImage(base64Data);
        }
        
        // Set position
        image.left = left;
        image.top = top;
        
        // Set dimensions
        if (data.width) image.width = data.width;
        if (data.height) image.height = data.height;
        
        // Lock aspect ratio by default
        image.lockAspectRatio = data.lockAspectRatio !== false;
        
        // Set name and alt text
        if (data.name) image.name = data.name;
        if (data.altText) image.altTextDescription = data.altText;
        
        await ctx.sync();
        
        logDiag(`[insertImage] Successfully inserted image at position (${left}, ${top})`);
    } catch (error) {
        logDiag(`[insertImage] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Inserts a text box at a specified cell position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (cell position) and data (JSON with text box options)
 */
async function insertTextBox(ctx, sheet, action) {
    logDiag(`[insertTextBox] Starting text box insertion at ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Validate text
        if (!data.text) {
            throw new Error("insertTextBox requires text content in data.text");
        }
        
        // Get position from target cell
        let left = 100, top = 100;
        if (action.target) {
            try {
                const posRange = sheet.getRange(action.target);
                posRange.load(["left", "top"]);
                await ctx.sync();
                left = posRange.left;
                top = posRange.top;
            } catch (posError) {
                logDiag(`[insertTextBox] Warning: Could not parse position "${action.target}", using default`);
            }
        }
        
        // Create text box (rectangle shape with text)
        const textBox = sheet.shapes.addTextBox(data.text);
        
        // Set position
        textBox.left = left;
        textBox.top = top;
        
        // Set dimensions
        textBox.width = data.width || 150;
        textBox.height = data.height || 50;
        
        // Apply fill
        if (data.fill && data.fill !== "none") {
            textBox.fill.setSolidColor(data.fill);
        } else if (data.fill === "none") {
            textBox.fill.clear();
        }
        
        // Apply border
        if (data.lineColor === "none") {
            textBox.lineFormat.visible = false;
        } else if (data.lineColor) {
            textBox.lineFormat.color = data.lineColor;
        }
        
        // Set name
        if (data.name) textBox.name = data.name;
        
        await ctx.sync();
        
        // Apply text formatting (requires separate sync)
        if (data.fontSize || data.fontColor || data.horizontalAlignment || data.verticalAlignment) {
            const textRange = textBox.textFrame.textRange;
            if (data.fontSize) textRange.font.size = data.fontSize;
            if (data.fontColor) textRange.font.color = data.fontColor;
            if (data.horizontalAlignment) textBox.textFrame.horizontalAlignment = data.horizontalAlignment;
            if (data.verticalAlignment) textBox.textFrame.verticalAlignment = data.verticalAlignment;
            await ctx.sync();
        }
        
        logDiag(`[insertTextBox] Successfully created text box at position (${left}, ${top})`);
    } catch (error) {
        logDiag(`[insertTextBox] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Formats an existing shape
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {string} target - Shape name
 * @param {string} data - JSON string with formatting options
 */
async function formatShape(ctx, sheet, target, data) {
    logDiag(`[formatShape] Formatting shape "${target}"`);
    
    if (!target) {
        throw new Error("formatShape requires a shape name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        
        // Get shape
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load("isNullObject");
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        // Apply fill
        if (options.fill !== undefined) {
            if (options.fill === "none") {
                shape.fill.clear();
            } else {
                shape.fill.setSolidColor(options.fill);
            }
        }
        
        // Apply transparency
        if (options.transparency !== undefined) {
            shape.fill.transparency = Math.max(0, Math.min(1, options.transparency));
        }
        
        // Apply line format
        if (options.lineColor !== undefined) {
            if (options.lineColor === "none") {
                shape.lineFormat.visible = false;
            } else {
                shape.lineFormat.visible = true;
                shape.lineFormat.color = options.lineColor;
            }
        }
        if (options.lineWeight !== undefined) {
            shape.lineFormat.weight = options.lineWeight;
        }
        if (options.lineStyle !== undefined) {
            shape.lineFormat.dashStyle = options.lineStyle;
        }
        
        // Apply dimensions
        if (options.width !== undefined) shape.width = options.width;
        if (options.height !== undefined) shape.height = options.height;
        if (options.rotation !== undefined) shape.rotation = options.rotation;
        
        await ctx.sync();
        
        logDiag(`[formatShape] Successfully formatted shape "${target}"`);
    } catch (error) {
        logDiag(`[formatShape] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Deletes a shape by name
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {string} target - Shape name
 */
async function deleteShape(ctx, sheet, target) {
    logDiag(`[deleteShape] Deleting shape "${target}"`);
    
    if (!target) {
        throw new Error("deleteShape requires a shape name in target");
    }
    
    try {
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load(["isNullObject", "name"]);
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        shape.delete();
        await ctx.sync();
        
        logDiag(`[deleteShape] Successfully deleted shape "${target}"`);
    } catch (error) {
        logDiag(`[deleteShape] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Groups multiple shapes together
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target (comma-separated shape names) and data (JSON with group options)
 */
async function groupShapes(ctx, sheet, action) {
    logDiag(`[groupShapes] Grouping shapes: ${action.target}`);
    
    if (!action.target) {
        throw new Error("groupShapes requires shape names in target (comma-separated)");
    }
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Parse shape names
        const shapeNames = action.target.split(",").map(s => s.trim()).filter(s => s);
        
        if (shapeNames.length < 2) {
            throw new Error("groupShapes requires at least 2 shapes to group");
        }
        
        // Get all shapes and collect their IDs
        const shapes = [];
        for (const name of shapeNames) {
            const shape = sheet.shapes.getItemOrNullObject(name);
            shape.load(["isNullObject", "id"]);
            shapes.push({ name, shape });
        }
        await ctx.sync();
        
        // Validate all shapes exist
        const shapeIds = [];
        for (const { name, shape } of shapes) {
            if (shape.isNullObject) {
                throw new Error(`Shape "${name}" not found`);
            }
            shapeIds.push(shape.id);
        }
        
        // Create group
        const group = sheet.shapes.addGroup(shapeIds);
        
        // Set group name if provided
        if (data.groupName) {
            group.name = data.groupName;
        }
        
        await ctx.sync();
        
        logDiag(`[groupShapes] Successfully grouped ${shapeNames.length} shapes`);
    } catch (error) {
        logDiag(`[groupShapes] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Arranges shape z-order (layering)
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {string} target - Shape name
 * @param {string} data - JSON string with order option
 */
async function arrangeShapes(ctx, sheet, target, data) {
    logDiag(`[arrangeShapes] Arranging shape "${target}"`);
    
    if (!target) {
        throw new Error("arrangeShapes requires a shape name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        
        if (!options.order) {
            throw new Error("arrangeShapes requires an order option: bringToFront, sendToBack, bringForward, sendBackward");
        }
        
        // Get shape
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load("isNullObject");
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        // Map order to Excel enum
        const orderMap = {
            "bringToFront": "BringToFront",
            "sendToBack": "SendToBack",
            "bringForward": "BringForward",
            "sendBackward": "SendBackward"
        };
        
        const excelOrder = orderMap[options.order];
        if (!excelOrder) {
            throw new Error(`Invalid order "${options.order}". Valid options: bringToFront, sendToBack, bringForward, sendBackward`);
        }
        
        shape.incrementZOrder(excelOrder);
        await ctx.sync();
        
        logDiag(`[arrangeShapes] Successfully applied ${options.order} to shape "${target}"`);
    } catch (error) {
        logDiag(`[arrangeShapes] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Ungroups a shape group back into individual shapes
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {string} target - Group shape name
 */
async function ungroupShapes(ctx, sheet, target) {
    logDiag(`[ungroupShapes] Ungrouping shape group "${target}"`);
    
    if (!target) {
        throw new Error("ungroupShapes requires a group name in target");
    }
    
    try {
        // Get the shape group
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load(["isNullObject", "type"]);
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        // Verify it's a group
        if (shape.type !== "Group") {
            throw new Error(`Shape "${target}" is not a group. Only grouped shapes can be ungrouped.`);
        }
        
        // Get the group and ungroup it
        const group = shape.group;
        group.ungroup();
        await ctx.sync();
        
        logDiag(`[ungroupShapes] Successfully ungrouped "${target}"`);
    } catch (error) {
        logDiag(`[ungroupShapes] Error: ${error.message}`);
        throw error;
    }
}

// ============================================================================
// Comments and Notes
// ============================================================================

/**
 * Adds a threaded comment to a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function addComment(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[addComment] Adding comment to cell "${target}"`);
    
    if (!target) {
        throw new Error("addComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            logDiag(`[addComment] Warning: Sheet is protected, comment may not be added`);
        }
        
        const options = data ? JSON.parse(data) : {};
        const content = options.content || options.text || "";
        const contentType = options.contentType === "Mention" ? Excel.ContentType.mention : Excel.ContentType.plain;
        
        if (!content) {
            throw new Error("addComment requires content in data");
        }
        
        // Add comment to the cell
        const comment = sheet.comments.add(target, content, contentType);
        comment.load(["id", "authorName", "creationDate"]);
        await ctx.sync();
        
        logDiag(`[addComment] Successfully added comment (ID: ${comment.id}) to "${target}" by ${comment.authorName}`);
    } catch (error) {
        // Map known errors to user-friendly messages
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("not available") || errorMsg.includes("undefined")) {
            logDiag(`[addComment] API not available: ${errorMsg}`);
            throw new Error("Comments feature is not available in this Excel version");
        }
        logDiag(`[addComment] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Adds a legacy note to a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function addNote(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[addNote] Adding note to cell "${target}"`);
    
    if (!target) {
        throw new Error("addNote requires a cell address in target");
    }
    
    try {
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            logDiag(`[addNote] Warning: Sheet is protected, note may not be added`);
        }
        
        const options = data ? JSON.parse(data) : {};
        const text = options.text || options.content || "";
        
        if (!text) {
            throw new Error("addNote requires text in data");
        }
        
        // Get the range and check if note API is available
        const range = sheet.getRange(target);
        
        // Check if note property exists (requires ExcelApi 1.11+)
        if (range.note === undefined) {
            throw new Error("Notes API is not available in this Excel version (requires ExcelApi 1.11+)");
        }
        
        range.note = text;
        await ctx.sync();
        
        logDiag(`[addNote] Successfully added note to "${target}"`);
    } catch (error) {
        // Map known errors to user-friendly messages
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("not available") || errorMsg.includes("undefined")) {
            logDiag(`[addNote] API not available: ${errorMsg}`);
            throw new Error("Notes feature is not available in this Excel version");
        }
        logDiag(`[addNote] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Edits an existing comment on a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function editComment(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[editComment] Editing comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("editComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot modify comments on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const content = options.content || options.text || "";
        
        if (!content) {
            throw new Error("editComment requires content in data");
        }
        
        // Get comment by cell address
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            throw new Error(`No comment found at cell "${target}"`);
        }
        
        // Update the comment content
        comment.content = content;
        await ctx.sync();
        
        logDiag(`[editComment] Successfully edited comment at "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("protected")) {
            logDiag(`[editComment] Protection error: ${errorMsg}`);
        }
        logDiag(`[editComment] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Edits an existing note on a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function editNote(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[editNote] Editing note at cell "${target}"`);
    
    if (!target) {
        throw new Error("editNote requires a cell address in target");
    }
    
    try {
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot modify notes on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const text = options.text || options.content || "";
        
        if (!text) {
            throw new Error("editNote requires text in data");
        }
        
        // Get the range and check if note API is available
        const range = sheet.getRange(target);
        
        // Check if note property exists
        if (range.note === undefined) {
            throw new Error("Notes API is not available in this Excel version (requires ExcelApi 1.11+)");
        }
        
        range.note = text;
        await ctx.sync();
        
        logDiag(`[editNote] Successfully edited note at "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("protected")) {
            logDiag(`[editNote] Protection error: ${errorMsg}`);
        }
        logDiag(`[editNote] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Deletes a comment from a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target
 */
async function deleteComment(ctx, sheet, action) {
    const { target } = action;
    logDiag(`[deleteComment] Deleting comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("deleteComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot delete comments on a protected sheet");
        }
        
        // Get comment by cell address using getItemByCell
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            logDiag(`[deleteComment] No comment found at "${target}" - nothing to delete`);
            return;
        }
        
        // Delete the comment
        comment.delete();
        await ctx.sync();
        
        logDiag(`[deleteComment] Successfully deleted comment at "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("protected")) {
            logDiag(`[deleteComment] Protection error: ${errorMsg}`);
        }
        logDiag(`[deleteComment] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Deletes a note from a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target
 */
async function deleteNote(ctx, sheet, action) {
    const { target } = action;
    logDiag(`[deleteNote] Deleting note at cell "${target}"`);
    
    if (!target) {
        throw new Error("deleteNote requires a cell address in target");
    }
    
    try {
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot delete notes on a protected sheet");
        }
        
        // Get the range and check if note API is available
        const range = sheet.getRange(target);
        
        // Check if note property exists
        if (range.note === undefined) {
            throw new Error("Notes API is not available in this Excel version (requires ExcelApi 1.11+)");
        }
        
        range.note = "";
        await ctx.sync();
        
        logDiag(`[deleteNote] Successfully deleted note at "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("protected")) {
            logDiag(`[deleteNote] Protection error: ${errorMsg}`);
        }
        logDiag(`[deleteNote] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Adds a reply to an existing comment thread
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function replyToComment(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[replyToComment] Adding reply to comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("replyToComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot reply to comments on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const content = options.content || options.text || "";
        const contentType = options.contentType === "Mention" ? Excel.ContentType.mention : Excel.ContentType.plain;
        
        if (!content) {
            throw new Error("replyToComment requires content in data");
        }
        
        // Get the parent comment by cell address
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            throw new Error(`No comment found at cell "${target}" to reply to`);
        }
        
        // Add reply to the comment
        const reply = comment.replies.add(content, contentType);
        reply.load(["id", "authorName", "creationDate"]);
        await ctx.sync();
        
        logDiag(`[replyToComment] Successfully added reply (ID: ${reply.id}) to comment at "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("protected")) {
            logDiag(`[replyToComment] Protection error: ${errorMsg}`);
        }
        logDiag(`[replyToComment] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Resolves or reopens a comment thread
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function resolveComment(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[resolveComment] Resolving/reopening comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("resolveComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot modify comments on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const resolved = options.resolved !== false; // Default to true
        
        // Get comment by cell address
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            throw new Error(`No comment found at cell "${target}"`);
        }
        
        // Set resolution status
        comment.resolved = resolved;
        await ctx.sync();
        
        logDiag(`[resolveComment] Successfully ${resolved ? "resolved" : "reopened"} comment at "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        if (errorMsg.includes("protected")) {
            logDiag(`[resolveComment] Protection error: ${errorMsg}`);
        }
        logDiag(`[resolveComment] Error: ${errorMsg}`);
        throw error;
    }
}

// ============================================================================
// Sparkline Operations
// ============================================================================

/**
 * Checks if sparkline API is supported in the current Excel version
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @returns {Promise<boolean>} True if sparklines are supported
 */
async function isSparklineSupported(ctx, sheet) {
    try {
        // Check if sparklineGroups API exists
        if (!sheet.sparklineGroups) {
            return false;
        }
        sheet.sparklineGroups.load("count");
        await ctx.sync();
        return true;
    } catch (error) {
        return false;
    }
}

/**
 * Creates sparkline(s) at the specified location
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 * 
 * Data JSON Schema:
 * {
 *   type: "Line" | "Column" | "WinLoss",
 *   sourceData: "B2:F2" (contiguous range),
 *   axes?: { horizontal: boolean, min?: number, max?: number },
 *   markers?: { high: boolean, low: boolean, first: boolean, last: boolean, negative: boolean },
 *   colors?: { series: "#hex", negative: "#hex", high: "#hex", low: "#hex", first: "#hex", last: "#hex" }
 * }
 * 
 * @example
 * // Basic line sparkline
 * <ACTION type="createSparkline" target="G2">{"type":"Line","sourceData":"B2:F2"}</ACTION>
 * 
 * // Column sparkline with custom colors
 * <ACTION type="createSparkline" target="H3">{"type":"Column","sourceData":"C3:C14","colors":{"series":"#70AD47"}}</ACTION>
 * 
 * // Win/Loss sparkline
 * <ACTION type="createSparkline" target="I5">{"type":"WinLoss","sourceData":"D5:D16"}</ACTION>
 */
async function createSparkline(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[createSparkline] Creating sparkline at "${target}"`);
    
    if (!target) {
        throw new Error("createSparkline requires a cell address in target");
    }
    
    try {
        // Check if sparklines API is available
        const supported = await isSparklineSupported(ctx, sheet);
        if (!supported) {
            logDiag(`[createSparkline] Sparklines require ExcelApi 1.10+. Consider using data bars for inline visualization.`);
            throw new Error("Sparklines are not available in this Excel version (requires ExcelApi 1.10+). Consider using data bars as an alternative.");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            logDiag(`[createSparkline] Warning: Sheet is protected, sparkline may not be added`);
        }
        
        // Parse options
        let options = {};
        if (data) {
            if (typeof data === "string") {
                try {
                    options = JSON.parse(data);
                } catch (e) {
                    logDiag(`[createSparkline] Warning: Could not parse data JSON: ${e.message}`);
                }
            } else if (typeof data === "object") {
                options = data;
            }
        }
        
        const sparklineType = options.type || "Line";
        const sourceData = options.sourceData;
        
        if (!sourceData) {
            throw new Error("createSparkline requires sourceData in data (e.g., 'B2:F2')");
        }
        
        // Validate sparkline type
        const typeMap = {
            "Line": Excel.SparklineType.line,
            "Column": Excel.SparklineType.column,
            "WinLoss": Excel.SparklineType.winLoss
        };
        
        const excelType = typeMap[sparklineType];
        if (!excelType) {
            throw new Error(`Invalid sparkline type "${sparklineType}". Valid types: Line, Column, WinLoss`);
        }
        
        // Validate source data range format
        const rangePattern = /^[A-Z]+\d+(:[A-Z]+\d+)?$/i;
        if (!rangePattern.test(sourceData)) {
            throw new Error(`Invalid sourceData range format "${sourceData}". Expected format: B2:F2 or C3:C20`);
        }
        
        // Create the sparkline group
        const sparklineGroup = sheet.sparklineGroups.add(excelType, sourceData, target);
        
        // Apply optional configurations
        if (options.axes) {
            if (options.axes.horizontal !== undefined) {
                sparklineGroup.axes.horizontal.axis.visible = options.axes.horizontal;
            }
            // Note: min/max axis settings require additional API support
        }
        
        // Apply marker settings (Line sparklines only)
        if (options.markers && sparklineType === "Line") {
            const points = sparklineGroup.points;
            if (options.markers.high !== undefined) {
                points.highPoint.visible = options.markers.high;
            }
            if (options.markers.low !== undefined) {
                points.lowPoint.visible = options.markers.low;
            }
            if (options.markers.first !== undefined) {
                points.firstPoint.visible = options.markers.first;
            }
            if (options.markers.last !== undefined) {
                points.lastPoint.visible = options.markers.last;
            }
            if (options.markers.negative !== undefined) {
                points.negativePoints.visible = options.markers.negative;
            }
        }
        
        // Apply color settings
        if (options.colors) {
            // Validate hex colors
            const hexPattern = /^#[0-9A-Fa-f]{6}$/;
            
            if (options.colors.series && hexPattern.test(options.colors.series)) {
                sparklineGroup.seriesColor = options.colors.series;
            }
            if (options.colors.negative && hexPattern.test(options.colors.negative)) {
                sparklineGroup.negativePointsColor = options.colors.negative;
            }
            if (options.colors.high && hexPattern.test(options.colors.high)) {
                sparklineGroup.highPointColor = options.colors.high;
            }
            if (options.colors.low && hexPattern.test(options.colors.low)) {
                sparklineGroup.lowPointColor = options.colors.low;
            }
            if (options.colors.first && hexPattern.test(options.colors.first)) {
                sparklineGroup.firstPointColor = options.colors.first;
            }
            if (options.colors.last && hexPattern.test(options.colors.last)) {
                sparklineGroup.lastPointColor = options.colors.last;
            }
        }
        
        await ctx.sync();
        
        logDiag(`[createSparkline] Successfully created ${sparklineType} sparkline at "${target}" from "${sourceData}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[createSparkline] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Normalizes a cell address for comparison by uppercasing, removing $ signs,
 * and extracting only the address part (ignoring sheet name if target has none)
 * @param {string} address - Cell address to normalize
 * @param {boolean} hasSheetInTarget - Whether the target has a sheet name
 * @returns {string} Normalized address
 */
function normalizeSparklineAddress(address, hasSheetInTarget) {
    if (!address) return "";
    let normalized = address.toUpperCase().replace(/\$/g, "");
    // If target has no sheet name, strip sheet name from address for comparison
    if (!hasSheetInTarget && normalized.includes("!")) {
        normalized = normalized.split("!")[1] || normalized;
    }
    return normalized;
}

/**
 * Configures an existing sparkline's properties
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 * 
 * Data JSON Schema:
 * {
 *   axes?: { horizontal: boolean },
 *   markers?: { high: boolean, low: boolean, first: boolean, last: boolean, negative: boolean },
 *   colors?: { series: "#hex", negative: "#hex", high: "#hex", low: "#hex" }
 * }
 */
async function configureSparkline(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[configureSparkline] Configuring sparkline at "${target}"`);
    
    if (!target) {
        throw new Error("configureSparkline requires a cell address in target");
    }
    
    try {
        // Check if sparklines API is available
        const supported = await isSparklineSupported(ctx, sheet);
        if (!supported) {
            throw new Error("Sparklines are not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Parse options
        let options = {};
        if (data) {
            if (typeof data === "string") {
                try {
                    options = JSON.parse(data);
                } catch (e) {
                    logDiag(`[configureSparkline] Warning: Could not parse data JSON: ${e.message}`);
                }
            } else if (typeof data === "object") {
                options = data;
            }
        }
        
        // Normalize target address for comparison
        const hasSheetInTarget = target.includes("!");
        const normalizedTarget = normalizeSparklineAddress(target, hasSheetInTarget);
        
        // Load sparkline groups and batch load all sparkline locations
        sheet.sparklineGroups.load("items");
        await ctx.sync();
        
        // Batch load all sparkline locations before iterating
        for (const group of sheet.sparklineGroups.items) {
            group.load("sparklines/items/location");
        }
        await ctx.sync();
        
        // Batch load all location addresses
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                sparkline.location.load("address");
            }
        }
        await ctx.sync();
        
        // Find sparkline group at the target location using strict equality
        let foundSparkline = null;
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                const normalizedAddress = normalizeSparklineAddress(sparkline.location.address, hasSheetInTarget);
                if (normalizedAddress === normalizedTarget) {
                    foundSparkline = group;
                    break;
                }
            }
            if (foundSparkline) break;
        }
        
        if (!foundSparkline) {
            logDiag(`[configureSparkline] No sparkline found at "${target}"`);
            throw new Error(`No sparkline found at cell "${target}"`);
        }
        
        // Apply configurations
        const hexPattern = /^#[0-9A-Fa-f]{6}$/;
        
        if (options.colors) {
            if (options.colors.series && hexPattern.test(options.colors.series)) {
                foundSparkline.seriesColor = options.colors.series;
            }
            if (options.colors.negative && hexPattern.test(options.colors.negative)) {
                foundSparkline.negativePointsColor = options.colors.negative;
            }
            if (options.colors.high && hexPattern.test(options.colors.high)) {
                foundSparkline.highPointColor = options.colors.high;
            }
            if (options.colors.low && hexPattern.test(options.colors.low)) {
                foundSparkline.lowPointColor = options.colors.low;
            }
        }
        
        if (options.markers) {
            const points = foundSparkline.points;
            if (options.markers.high !== undefined) {
                points.highPoint.visible = options.markers.high;
            }
            if (options.markers.low !== undefined) {
                points.lowPoint.visible = options.markers.low;
            }
            if (options.markers.first !== undefined) {
                points.firstPoint.visible = options.markers.first;
            }
            if (options.markers.last !== undefined) {
                points.lastPoint.visible = options.markers.last;
            }
            if (options.markers.negative !== undefined) {
                points.negativePoints.visible = options.markers.negative;
            }
        }
        
        if (options.axes && options.axes.horizontal !== undefined) {
            foundSparkline.axes.horizontal.axis.visible = options.axes.horizontal;
        }
        
        await ctx.sync();
        
        logDiag(`[configureSparkline] Successfully configured sparkline at "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[configureSparkline] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Deletes sparkline(s) at the specified location
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target
 */
async function deleteSparkline(ctx, sheet, action) {
    const { target } = action;
    logDiag(`[deleteSparkline] Deleting sparkline at "${target}"`);
    
    if (!target) {
        throw new Error("deleteSparkline requires a cell address in target");
    }
    
    try {
        // Check if sparklines API is available
        const supported = await isSparklineSupported(ctx, sheet);
        if (!supported) {
            throw new Error("Sparklines are not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Normalize target address for comparison
        const hasSheetInTarget = target.includes("!");
        const normalizedTarget = normalizeSparklineAddress(target, hasSheetInTarget);
        
        // Load sparkline groups and batch load all sparkline locations
        sheet.sparklineGroups.load("items");
        await ctx.sync();
        
        // Batch load all sparkline locations before iterating
        for (const group of sheet.sparklineGroups.items) {
            group.load("sparklines/items/location");
        }
        await ctx.sync();
        
        // Batch load all location addresses
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                sparkline.location.load("address");
            }
        }
        await ctx.sync();
        
        // Find sparkline groups at the target location using strict equality
        let deletedCount = 0;
        const groupsToDelete = [];
        
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                const normalizedAddress = normalizeSparklineAddress(sparkline.location.address, hasSheetInTarget);
                if (normalizedAddress === normalizedTarget) {
                    groupsToDelete.push(group);
                    break;
                }
            }
        }
        
        // Delete the found groups
        for (const group of groupsToDelete) {
            group.delete();
            deletedCount++;
        }
        
        await ctx.sync();
        
        if (deletedCount === 0) {
            logDiag(`[deleteSparkline] No sparkline found at "${target}" - nothing to delete`);
        } else {
            logDiag(`[deleteSparkline] Successfully deleted ${deletedCount} sparkline group(s) at "${target}"`);
        }
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[deleteSparkline] Error: ${errorMsg}`);
        throw error;
    }
}

// ============================================================================
// Worksheet Management Operations
// ============================================================================

/**
 * Renames a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet (may not be target)
 * @param {Object} action - Action object with target (old name) and data (newName)
 */
async function renameSheet(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[renameSheet] Renaming sheet "${target}"`);
    
    if (!target) {
        throw new Error("renameSheet requires a sheet name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const newName = options.newName;
        
        if (!newName) {
            throw new Error("renameSheet requires newName in data");
        }
        
        // Validate new name
        if (newName.length > 31) {
            throw new Error("Sheet name cannot exceed 31 characters");
        }
        
        const invalidChars = /[\\\/\?\*\[\]]/;
        if (invalidChars.test(newName)) {
            throw new Error("Sheet name cannot contain \\ / ? * [ ] characters");
        }
        
        if (newName.trim() === "") {
            throw new Error("Sheet name cannot be empty");
        }
        
        // Get the target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load("name");
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        // Check for duplicate name
        const existingSheet = ctx.workbook.worksheets.getItemOrNullObject(newName);
        existingSheet.load("name");
        await ctx.sync();
        
        if (!existingSheet.isNullObject && existingSheet.name.toLowerCase() !== target.toLowerCase()) {
            throw new Error(`A sheet named "${newName}" already exists`);
        }
        
        // Rename the sheet
        targetSheet.name = newName;
        await ctx.sync();
        
        logDiag(`[renameSheet] Successfully renamed sheet "${target}" to "${newName}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[renameSheet] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Moves a worksheet to a new position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target and data (position, referenceSheet)
 */
async function moveSheet(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[moveSheet] Moving sheet "${target}"`);
    
    if (!target) {
        throw new Error("moveSheet requires a sheet name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const position = options.position || "last";
        const referenceSheet = options.referenceSheet;
        
        // Get the target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load("name");
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        // Get all sheets to determine positions
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let newPosition;
        
        switch (position.toLowerCase()) {
            case "first":
                newPosition = 0;
                break;
            case "last":
                newPosition = sheets.items.length - 1;
                break;
            case "before":
            case "after":
                if (!referenceSheet) {
                    throw new Error(`moveSheet with position "${position}" requires referenceSheet`);
                }
                const refSheet = ctx.workbook.worksheets.getItemOrNullObject(referenceSheet);
                refSheet.load("position");
                await ctx.sync();
                
                if (refSheet.isNullObject) {
                    throw new Error(`Reference sheet "${referenceSheet}" not found`);
                }
                
                newPosition = position.toLowerCase() === "before" 
                    ? refSheet.position 
                    : refSheet.position + 1;
                break;
            default:
                throw new Error(`Invalid position "${position}". Use: first, last, before, after`);
        }
        
        // Move the sheet
        targetSheet.position = newPosition;
        await ctx.sync();
        
        logDiag(`[moveSheet] Successfully moved sheet "${target}" to position ${newPosition}`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[moveSheet] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Hides a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name)
 */
async function hideSheet(ctx, sheet, action) {
    const { target } = action;
    logDiag(`[hideSheet] Hiding sheet "${target}"`);
    
    if (!target) {
        throw new Error("hideSheet requires a sheet name in target");
    }
    
    try {
        // First, get the target sheet and check its current visibility
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load(["name", "visibility"]);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        // If already hidden, return early without checking visible sheet count
        if (targetSheet.visibility !== Excel.SheetVisibility.visible) {
            logDiag(`[hideSheet] Sheet "${target}" is already hidden`);
            return;
        }
        
        // Only check visible sheet count if we're about to hide a visible sheet
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/visibility");
        await ctx.sync();
        
        const visibleSheets = sheets.items.filter(s => s.visibility === Excel.SheetVisibility.visible);
        
        if (visibleSheets.length <= 1) {
            throw new Error("Cannot hide the only visible sheet");
        }
        
        // Hide the sheet
        targetSheet.visibility = Excel.SheetVisibility.hidden;
        await ctx.sync();
        
        logDiag(`[hideSheet] Successfully hidden sheet "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[hideSheet] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Unhides a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name)
 */
async function unhideSheet(ctx, sheet, action) {
    const { target } = action;
    logDiag(`[unhideSheet] Unhiding sheet "${target}"`);
    
    if (!target) {
        throw new Error("unhideSheet requires a sheet name in target");
    }
    
    try {
        // Get the target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load(["name", "visibility"]);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        if (targetSheet.visibility === Excel.SheetVisibility.visible) {
            logDiag(`[unhideSheet] Sheet "${target}" is already visible`);
            return;
        }
        
        // Unhide the sheet
        targetSheet.visibility = Excel.SheetVisibility.visible;
        await ctx.sync();
        
        logDiag(`[unhideSheet] Successfully unhidden sheet "${target}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[unhideSheet] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Freezes panes at a specified cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet to freeze
 * @param {Object} action - Action object with target (cell) and data (freezeType)
 */
async function freezePanes(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[freezePanes] Freezing panes at "${target}"`);
    
    if (!target) {
        throw new Error("freezePanes requires a cell address in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const freezeType = options.freezeType || "both";
        
        // Get the freeze range
        const range = sheet.getRange(target);
        range.load(["rowIndex", "columnIndex"]);
        await ctx.sync();
        
        const rowCount = range.rowIndex;
        const colCount = range.columnIndex;
        
        // Apply freeze based on type
        switch (freezeType.toLowerCase()) {
            case "rows":
                if (rowCount > 0) {
                    sheet.freezePanes.freezeRows(rowCount);
                } else {
                    throw new Error("Cannot freeze rows: target cell is in row 1");
                }
                break;
            case "columns":
                if (colCount > 0) {
                    sheet.freezePanes.freezeColumns(colCount);
                } else {
                    throw new Error("Cannot freeze columns: target cell is in column A");
                }
                break;
            case "both":
            default:
                sheet.freezePanes.freezeAt(range);
                break;
        }
        
        await ctx.sync();
        
        logDiag(`[freezePanes] Successfully froze panes at "${target}" (type: ${freezeType})`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[freezePanes] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Unfreezes all panes on a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target ("current" or sheet name)
 */
async function unfreezePane(ctx, sheet, action) {
    const { target } = action;
    logDiag(`[unfreezePane] Unfreezing panes on "${target}"`);
    
    try {
        let targetSheet = sheet;
        
        if (target && target.toLowerCase() !== "current") {
            targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
            targetSheet.load("name");
            await ctx.sync();
            
            if (targetSheet.isNullObject) {
                throw new Error(`Sheet "${target}" not found`);
            }
        }
        
        // Unfreeze all panes
        targetSheet.freezePanes.unfreeze();
        await ctx.sync();
        
        logDiag(`[unfreezePane] Successfully unfroze panes`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[unfreezePane] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Sets the zoom level for a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target and data (zoomLevel)
 */
async function setZoom(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[setZoom] Setting zoom on "${target}"`);
    
    try {
        const options = data ? JSON.parse(data) : {};
        const zoomLevel = options.zoomLevel;
        
        if (zoomLevel === undefined || zoomLevel === null) {
            throw new Error("setZoom requires zoomLevel in data");
        }
        
        const zoom = parseInt(zoomLevel);
        if (isNaN(zoom) || zoom < 10 || zoom > 400) {
            throw new Error("zoomLevel must be between 10 and 400");
        }
        
        let targetSheet = sheet;
        
        if (target && target.toLowerCase() !== "current") {
            targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
            targetSheet.load("name");
            await ctx.sync();
            
            if (targetSheet.isNullObject) {
                throw new Error(`Sheet "${target}" not found`);
            }
        }
        
        // Set zoom level via the worksheet's pageLayout (Office.js workaround)
        // Note: Direct zoom setting may require specific API version
        try {
            // Try using the view API if available
            if (targetSheet.view) {
                targetSheet.view.zoom = zoom;
            } else {
                // Fallback: use pageLayout zoom
                targetSheet.pageLayout.zoom = { scale: zoom };
            }
            await ctx.sync();
        } catch (zoomError) {
            // If view API not available, try pageLayout
            targetSheet.pageLayout.zoom = { scale: zoom };
            await ctx.sync();
        }
        
        logDiag(`[setZoom] Successfully set zoom to ${zoom}%`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[setZoom] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Splits panes at a specified cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet to split
 * @param {Object} action - Action object with target (cell) and data (horizontal, vertical)
 */
async function splitPane(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[splitPane] Splitting panes at "${target}"`);
    
    if (!target) {
        throw new Error("splitPane requires a cell address in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const horizontal = options.horizontal !== false;
        const vertical = options.vertical !== false;
        
        // Get the split position
        const range = sheet.getRange(target);
        range.load(["rowIndex", "columnIndex"]);
        await ctx.sync();
        
        // Note: Office.js has limited support for split panes
        // Using freezeAt as a workaround which creates a similar effect
        if (horizontal && vertical) {
            sheet.freezePanes.freezeAt(range);
        } else if (horizontal && !vertical) {
            // Horizontal-only split: guard against row 1
            if (range.rowIndex === 0) {
                throw new Error("Cannot split horizontally at row 1; choose a cell below the first row");
            }
            sheet.freezePanes.freezeRows(range.rowIndex);
        } else if (vertical && !horizontal) {
            // Vertical-only split: guard against column A
            if (range.columnIndex === 0) {
                throw new Error("Cannot split vertically at column A; choose a cell to the right of the first column");
            }
            sheet.freezePanes.freezeColumns(range.columnIndex);
        }
        
        await ctx.sync();
        
        logDiag(`[splitPane] Successfully split panes at "${target}" (H:${horizontal}, V:${vertical})`);
        logDiag(`[splitPane] Note: Using freeze panes as Office.js split pane API is limited`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[splitPane] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Creates a custom view (limited API support)
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (view name) and data (options)
 */
async function createView(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[createView] Creating view "${target}"`);
    
    if (!target) {
        throw new Error("createView requires a view name in target");
    }
    
    try {
        // Note: Office.js has very limited support for custom views
        // This is a placeholder that logs the current state
        const options = data ? JSON.parse(data) : {};
        
        // Load current view state for documentation
        sheet.load("name");
        await ctx.sync();
        
        logDiag(`[createView] Custom view "${target}" requested for sheet "${sheet.name}"`);
        logDiag(`[createView] Options: includeHidden=${options.includeHidden}, includePrint=${options.includePrint}, includeFilter=${options.includeFilter}`);
        logDiag(`[createView] Note: Office.js has limited custom view API support. Use Excel UI: View > Custom Views > Add`);
        
        // Return without error but with warning
        console.warn(`Custom view "${target}" creation requires manual Excel UI. Go to View > Custom Views > Add.`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[createView] Error: ${errorMsg}`);
        throw error;
    }
}

// ============================================================================
// Page Setup Functions
// ============================================================================

/**
 * Converts column number to Excel column letter (1 → A, 27 → AA)
 * @param {number} colNum - Column number (1-based)
 * @returns {string} Column letter
 */
function columnNumberToLetter(colNum) {
    let letter = '';
    while (colNum > 0) {
        const remainder = (colNum - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        colNum = Math.floor((colNum - 1) / 26);
    }
    return letter;
}

/**
 * Sets comprehensive page setup for printing
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name) and data (orientation, paperSize, scaling, etc.)
 * @returns {Promise<void>}
 */
async function setPageSetup(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[setPageSetup] Setting page setup for "${target || 'current sheet'}"`);
    
    try {
        // Parse options from data
        const options = data ? JSON.parse(data) : {};
        
        // Resolve target sheet
        let targetSheet = sheet;
        if (target && target !== "current") {
            const sheetObj = ctx.workbook.worksheets.getItemOrNullObject(target);
            await ctx.sync();
            if (!sheetObj.isNullObject) {
                targetSheet = sheetObj;
            } else {
                logDiag(`[setPageSetup] Sheet "${target}" not found, using current sheet`);
            }
        }
        
        const pageLayout = targetSheet.pageLayout;
        const appliedSettings = [];
        
        // Set orientation
        if (options.orientation) {
            const orientation = options.orientation.toLowerCase();
            if (orientation === "portrait") {
                pageLayout.orientation = Excel.PageOrientation.portrait;
                appliedSettings.push("orientation: portrait");
            } else if (orientation === "landscape") {
                pageLayout.orientation = Excel.PageOrientation.landscape;
                appliedSettings.push("orientation: landscape");
            } else {
                logDiag(`[setPageSetup] Invalid orientation: ${options.orientation}`);
            }
        }
        
        // Set paper size
        if (options.paperSize) {
            const paperSizeMap = {
                "letter": Excel.PaperType.letter,
                "a4": Excel.PaperType.a4,
                "legal": Excel.PaperType.legal,
                "tabloid": Excel.PaperType.tabloid,
                "a3": Excel.PaperType.a3,
                "a5": Excel.PaperType.a5
            };
            const paperType = paperSizeMap[options.paperSize.toLowerCase()];
            if (paperType !== undefined) {
                pageLayout.paperSize = paperType;
                appliedSettings.push(`paperSize: ${options.paperSize}`);
            } else {
                logDiag(`[setPageSetup] Invalid paper size: ${options.paperSize}`);
            }
        }
        
        // Set scaling or fit-to-pages (mutually exclusive)
        // fitToPages takes precedence over scaling
        if (options.fitToPages) {
            const width = parseInt(options.fitToPages.width) || 1;
            const height = parseInt(options.fitToPages.height) || 1;
            // Validate positive integers
            if (width > 0 && height > 0) {
                pageLayout.zoom = { 
                    horizontalFitToPages: width, 
                    verticalFitToPages: height 
                };
                appliedSettings.push(`fitToPages: ${width}×${height}`);
            } else {
                logDiag(`[setPageSetup] Invalid fitToPages: width=${width}, height=${height} (must be positive integers)`);
            }
        } else if (options.scaling !== undefined) {
            const scale = parseInt(options.scaling);
            if (scale >= 10 && scale <= 400) {
                pageLayout.zoom = { scale: scale };
                appliedSettings.push(`scaling: ${scale}%`);
            } else {
                logDiag(`[setPageSetup] Invalid scaling: ${options.scaling} (must be 10-400)`);
            }
        }
        // Note: If neither fitToPages nor scaling is provided, existing zoom settings are retained
        
        // Set print gridlines
        if (options.printGridlines !== undefined) {
            pageLayout.printGridlines = options.printGridlines;
            appliedSettings.push(`printGridlines: ${options.printGridlines}`);
        }
        
        // Set print headings (row/column headers)
        if (options.printHeadings !== undefined) {
            pageLayout.printHeadings = options.printHeadings;
            appliedSettings.push(`printHeadings: ${options.printHeadings}`);
        }
        
        await ctx.sync();
        
        logDiag(`[setPageSetup] Applied settings: ${appliedSettings.join(", ") || "none"}`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[setPageSetup] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Configures page margins for printing
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name) and data (margins in inches)
 * @returns {Promise<void>}
 */
async function setPageMargins(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[setPageMargins] Setting page margins for "${target || 'current sheet'}"`);
    
    try {
        // Parse options from data
        const options = data ? JSON.parse(data) : {};
        
        // Resolve target sheet
        let targetSheet = sheet;
        if (target && target !== "current") {
            const sheetObj = ctx.workbook.worksheets.getItemOrNullObject(target);
            await ctx.sync();
            if (!sheetObj.isNullObject) {
                targetSheet = sheetObj;
            } else {
                logDiag(`[setPageMargins] Sheet "${target}" not found, using current sheet`);
            }
        }
        
        const pageLayout = targetSheet.pageLayout;
        const appliedMargins = [];
        
        // Convert inches to points (1 inch = 72 points)
        const inchesToPoints = (inches) => inches * 72;
        
        // Validate margin value
        const validateMargin = (value, name) => {
            if (value < 0) {
                logDiag(`[setPageMargins] Invalid ${name}: ${value} (must be >= 0)`);
                return false;
            }
            if (value > 5) {
                logDiag(`[setPageMargins] Warning: ${name} ${value}" is unusually large`);
            }
            return true;
        };
        
        // Set top margin
        if (options.top !== undefined && validateMargin(options.top, "top")) {
            pageLayout.topMargin = inchesToPoints(options.top);
            appliedMargins.push(`top: ${options.top}"`);
        }
        
        // Set bottom margin
        if (options.bottom !== undefined && validateMargin(options.bottom, "bottom")) {
            pageLayout.bottomMargin = inchesToPoints(options.bottom);
            appliedMargins.push(`bottom: ${options.bottom}"`);
        }
        
        // Set left margin
        if (options.left !== undefined && validateMargin(options.left, "left")) {
            pageLayout.leftMargin = inchesToPoints(options.left);
            appliedMargins.push(`left: ${options.left}"`);
        }
        
        // Set right margin
        if (options.right !== undefined && validateMargin(options.right, "right")) {
            pageLayout.rightMargin = inchesToPoints(options.right);
            appliedMargins.push(`right: ${options.right}"`);
        }
        
        // Set header margin
        if (options.header !== undefined && validateMargin(options.header, "header")) {
            pageLayout.headerMargin = inchesToPoints(options.header);
            appliedMargins.push(`header: ${options.header}"`);
        }
        
        // Set footer margin
        if (options.footer !== undefined && validateMargin(options.footer, "footer")) {
            pageLayout.footerMargin = inchesToPoints(options.footer);
            appliedMargins.push(`footer: ${options.footer}"`);
        }
        
        await ctx.sync();
        
        logDiag(`[setPageMargins] Applied margins: ${appliedMargins.join(", ") || "none"}`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[setPageMargins] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Sets page orientation (portrait/landscape)
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name) and data (orientation)
 * @returns {Promise<void>}
 */
async function setPageOrientation(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[setPageOrientation] Setting orientation for "${target || 'current sheet'}"`);
    
    try {
        // Parse options from data or action attributes
        let orientation = null;
        if (data) {
            try {
                const options = JSON.parse(data);
                orientation = options.orientation;
            } catch {
                // data might be the orientation string directly
                orientation = data;
            }
        }
        orientation = orientation || action.orientation || "portrait";
        
        // Resolve target sheet
        let targetSheet = sheet;
        if (target && target !== "current") {
            const sheetObj = ctx.workbook.worksheets.getItemOrNullObject(target);
            await ctx.sync();
            if (!sheetObj.isNullObject) {
                targetSheet = sheetObj;
            } else {
                logDiag(`[setPageOrientation] Sheet "${target}" not found, using current sheet`);
            }
        }
        
        // Set orientation
        const orientationLower = orientation.toLowerCase();
        if (orientationLower === "portrait") {
            targetSheet.pageLayout.orientation = Excel.PageOrientation.portrait;
            logDiag(`[setPageOrientation] Set orientation to portrait`);
        } else if (orientationLower === "landscape") {
            targetSheet.pageLayout.orientation = Excel.PageOrientation.landscape;
            logDiag(`[setPageOrientation] Set orientation to landscape`);
        } else {
            throw new Error(`Invalid orientation: ${orientation}. Use "portrait" or "landscape".`);
        }
        
        await ctx.sync();
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[setPageOrientation] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Defines the print area for a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (range address or "clear")
 * @returns {Promise<void>}
 */
async function setPrintArea(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[setPrintArea] Setting print area: "${target}"`);
    
    try {
        // Parse options from data if present
        let sheetName = null;
        let rangeAddress = target;
        
        if (data) {
            try {
                const options = JSON.parse(data);
                sheetName = options.sheet;
                if (options.range) {
                    rangeAddress = options.range;
                }
            } catch {
                // data is not JSON, ignore
            }
        }
        
        // Resolve target sheet
        let targetSheet = sheet;
        if (sheetName) {
            const sheetObj = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
            await ctx.sync();
            if (!sheetObj.isNullObject) {
                targetSheet = sheetObj;
            } else {
                logDiag(`[setPrintArea] Sheet "${sheetName}" not found, using current sheet`);
            }
        }
        
        // Handle clear print area
        if (!rangeAddress || rangeAddress.toLowerCase() === "clear") {
            targetSheet.pageLayout.setPrintArea("");
            await ctx.sync();
            logDiag(`[setPrintArea] Cleared print area`);
            return;
        }
        
        // Set print area
        targetSheet.pageLayout.setPrintArea(rangeAddress);
        await ctx.sync();
        
        logDiag(`[setPrintArea] Set print area to "${rangeAddress}"`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[setPrintArea] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Sets headers and footers with dynamic fields
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name) and data (header/footer content)
 * @returns {Promise<void>}
 */
async function setHeaderFooter(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[setHeaderFooter] Setting headers/footers for "${target || 'current sheet'}"`);
    
    try {
        // Parse options from data
        const options = data ? JSON.parse(data) : {};
        
        // Resolve target sheet
        let targetSheet = sheet;
        if (target && target !== "current") {
            const sheetObj = ctx.workbook.worksheets.getItemOrNullObject(target);
            await ctx.sync();
            if (!sheetObj.isNullObject) {
                targetSheet = sheetObj;
            } else {
                logDiag(`[setHeaderFooter] Sheet "${target}" not found, using current sheet`);
            }
        }
        
        // Get headers/footers group based on page type
        const headersFooters = targetSheet.pageLayout.headersFooters;
        let headerFooter;
        
        const pageType = (options.pageType || "default").toLowerCase();
        switch (pageType) {
            case "first":
                headerFooter = headersFooters.firstPage;
                break;
            case "even":
                headerFooter = headersFooters.evenPages;
                break;
            case "odd":
                headerFooter = headersFooters.oddPages;
                break;
            default:
                headerFooter = headersFooters.defaultForAllPages;
        }
        
        const appliedSettings = [];
        
        // Set headers
        if (options.leftHeader !== undefined) {
            headerFooter.leftHeader = options.leftHeader;
            appliedSettings.push("leftHeader");
        }
        if (options.centerHeader !== undefined) {
            headerFooter.centerHeader = options.centerHeader;
            appliedSettings.push("centerHeader");
        }
        if (options.rightHeader !== undefined) {
            headerFooter.rightHeader = options.rightHeader;
            appliedSettings.push("rightHeader");
        }
        
        // Set footers
        if (options.leftFooter !== undefined) {
            headerFooter.leftFooter = options.leftFooter;
            appliedSettings.push("leftFooter");
        }
        if (options.centerFooter !== undefined) {
            headerFooter.centerFooter = options.centerFooter;
            appliedSettings.push("centerFooter");
        }
        if (options.rightFooter !== undefined) {
            headerFooter.rightFooter = options.rightFooter;
            appliedSettings.push("rightFooter");
        }
        
        await ctx.sync();
        
        logDiag(`[setHeaderFooter] Applied: ${appliedSettings.join(", ") || "none"} (pageType: ${pageType})`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        // Check for API version issues
        if (errorMsg.includes("headersFooters") || errorMsg.includes("not supported")) {
            logDiag(`[setHeaderFooter] Headers/footers require ExcelApi 1.9+ (Excel 2019/365/Online)`);
            console.warn("Headers/footers require Excel 2019 or later. Please update Excel or use the manual Page Layout menu.");
        }
        logDiag(`[setHeaderFooter] Error: ${errorMsg}`);
        throw error;
    }
}

/**
 * Manages manual page breaks (add/remove/clear)
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name) and data (breaks array, action type)
 * @returns {Promise<void>}
 */
async function setPageBreaks(ctx, sheet, action) {
    const { target, data } = action;
    logDiag(`[setPageBreaks] Managing page breaks for "${target || 'current sheet'}"`);
    
    try {
        // Parse options from data
        const options = data ? JSON.parse(data) : {};
        const breakAction = (options.action || "add").toLowerCase();
        const breaks = options.breaks || [];
        
        // Resolve target sheet
        let targetSheet = sheet;
        if (target && target !== "current") {
            const sheetObj = ctx.workbook.worksheets.getItemOrNullObject(target);
            await ctx.sync();
            if (!sheetObj.isNullObject) {
                targetSheet = sheetObj;
            } else {
                logDiag(`[setPageBreaks] Sheet "${target}" not found, using current sheet`);
            }
        }
        
        const hBreaks = targetSheet.horizontalPageBreaks;
        const vBreaks = targetSheet.verticalPageBreaks;
        
        if (breakAction === "clear") {
            // Clear all page breaks
            hBreaks.load("items");
            vBreaks.load("items");
            await ctx.sync();
            
            // Delete all horizontal breaks
            const hItems = hBreaks.items;
            for (let i = hItems.length - 1; i >= 0; i--) {
                hItems[i].delete();
            }
            
            // Delete all vertical breaks
            const vItems = vBreaks.items;
            for (let i = vItems.length - 1; i >= 0; i--) {
                vItems[i].delete();
            }
            
            await ctx.sync();
            logDiag(`[setPageBreaks] Cleared all page breaks`);
            return;
        }
        
        if (breakAction === "add") {
            let addedCount = 0;
            
            for (const brk of breaks) {
                const breakType = (brk.type || "horizontal").toLowerCase();
                
                if (breakType === "horizontal" && brk.row) {
                    // Add horizontal page break (break above the specified row)
                    const row = parseInt(brk.row);
                    if (row > 1) {
                        const breakRange = targetSheet.getRange(`A${row}`);
                        hBreaks.add(breakRange);
                        addedCount++;
                        logDiag(`[setPageBreaks] Added horizontal break above row ${row}`);
                    } else {
                        logDiag(`[setPageBreaks] Cannot add horizontal break at row ${row} (must be > 1)`);
                    }
                } else if (breakType === "vertical" && brk.col) {
                    // Add vertical page break (break left of the specified column)
                    const col = parseInt(brk.col);
                    if (col > 1) {
                        const colLetter = columnNumberToLetter(col);
                        const breakRange = targetSheet.getRange(`${colLetter}1`);
                        vBreaks.add(breakRange);
                        addedCount++;
                        logDiag(`[setPageBreaks] Added vertical break left of column ${colLetter} (${col})`);
                    } else {
                        logDiag(`[setPageBreaks] Cannot add vertical break at column ${col} (must be > 1)`);
                    }
                }
            }
            
            await ctx.sync();
            logDiag(`[setPageBreaks] Added ${addedCount} page break(s)`);
            return;
        }
        
        if (breakAction === "remove") {
            // Load existing breaks and their indices once before processing
            hBreaks.load("items");
            vBreaks.load("items");
            await ctx.sync();
            
            // Load rowIndex for all horizontal breaks
            for (const hBreak of hBreaks.items) {
                hBreak.load("rowIndex");
            }
            // Load columnIndex for all vertical breaks
            for (const vBreak of vBreaks.items) {
                vBreak.load("columnIndex");
            }
            await ctx.sync();
            
            let removedCount = 0;
            
            // Process all breaks without additional sync calls
            for (const brk of breaks) {
                const breakType = (brk.type || "horizontal").toLowerCase();
                
                if (breakType === "horizontal" && brk.row) {
                    const row = parseInt(brk.row);
                    if (row <= 1) {
                        logDiag(`[setPageBreaks] Invalid row ${row} for removal (must be > 1)`);
                        continue;
                    }
                    // Find and delete matching horizontal break
                    for (const hBreak of hBreaks.items) {
                        if (hBreak.rowIndex === row - 1) { // rowIndex is 0-based
                            hBreak.delete();
                            removedCount++;
                            logDiag(`[setPageBreaks] Removed horizontal break at row ${row}`);
                            break;
                        }
                    }
                } else if (breakType === "vertical" && brk.col) {
                    const col = parseInt(brk.col);
                    if (col <= 1) {
                        logDiag(`[setPageBreaks] Invalid column ${col} for removal (must be > 1)`);
                        continue;
                    }
                    // Find and delete matching vertical break
                    for (const vBreak of vBreaks.items) {
                        if (vBreak.columnIndex === col - 1) { // columnIndex is 0-based
                            vBreak.delete();
                            removedCount++;
                            logDiag(`[setPageBreaks] Removed vertical break at column ${col}`);
                            break;
                        }
                    }
                }
            }
            
            await ctx.sync();
            logDiag(`[setPageBreaks] Removed ${removedCount} page break(s)`);
            return;
        }
        
        logDiag(`[setPageBreaks] Unknown action: ${breakAction}. Use "add", "remove", or "clear".`);
    } catch (error) {
        const errorMsg = error.message || String(error);
        logDiag(`[setPageBreaks] Error: ${errorMsg}`);
        throw error;
    }
}

// ============================================================================
// Hyperlink Operations
// ============================================================================

// Cache for hyperlink API support check
let hyperlinkSupportChecked = false;
let hyperlinkSupported = false;

/**
 * Checks if the Range.hyperlink API is supported (ExcelApi 1.7+)
 * @param {Excel.RequestContext} ctx - Excel context
 * @returns {Promise<boolean>} True if hyperlinks are supported
 */
async function isHyperlinkSupported(ctx) {
    if (hyperlinkSupportChecked) {
        return hyperlinkSupported;
    }
    
    try {
        // Check using Office.context.requirements if available
        if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
            hyperlinkSupported = Office.context.requirements.isSetSupported('ExcelApi', '1.7');
            hyperlinkSupportChecked = true;
            return hyperlinkSupported;
        }
        
        // Fallback: try a lightweight operation to test support
        const testRange = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1");
        testRange.load("hyperlink");
        await ctx.sync();
        hyperlinkSupported = true;
        hyperlinkSupportChecked = true;
        return true;
    } catch (e) {
        hyperlinkSupported = false;
        hyperlinkSupportChecked = true;
        return false;
    }
}

/**
 * Adds a hyperlink to a cell or range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string with hyperlink options
 * 
 * Supported options:
 * - url: Web URL (e.g., "https://example.com")
 * - email: Email address (automatically prefixed with "mailto:")
 * - documentReference: Internal link (e.g., "'Sheet2'!A1")
 * - displayText: Text to display in cell (defaults to URL/email/reference)
 * - tooltip: Hover tooltip text (screenTip)
 * 
 * Note: Only one of url, email, or documentReference should be provided.
 * Requires ExcelApi 1.7+ (Excel 2016+, Excel Online, Excel 365)
 */
async function addHyperlink(ctx, range, data) {
    logDiag(`[addHyperlink] Starting hyperlink addition`);
    
    // Check API support
    const supported = await isHyperlinkSupported(ctx);
    if (!supported) {
        throw new Error("Hyperlinks require ExcelApi 1.7+; your version does not support this feature.");
    }
    
    let options = { url: null, email: null, documentReference: null, displayText: null, tooltip: "" };
    if (data) {
        try {
            options = { ...options, ...JSON.parse(data) };
        } catch (e) {
            logDiag(`[addHyperlink] Warning: Failed to parse data: ${e.message}`);
        }
    }
    
    // Validate: must have exactly one of url, email, or documentReference
    const linkTypes = [options.url, options.email, options.documentReference].filter(v => v);
    if (linkTypes.length === 0) {
        throw new Error("Invalid hyperlink data: must provide url, email, or documentReference");
    }
    if (linkTypes.length > 1) {
        throw new Error("Invalid hyperlink data: provide only one of url, email, or documentReference");
    }
    
    try {
        let hyperlinkObj = { screenTip: options.tooltip || "" };
        
        if (options.url) {
            // Validate URL format
            if (!options.url.match(/^https?:\/\//i) && !options.url.startsWith("//")) {
                options.url = "https://" + options.url;
            }
            hyperlinkObj.address = options.url;
            hyperlinkObj.textToDisplay = options.displayText || options.url;
            logDiag(`[addHyperlink] Adding web URL: ${options.url}`);
        } else if (options.email) {
            // Automatically add mailto: prefix
            const emailAddress = options.email.startsWith("mailto:") ? options.email : "mailto:" + options.email;
            hyperlinkObj.address = emailAddress;
            hyperlinkObj.textToDisplay = options.displayText || options.email;
            logDiag(`[addHyperlink] Adding email link: ${options.email}`);
        } else if (options.documentReference) {
            hyperlinkObj.documentReference = options.documentReference;
            hyperlinkObj.textToDisplay = options.displayText || options.documentReference;
            logDiag(`[addHyperlink] Adding internal link: ${options.documentReference}`);
        }
        
        range.hyperlink = hyperlinkObj;
        await ctx.sync();
        
        logDiag(`[addHyperlink] Successfully added hyperlink`);
    } catch (e) {
        throw new Error(`Failed to add hyperlink: ${e.message}`);
    }
}

/**
 * Removes hyperlink(s) from a cell or range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * 
 * Note: This clears only the hyperlink, preserving cell values and formatting.
 * Always clears the entire range even if only some cells have hyperlinks.
 * Requires ExcelApi 1.7+
 */
async function removeHyperlink(ctx, range) {
    logDiag(`[removeHyperlink] Starting hyperlink removal`);
    
    // Check API support
    const supported = await isHyperlinkSupported(ctx);
    if (!supported) {
        throw new Error("Hyperlinks require ExcelApi 1.7+; your version does not support this feature.");
    }
    
    try {
        // Clear hyperlinks from entire range using clear method
        // This works even if only some cells in the range have hyperlinks
        range.clear(Excel.ClearApplyTo.hyperlinks);
        await ctx.sync();
        
        logDiag(`[removeHyperlink] Successfully removed hyperlinks from range`);
    } catch (e) {
        throw new Error(`Failed to remove hyperlink: ${e.message}`);
    }
}

/**
 * Edits an existing hyperlink or adds a new one if none exists
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string with hyperlink options to update
 * 
 * Supported options (all optional - only provided fields are updated):
 * - url: New web URL
 * - email: New email address
 * - documentReference: New internal link
 * - displayText: New display text
 * - tooltip: New tooltip text
 * 
 * Note: If changing link type (e.g., url to documentReference), the old type is cleared.
 * Requires ExcelApi 1.7+
 */
async function editHyperlink(ctx, range, data) {
    logDiag(`[editHyperlink] Starting hyperlink edit`);
    
    // Check API support
    const supported = await isHyperlinkSupported(ctx);
    if (!supported) {
        throw new Error("Hyperlinks require ExcelApi 1.7+; your version does not support this feature.");
    }
    
    let options = {};
    if (data) {
        try {
            options = JSON.parse(data);
        } catch (e) {
            logDiag(`[editHyperlink] Warning: Failed to parse data: ${e.message}`);
        }
    }
    
    try {
        // Load existing hyperlink
        range.load("hyperlink");
        await ctx.sync();
        
        const existingHyperlink = range.hyperlink || {};
        let hyperlinkObj = {
            screenTip: options.tooltip !== undefined ? options.tooltip : (existingHyperlink.screenTip || ""),
            textToDisplay: options.displayText !== undefined ? options.displayText : existingHyperlink.textToDisplay
        };
        
        // Determine link type - new value takes precedence
        if (options.url) {
            if (!options.url.match(/^https?:\/\//i) && !options.url.startsWith("//")) {
                options.url = "https://" + options.url;
            }
            hyperlinkObj.address = options.url;
            if (!options.displayText && !existingHyperlink.textToDisplay) {
                hyperlinkObj.textToDisplay = options.url;
            }
            logDiag(`[editHyperlink] Updating to web URL: ${options.url}`);
        } else if (options.email) {
            const emailAddress = options.email.startsWith("mailto:") ? options.email : "mailto:" + options.email;
            hyperlinkObj.address = emailAddress;
            if (!options.displayText && !existingHyperlink.textToDisplay) {
                hyperlinkObj.textToDisplay = options.email;
            }
            logDiag(`[editHyperlink] Updating to email link: ${options.email}`);
        } else if (options.documentReference) {
            hyperlinkObj.documentReference = options.documentReference;
            if (!options.displayText && !existingHyperlink.textToDisplay) {
                hyperlinkObj.textToDisplay = options.documentReference;
            }
            logDiag(`[editHyperlink] Updating to internal link: ${options.documentReference}`);
        } else {
            // Keep existing link type
            if (existingHyperlink.address) {
                hyperlinkObj.address = existingHyperlink.address;
            } else if (existingHyperlink.documentReference) {
                hyperlinkObj.documentReference = existingHyperlink.documentReference;
            } else {
                throw new Error("No existing hyperlink to edit and no new link provided");
            }
        }
        
        range.hyperlink = hyperlinkObj;
        await ctx.sync();
        
        logDiag(`[editHyperlink] Successfully edited hyperlink`);
    } catch (e) {
        throw new Error(`Failed to edit hyperlink: ${e.message}`);
    }
}

// ============================================================================
// Data Type Functions
// ============================================================================

/**
 * Checks if data types API is available (ExcelApi 1.16+)
 * @param {Excel.RequestContext} ctx - Excel context
 * @returns {Promise<boolean>}
 */
async function isDataTypesSupported(ctx) {
    try {
        const testRange = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1");
        testRange.load("valuesAsJson");
        await ctx.sync();
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * Inserts a custom data type (EntityCellValue) into a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} range - Target cell (single cell only)
 * @param {Object} action - Action with entity data
 */
async function insertDataType(ctx, sheet, range, action) {
    logDiag(`[insertDataType] Starting at "${action.target}"`);
    
    // Check API support
    const supported = await isDataTypesSupported(ctx);
    if (!supported) {
        throw new Error("Data types require Excel 365, Excel 2021, or Excel Online. Your version does not support this feature.");
    }
    
    let options = { text: "", basicValue: "", properties: {} };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`[insertDataType] Warning: Failed to parse action.data`);
        }
    }
    
    // Validate single cell
    range.load(["rowCount", "columnCount"]);
    await ctx.sync();
    
    if (range.rowCount !== 1 || range.columnCount !== 1) {
        throw new Error(`insertDataType requires single cell target, got ${range.rowCount}x${range.columnCount}`);
    }
    
    // Build EntityCellValue JSON
    const entityValue = {
        type: "Entity",
        text: options.text || "Entity",
        basicType: "String",
        basicValue: options.basicValue || options.text || "Entity",
        properties: {}
    };
    
    // Add properties (convert to CellValue types)
    for (const [key, value] of Object.entries(options.properties || {})) {
        if (typeof value === "number") {
            entityValue.properties[key] = { type: "Double", basicValue: value };
        } else if (typeof value === "boolean") {
            entityValue.properties[key] = { type: "Boolean", basicValue: value };
        } else {
            entityValue.properties[key] = { type: "String", basicValue: String(value) };
        }
    }
    
    try {
        range.valuesAsJson = [[entityValue]];
        await ctx.sync();
        logDiag(`[insertDataType] Successfully inserted entity "${options.text}" at ${action.target}`);
    } catch (e) {
        throw new Error(`Failed to insert data type: ${e.message}`);
    }
}

/**
 * Refreshes a data type cell by updating its properties
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} range - Target cell
 * @param {Object} action - Action with updated properties
 */
async function refreshDataType(ctx, sheet, range, action) {
    logDiag(`[refreshDataType] Starting at "${action.target}"`);
    
    // Check API support
    const supported = await isDataTypesSupported(ctx);
    if (!supported) {
        throw new Error("Data types require Excel 365, Excel 2021, or Excel Online. Your version does not support this feature.");
    }
    
    // Load current cell value
    range.load(["valueTypes", "valuesAsJson"]);
    await ctx.sync();
    
    const cellType = range.valueTypes[0][0];
    if (cellType !== "Entity" && cellType !== "LinkedEntity") {
        throw new Error(`Cell ${action.target} is not a data type (type: ${cellType})`);
    }
    
    if (cellType === "LinkedEntity") {
        logDiag(`[refreshDataType] Warning: LinkedEntity cells (Stocks, Geography) auto-refresh from service. Manual refresh not supported.`);
        return; // No-op for linked entities
    }
    
    // Update custom entity properties
    let options = { properties: {} };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logDiag(`[refreshDataType] Warning: Failed to parse action.data`);
        }
    }
    
    const currentEntity = range.valuesAsJson[0][0];
    const updatedEntity = { ...currentEntity };
    
    // Ensure properties object exists before merging
    updatedEntity.properties = currentEntity.properties ? { ...currentEntity.properties } : {};
    
    // Merge updated properties
    for (const [key, value] of Object.entries(options.properties || {})) {
        if (typeof value === "number") {
            updatedEntity.properties[key] = { type: "Double", basicValue: value };
        } else if (typeof value === "boolean") {
            updatedEntity.properties[key] = { type: "Boolean", basicValue: value };
        } else {
            updatedEntity.properties[key] = { type: "String", basicValue: String(value) };
        }
    }
    
    try {
        range.valuesAsJson = [[updatedEntity]];
        await ctx.sync();
        logDiag(`[refreshDataType] Successfully refreshed entity at ${action.target}`);
    } catch (e) {
        throw new Error(`Failed to refresh data type: ${e.message}`);
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
    deleteSlicer,
    createNamedRange,
    deleteNamedRange,
    updateNamedRange,
    listNamedRanges,
    protectWorksheet,
    unprotectWorksheet,
    protectRange,
    unprotectRange,
    protectWorkbook,
    unprotectWorkbook,
    insertShape,
    insertImage,
    insertTextBox,
    formatShape,
    deleteShape,
    groupShapes,
    arrangeShapes,
    ungroupShapes,
    addComment,
    addNote,
    editComment,
    editNote,
    deleteComment,
    deleteNote,
    replyToComment,
    resolveComment,
    createSparkline,
    configureSparkline,
    deleteSparkline,
    renameSheet,
    moveSheet,
    hideSheet,
    unhideSheet,
    freezePanes,
    unfreezePane,
    setZoom,
    splitPane,
    createView,
    setPageSetup,
    setPageMargins,
    setPageOrientation,
    setPrintArea,
    setHeaderFooter,
    setPageBreaks,
    insertDataType,
    refreshDataType,
    addHyperlink,
    removeHyperlink,
    editHyperlink
};
