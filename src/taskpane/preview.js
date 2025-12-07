/**
 * Preview Panel Module
 * Contains functions for rendering and managing the preview panel
 * 
 * Supported Action Types (90 total):
 * 
 * BASIC OPERATIONS (6):
 *   formula, values, format, validation, sort, autofill
 * ADVANCED FORMATTING (2):
 *   conditionalFormat, clearFormat
 * CHARTS (2):
 *   chart, pivotChart
 * COPY/FILTER/DUPLICATES (5):
 *   copy, copyValues, filter, clearFilter, removeDuplicates
 * SHEET MANAGEMENT (1):
 *   sheet
 * TABLE OPERATIONS (7):
 *   createTable, styleTable, addTableRow, addTableColumn,
 *   resizeTable, convertToRange, toggleTableTotals
 * DATA MANIPULATION (8):
 *   insertRows, insertColumns, deleteRows, deleteColumns,
 *   mergeCells, unmergeCells, findReplace, textToColumns
 * PIVOTTABLE OPERATIONS (5):
 *   createPivotTable, addPivotField, configurePivotLayout,
 *   refreshPivotTable, deletePivotTable
 * SLICER OPERATIONS (5):
 *   createSlicer, configureSlicer, connectSlicerToTable,
 *   connectSlicerToPivot, deleteSlicer
 * NAMED RANGE OPERATIONS (4):
 *   createNamedRange, deleteNamedRange, updateNamedRange, listNamedRanges
 * PROTECTION OPERATIONS (6):
 *   protectWorksheet, unprotectWorksheet, protectRange, unprotectRange,
 *   protectWorkbook, unprotectWorkbook
 * SHAPE OPERATIONS (9):
 *   insertShape, insertImage, insertTextBox, formatShape, deleteShape,
 *   groupShapes, arrangeShapes, ungroupShapes
 * COMMENT OPERATIONS (8):
 *   addComment, addNote, editComment, editNote, deleteComment,
 *   deleteNote, replyToComment, resolveComment
 * SPARKLINE OPERATIONS (3):
 *   createSparkline, configureSparkline, deleteSparkline
 * WORKSHEET MANAGEMENT (9):
 *   renameSheet, moveSheet, hideSheet, unhideSheet, freezePanes,
 *   unfreezePane, setZoom, splitPane, createView
 * PAGE SETUP OPERATIONS (6):
 *   setPageSetup, setPageMargins, setPageOrientation, setPrintArea,
 *   setHeaderFooter, setPageBreaks
 * DATA TYPE OPERATIONS (2):
 *   insertDataType, refreshDataType
 * HYPERLINK OPERATIONS (3):
 *   addHyperlink, removeHyperlink, editHyperlink
 */

// Action types supported by the preview panel (90 total)
const ACTION_TYPES = [
    // Basic Operations
    'formula', 'values', 'format', 'validation', 'sort', 'autofill',
    // Advanced Formatting
    'conditionalFormat', 'clearFormat',
    // Charts
    'chart', 'pivotChart',
    // Copy/Filter/Duplicates
    'copy', 'copyValues', 'filter', 'clearFilter', 'removeDuplicates',
    // Sheet Management
    'sheet',
    // Table Operations
    'createTable', 'styleTable', 'addTableRow', 'addTableColumn', 'resizeTable', 'convertToRange', 'toggleTableTotals',
    // Data Manipulation
    'insertRows', 'insertColumns', 'deleteRows', 'deleteColumns', 'mergeCells', 'unmergeCells', 'findReplace', 'textToColumns',
    // PivotTable Operations
    'createPivotTable', 'addPivotField', 'configurePivotLayout', 'refreshPivotTable', 'deletePivotTable',
    // Slicer Operations
    'createSlicer', 'configureSlicer', 'connectSlicerToTable', 'connectSlicerToPivot', 'deleteSlicer',
    // Named Range Operations
    'createNamedRange', 'deleteNamedRange', 'updateNamedRange', 'listNamedRanges',
    // Protection Operations
    'protectWorksheet', 'unprotectWorksheet', 'protectRange', 'unprotectRange', 'protectWorkbook', 'unprotectWorkbook',
    // Shape Operations
    'insertShape', 'insertImage', 'insertTextBox', 'formatShape', 'deleteShape', 'groupShapes', 'arrangeShapes', 'ungroupShapes',
    // Comment Operations
    'addComment', 'addNote', 'editComment', 'editNote', 'deleteComment', 'deleteNote', 'replyToComment', 'resolveComment',
    // Sparkline Operations
    'createSparkline', 'configureSparkline', 'deleteSparkline',
    // Worksheet Management
    'renameSheet', 'moveSheet', 'hideSheet', 'unhideSheet', 'freezePanes', 'unfreezePane', 'setZoom', 'splitPane', 'createView',
    // Page Setup Operations
    'setPageSetup', 'setPageMargins', 'setPageOrientation', 'setPrintArea', 'setHeaderFooter', 'setPageBreaks',
    // Data Type Operations
    'insertDataType', 'refreshDataType',
    // Hyperlink Operations
    'addHyperlink', 'removeHyperlink', 'editHyperlink'
];

/**
 * Human-readable labels for all 87 action types
 * Shared constant to avoid recreation on each getActionSummary call
 */
const ACTION_TYPE_LABELS = {
    // Basic Operations
    formula: "Formula",
    values: "Values",
    format: "Format",
    chart: "Chart",
    validation: "Dropdown",
    sort: "Sort",
    autofill: "Autofill",
    // Advanced Formatting
    conditionalFormat: "Conditional Format",
    clearFormat: "Clear Format",
    // Charts
    pivotChart: "Pivot Chart",
    // Copy/Filter/Duplicates
    copy: "Copy",
    copyValues: "Copy Values",
    filter: "Filter",
    clearFilter: "Clear Filter",
    removeDuplicates: "Remove Duplicates",
    // Sheet Management
    sheet: "Create Sheet",
    // Table Operations
    createTable: "Create Table",
    styleTable: "Style Table",
    addTableRow: "Add Table Row",
    addTableColumn: "Add Table Column",
    resizeTable: "Resize Table",
    convertToRange: "Convert to Range",
    toggleTableTotals: "Toggle Totals",
    // Data Manipulation
    insertRows: "Insert Rows",
    insertColumns: "Insert Columns",
    deleteRows: "Delete Rows",
    deleteColumns: "Delete Columns",
    mergeCells: "Merge Cells",
    unmergeCells: "Unmerge Cells",
    findReplace: "Find & Replace",
    textToColumns: "Text to Columns",
    // PivotTable Operations
    createPivotTable: "Create PivotTable",
    addPivotField: "Add Pivot Field",
    configurePivotLayout: "Configure Pivot",
    refreshPivotTable: "Refresh PivotTable",
    deletePivotTable: "Delete PivotTable",
    // Slicer Operations
    createSlicer: "Create Slicer",
    configureSlicer: "Configure Slicer",
    connectSlicerToTable: "Connect Slicer to Table",
    connectSlicerToPivot: "Connect Slicer to Pivot",
    deleteSlicer: "Delete Slicer",
    // Named Range Operations
    createNamedRange: "Create Named Range",
    deleteNamedRange: "Delete Named Range",
    updateNamedRange: "Update Named Range",
    listNamedRanges: "List Named Ranges",
    // Protection Operations
    protectWorksheet: "Protect Sheet",
    unprotectWorksheet: "Unprotect Sheet",
    protectRange: "Protect Range",
    unprotectRange: "Unprotect Range",
    protectWorkbook: "Protect Workbook",
    unprotectWorkbook: "Unprotect Workbook",
    // Shape Operations
    insertShape: "Insert Shape",
    insertImage: "Insert Image",
    insertTextBox: "Insert Text Box",
    formatShape: "Format Shape",
    deleteShape: "Delete Shape",
    groupShapes: "Group Shapes",
    arrangeShapes: "Arrange Shapes",
    ungroupShapes: "Ungroup Shapes",
    // Comment Operations
    addComment: "Add Comment",
    addNote: "Add Note",
    editComment: "Edit Comment",
    editNote: "Edit Note",
    deleteComment: "Delete Comment",
    deleteNote: "Delete Note",
    replyToComment: "Reply to Comment",
    resolveComment: "Resolve Comment",
    // Sparkline Operations
    createSparkline: "Create Sparkline",
    configureSparkline: "Configure Sparkline",
    deleteSparkline: "Delete Sparkline",
    // Worksheet Management
    renameSheet: "Rename Sheet",
    moveSheet: "Move Sheet",
    hideSheet: "Hide Sheet",
    unhideSheet: "Unhide Sheet",
    freezePanes: "Freeze Panes",
    unfreezePane: "Unfreeze Panes",
    setZoom: "Set Zoom",
    splitPane: "Split Panes",
    createView: "Create View",
    // Page Setup Operations
    setPageSetup: "Page Setup",
    setPageMargins: "Set Margins",
    setPageOrientation: "Set Orientation",
    setPrintArea: "Set Print Area",
    setHeaderFooter: "Set Header/Footer",
    setPageBreaks: "Set Page Breaks",
    // Data Type Operations
    insertDataType: "Insert Entity",
    refreshDataType: "Refresh Entity",
    // Hyperlink Operations
    addHyperlink: "Add Hyperlink",
    removeHyperlink: "Remove Hyperlink",
    editHyperlink: "Edit Hyperlink"
};

/**
 * Gets the SVG icon for an action type
 * @param {string} type - Action type
 * @returns {string} SVG icon markup
 */
function getActionIcon(type) {
    const icons = {
        formula: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7h16M4 12h16M4 17h10"/><text x="18" y="18" font-size="8" fill="currentColor">fx</text></svg>',
        values: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>',
        format: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 20h16M6 16l6-12 6 12M8 12h8"/></svg>',
        chart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>',
        validation: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>',
        sort: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h12M3 18h6"/></svg>',
        autofill: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>',
        // Sparkline icons - compact trend visualization
        createSparkline: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="8" width="18" height="8" rx="1"/><path d="M6 12l3-2 3 3 3-4 3 3"/></svg>',
        configureSparkline: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="8" width="18" height="8" rx="1"/><path d="M6 12l3-2 3 3 3-4 3 3"/><circle cx="18" cy="6" r="3"/></svg>',
        deleteSparkline: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="8" width="18" height="8" rx="1"/><path d="M6 12l3-2 3 3 3-4 3 3"/><path d="M16 4l4 4M20 4l-4 4" stroke="#C00000"/></svg>',
        // Worksheet management icons
        renameSheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="14" width="18" height="6" rx="1"/><path d="M7 17h4M15 17h2"/><path d="M11 4l2 2-2 2"/></svg>',
        moveSheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="14" width="18" height="6" rx="1"/><path d="M12 4v6M9 7l3 3 3-3"/></svg>',
        hideSheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="14" width="18" height="6" rx="1"/><path d="M4 8l16-4M4 4l16 4"/></svg>',
        unhideSheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="14" width="18" height="6" rx="1"/><circle cx="12" cy="6" r="3"/><path d="M6 6h2M16 6h2"/></svg>',
        freezePanes: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18" stroke-dasharray="2 2"/></svg>',
        unfreezePane: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M8 8l8 8M16 8l-8 8"/></svg>',
        setZoom: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="10" cy="10" r="6"/><path d="M14 14l6 6"/><path d="M8 10h4M10 8v4"/></svg>',
        splitPane: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M12 3v18M3 12h18"/></svg>',
        createView: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="12" cy="12" r="3"/><path d="M3 12h3M18 12h3"/></svg>',
        // Data type icons - entity cards
        insertDataType: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h6M7 16h8"/><circle cx="17" cy="14" r="2"/></svg>',
        refreshDataType: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h6"/><path d="M17 11v6M14 14h6"/></svg>',
        // Advanced Formatting icons
        conditionalFormat: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h10M7 16h10"/><circle cx="17" cy="8" r="2" fill="#4CAF50"/><circle cx="17" cy="12" r="2" fill="#FFC107"/><circle cx="17" cy="16" r="2" fill="#F44336"/></svg>',
        clearFormat: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 20h16M6 16l6-12 6 12"/><path d="M17 4l-6 6M11 10l6-6" stroke="#C00000"/></svg>',
        // Chart icons
        pivotChart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 20V10M12 20V4M6 20v-6"/><rect x="2" y="2" width="6" height="6" rx="1"/></svg>',
        // Copy/Filter/Duplicates icons
        copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>',
        copyValues: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/><text x="13" y="18" font-size="6" fill="currentColor">123</text></svg>',
        filter: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z"/></svg>',
        clearFilter: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z"/><path d="M17 8l4 4M21 8l-4 4" stroke="#C00000"/></svg>',
        removeDuplicates: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1" stroke-dasharray="2 2"/><path d="M16 17l4 4M20 17l-4 4" stroke="#C00000"/></svg>',
        // Sheet Management icons
        sheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M12 8v8M8 12h8"/></svg>',
        // Table Operations icons
        createTable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><path d="M17 6v0" stroke-width="3"/></svg>',
        styleTable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><circle cx="18" cy="18" r="3"/></svg>',
        addTableRow: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="14" rx="2"/><path d="M3 9h18M9 17V9"/><path d="M12 19v4M10 21h4"/></svg>',
        addTableColumn: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="14" height="18" rx="2"/><path d="M3 9h14M9 21V9"/><path d="M19 12h4M21 10v4"/></svg>',
        resizeTable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><path d="M15 15l6 6M21 15v6h-6"/></svg>',
        convertToRange: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="8" height="8" rx="1"/><path d="M14 7h7M14 12h7M14 17h7M3 14h8M3 19h8"/></svg>',
        toggleTableTotals: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 21V9"/><text x="14" y="19" font-size="8" fill="currentColor">Σ</text></svg>',
        // Data Manipulation icons
        insertRows: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h18M3 18h18"/><path d="M12 9v6M9 12h6" stroke="#4CAF50"/></svg>',
        insertColumns: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 3v18M12 3v18M18 3v18"/><path d="M9 12h6M12 9v6" stroke="#4CAF50"/></svg>',
        deleteRows: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h18M3 18h18"/><path d="M9 12h6" stroke="#C00000" stroke-width="3"/></svg>',
        deleteColumns: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 3v18M12 3v18M18 3v18"/><path d="M9 12h6" stroke="#C00000" stroke-width="3"/></svg>',
        mergeCells: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="8" height="8"/><rect x="13" y="3" width="8" height="8"/><rect x="3" y="13" width="18" height="8"/><path d="M8 7l4 4 4-4"/></svg>',
        unmergeCells: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="8"/><rect x="3" y="13" width="8" height="8"/><rect x="13" y="13" width="8" height="8"/><path d="M12 11l-4-4M12 11l4-4"/></svg>',
        findReplace: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="10" cy="10" r="6"/><path d="M14 14l6 6"/><path d="M17 10h4M19 8v4"/></svg>',
        textToColumns: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="6" width="6" height="12"/><rect x="11" y="6" width="4" height="12"/><rect x="17" y="6" width="4" height="12"/><path d="M6 3l6 3-6 3"/></svg>',
        // PivotTable Operations icons
        createPivotTable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/><text x="14" y="16" font-size="8" fill="currentColor">Σ</text></svg>',
        addPivotField: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/><path d="M15 12v6M12 15h6"/></svg>',
        configurePivotLayout: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/><circle cx="17" cy="17" r="3"/></svg>',
        refreshPivotTable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/><path d="M17 12a3 3 0 11-3 3M17 12v3h3"/></svg>',
        deletePivotTable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/><path d="M14 14l4 4M18 14l-4 4" stroke="#C00000"/></svg>',
        // Slicer Operations icons
        createSlicer: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h10M7 16h10"/><rect x="5" y="7" width="3" height="3" rx="1" fill="currentColor"/></svg>',
        configureSlicer: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h10M7 16h10"/><circle cx="18" cy="18" r="3"/></svg>',
        connectSlicerToTable: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="6" width="8" height="12" rx="1"/><rect x="14" y="6" width="8" height="12" rx="1"/><path d="M10 12h4M12 10l2 2-2 2"/></svg>',
        connectSlicerToPivot: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="6" width="8" height="12" rx="1"/><rect x="14" y="6" width="8" height="12" rx="1"/><path d="M14 9h8M17 6v12"/><path d="M10 12h4M12 10l2 2-2 2"/></svg>',
        deleteSlicer: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h10M7 16h10"/><path d="M15 6l4 4M19 6l-4 4" stroke="#C00000"/></svg>',
        // Named Range Operations icons
        createNamedRange: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="8" width="12" height="8" rx="1"/><path d="M17 10h4M17 14h4"/><path d="M7 11h4M7 13h2"/></svg>',
        deleteNamedRange: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="8" width="12" height="8" rx="1"/><path d="M7 11h4M7 13h2"/><path d="M17 10l4 4M21 10l-4 4" stroke="#C00000"/></svg>',
        updateNamedRange: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="8" width="12" height="8" rx="1"/><path d="M7 11h4M7 13h2"/><path d="M18 8v8M16 10l2-2 2 2M16 14l2 2 2-2"/></svg>',
        listNamedRanges: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 6h16M4 12h16M4 18h16"/><rect x="2" y="4" width="4" height="4" rx="1"/><rect x="2" y="10" width="4" height="4" rx="1"/><rect x="2" y="16" width="4" height="4" rx="1"/></svg>',
        // Protection Operations icons
        protectWorksheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>',
        unprotectWorksheet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0110 0"/></svg>',
        protectRange: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="10" height="10" rx="1"/><rect x="8" y="13" width="10" height="8" rx="2"/><path d="M11 13v-2a3 3 0 016 0v2"/></svg>',
        unprotectRange: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="10" height="10" rx="1"/><rect x="8" y="13" width="10" height="8" rx="2"/><path d="M11 13v-2a3 3 0 016 0"/></svg>',
        protectWorkbook: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19.5A2.5 2.5 0 016.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z"/><rect x="10" y="10" width="6" height="5" rx="1"/><path d="M12 10V8a2 2 0 014 0v2"/></svg>',
        unprotectWorkbook: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19.5A2.5 2.5 0 016.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z"/><rect x="10" y="10" width="6" height="5" rx="1"/><path d="M12 10V8a2 2 0 014 0"/></svg>',
        // Shape Operations icons
        insertShape: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="8" r="4"/><rect x="12" y="12" width="8" height="8"/><path d="M4 20l4-8 4 8z"/></svg>',
        insertImage: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/></svg>',
        insertTextBox: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M7 9h10M7 13h6"/></svg>',
        formatShape: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="12" height="12" rx="2"/><circle cx="18" cy="18" r="4"/><path d="M18 16v4M16 18h4"/></svg>',
        deleteShape: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="12" height="12" rx="2"/><path d="M17 7l4 4M21 7l-4 4" stroke="#C00000"/></svg>',
        groupShapes: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="1" y="1" width="22" height="22" rx="2" stroke-dasharray="3 3"/></svg>',
        arrangeShapes: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="10" width="8" height="8"/><rect x="8" y="6" width="8" height="8" fill="white"/><rect x="13" y="2" width="8" height="8"/></svg>',
        ungroupShapes: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><path d="M10 6l2-3 2 3M10 18l2 3 2-3M6 10l-3 2 3 2M18 10l3 2-3 2"/></svg>',
        // Comment Operations icons
        addComment: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/><path d="M12 8v4M10 10h4"/></svg>',
        addNote: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 8h10M7 12h10M7 16h6"/></svg>',
        editComment: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/><path d="M15 6l3 3M10 11l5-5 3 3-5 5H10v-3z"/></svg>',
        editNote: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M15 6l3 3M10 11l5-5 3 3-5 5H10v-3z"/></svg>',
        deleteComment: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/><path d="M9 9l6 6M15 9l-6 6" stroke="#C00000"/></svg>',
        deleteNote: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 9l6 6M15 9l-6 6" stroke="#C00000"/></svg>',
        replyToComment: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/><path d="M7 10l3 3-3 3"/></svg>',
        resolveComment: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/><path d="M9 11l2 2 4-4" stroke="#4CAF50"/></svg>',
        // Page Setup Operations icons
        setPageSetup: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="2" width="14" height="20" rx="2"/><circle cx="17" cy="17" r="4"/><path d="M17 15v4M15 17h4"/></svg>',
        setPageMargins: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="2" width="14" height="20" rx="2"/><rect x="7" y="5" width="10" height="14" stroke-dasharray="2 2"/></svg>',
        setPageOrientation: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="6" y="3" width="12" height="16" rx="2"/><path d="M15 19l3 3 3-3M18 22v-6"/></svg>',
        setPrintArea: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="2" width="14" height="20" rx="2"/><rect x="8" y="6" width="8" height="8" stroke-dasharray="3 3"/></svg>',
        setHeaderFooter: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="2" width="14" height="20" rx="2"/><path d="M5 6h14M5 18h14"/><path d="M8 4h8M8 20h8" stroke-width="1"/></svg>',
        setPageBreaks: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="2" width="14" height="20" rx="2"/><path d="M5 12h14" stroke-dasharray="4 2"/></svg>',
        // Hyperlink Operations icons
        addHyperlink: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>',
        removeHyperlink: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/><line x1="3" y1="3" x2="21" y2="21" stroke="#C00000"/></svg>',
        editHyperlink: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/><path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"/></svg>'
    };
    return icons[type] || icons.formula;
}

/**
 * Gets a summary string for an action
 * @param {Object} action - The action to summarize
 * @returns {string} Human-readable summary
 */
function getActionSummary(action) {
    return ACTION_TYPE_LABELS[action.type] || action.type;
}

/**
 * Gets detailed description for an action
 * @param {Object} action - The action to describe
 * @returns {string} Detailed description
 */
function getActionDetails(action) {
    switch (action.type) {
        case "formula":
            return action.data || "No formula";
        case "values":
            try {
                const vals = JSON.parse(action.data);
                return JSON.stringify(vals, null, 2);
            } catch {
                return action.data || "No values";
            }
        case "format":
            try {
                const fmt = JSON.parse(action.data);
                const parts = [];
                if (fmt.bold) parts.push("Bold");
                if (fmt.italic) parts.push("Italic");
                if (fmt.fill) parts.push(`Fill: ${fmt.fill}`);
                if (fmt.fontColor) parts.push(`Color: ${fmt.fontColor}`);
                if (fmt.fontSize) parts.push(`Size: ${fmt.fontSize}`);
                if (fmt.numberFormat) parts.push(`Format: ${fmt.numberFormat}`);
                return parts.join(", ") || "No formatting";
            } catch {
                return action.data || "No format data";
            }
        case "chart":
            return `Type: ${action.chartType}\nData: ${action.target}\nPosition: ${action.position}${action.title ? `\nTitle: ${action.title}` : ""}`;
        case "validation":
            return `Source: ${action.source}`;
        case "sort":
            return action.data || "Default sort";
        case "autofill":
            return `Source: ${action.source}`;
        case "createSparkline":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Type: ${opts.type || "Line"}`];
                if (opts.sourceData) parts.push(`Source: ${opts.sourceData}`);
                if (opts.colors && opts.colors.series) parts.push(`Color: ${opts.colors.series}`);
                return parts.join("\n");
            } catch {
                return action.data || "Create sparkline";
            }
        case "configureSparkline":
            try {
                const opts = JSON.parse(action.data);
                const parts = [];
                if (opts.colors) {
                    const colorParts = [];
                    if (opts.colors.series) colorParts.push(`Series: ${opts.colors.series}`);
                    if (opts.colors.high) colorParts.push(`High: ${opts.colors.high}`);
                    if (opts.colors.low) colorParts.push(`Low: ${opts.colors.low}`);
                    if (colorParts.length > 0) parts.push(`Colors: ${colorParts.join(", ")}`);
                }
                if (opts.markers) {
                    const markerParts = [];
                    if (opts.markers.high) markerParts.push("High");
                    if (opts.markers.low) markerParts.push("Low");
                    if (opts.markers.first) markerParts.push("First");
                    if (opts.markers.last) markerParts.push("Last");
                    if (markerParts.length > 0) parts.push(`Markers: ${markerParts.join(", ")}`);
                }
                if (opts.axes && opts.axes.horizontal !== undefined) {
                    parts.push(`Axis: ${opts.axes.horizontal ? "Visible" : "Hidden"}`);
                }
                return parts.join("\n") || "Configure sparkline properties";
            } catch {
                return action.data || "Configure sparkline";
            }
        case "deleteSparkline":
            return `Remove sparkline at ${action.target}`;
        case "renameSheet":
            try {
                const opts = JSON.parse(action.data);
                return `Rename "${action.target}" to "${opts.newName}"`;
            } catch {
                return `Rename sheet ${action.target}`;
            }
        case "moveSheet":
            try {
                const opts = JSON.parse(action.data);
                let desc = `Move "${action.target}" to ${opts.position}`;
                if (opts.referenceSheet) desc += ` "${opts.referenceSheet}"`;
                return desc;
            } catch {
                return `Move sheet ${action.target}`;
            }
        case "hideSheet":
            return `Hide sheet "${action.target}"`;
        case "unhideSheet":
            return `Unhide sheet "${action.target}"`;
        case "freezePanes":
            try {
                const opts = JSON.parse(action.data);
                return `Freeze ${opts.freezeType || "both"} at ${action.target}`;
            } catch {
                return `Freeze panes at ${action.target}`;
            }
        case "unfreezePane":
            return `Unfreeze all panes`;
        case "setZoom":
            try {
                const opts = JSON.parse(action.data);
                return `Set zoom to ${opts.zoomLevel}%`;
            } catch {
                return `Set zoom level`;
            }
        case "splitPane":
            return `Split panes at ${action.target}`;
        case "createView":
            return `Create custom view "${action.target}"`;
        case "insertDataType":
            try {
                const opts = JSON.parse(action.data);
                const propCount = Object.keys(opts.properties || {}).length;
                return `Insert entity "${opts.text || 'Entity'}" with ${propCount} properties`;
            } catch {
                return `Insert entity at ${action.target}`;
            }
        case "refreshDataType":
            try {
                const opts = JSON.parse(action.data);
                const propCount = Object.keys(opts.properties || {}).length;
                return `Update ${propCount} properties at ${action.target}`;
            } catch {
                return `Refresh entity at ${action.target}`;
            }
        // Advanced Formatting
        case "conditionalFormat":
            try {
                const opts = JSON.parse(action.data);
                const ruleType = opts.type || opts.ruleType || "custom";
                const parts = [`Rule: ${ruleType}`];
                if (opts.criteria) parts.push(`Criteria: ${opts.criteria}`);
                if (opts.format) parts.push(`Format: ${JSON.stringify(opts.format)}`);
                return parts.join("\n");
            } catch {
                return action.data || "Apply conditional formatting";
            }
        case "clearFormat":
            return "Clear all conditional formatting";
        // Charts
        case "pivotChart":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Type: ${opts.chartType || "Column"}`];
                if (opts.pivotTable) parts.push(`Source: ${opts.pivotTable}`);
                if (opts.title) parts.push(`Title: ${opts.title}`);
                return parts.join("\n");
            } catch {
                return action.data || "Create pivot chart";
            }
        // Copy/Filter/Duplicates
        case "copy":
            return `Copy ${action.source || action.target} → ${action.target}`;
        case "copyValues":
            return `Copy values from ${action.source} → ${action.target}`;
        case "filter":
            try {
                const opts = JSON.parse(action.data);
                const parts = [];
                if (opts.column) parts.push(`Column: ${opts.column}`);
                if (opts.criteria) parts.push(`Criteria: ${opts.criteria}`);
                if (opts.operator) parts.push(`Operator: ${opts.operator}`);
                return parts.join("\n") || "Apply filter";
            } catch {
                return action.data || "Apply filter";
            }
        case "clearFilter":
            return "Clear all filters";
        case "removeDuplicates":
            try {
                const opts = JSON.parse(action.data);
                const cols = opts.columns ? opts.columns.join(", ") : "all columns";
                return `Remove duplicates based on: ${cols}`;
            } catch {
                return action.data || "Remove duplicate rows";
            }
        // Sheet Management
        case "sheet":
            try {
                const opts = JSON.parse(action.data);
                return `Create sheet "${opts.name || action.target}"`;
            } catch {
                return `Create sheet "${action.target}"`;
            }
        // Table Operations
        case "createTable":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Range: ${action.target}`];
                if (opts.tableName) parts.push(`Name: ${opts.tableName}`);
                if (opts.style) parts.push(`Style: ${opts.style}`);
                if (opts.hasHeaders !== undefined) parts.push(`Headers: ${opts.hasHeaders ? "Yes" : "No"}`);
                return parts.join("\n");
            } catch {
                return `Create table at ${action.target}`;
            }
        case "styleTable":
            try {
                const opts = JSON.parse(action.data);
                return `Style "${action.target}" with ${opts.style || "default style"}`;
            } catch {
                return `Style table ${action.target}`;
            }
        case "addTableRow":
            try {
                const opts = JSON.parse(action.data);
                const pos = opts.position || "end";
                return `Add row at ${pos} of "${action.target}"`;
            } catch {
                return `Add row to ${action.target}`;
            }
        case "addTableColumn":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Table: ${action.target}`];
                if (opts.columnName) parts.push(`Column: ${opts.columnName}`);
                if (opts.position) parts.push(`Position: ${opts.position}`);
                return parts.join("\n");
            } catch {
                return `Add column to ${action.target}`;
            }
        case "resizeTable":
            try {
                const opts = JSON.parse(action.data);
                return `Resize "${action.target}" to ${opts.newRange}`;
            } catch {
                return `Resize table ${action.target}`;
            }
        case "convertToRange":
            return `Convert table "${action.target}" to range`;
        case "toggleTableTotals":
            try {
                const opts = JSON.parse(action.data);
                const state = opts.show ? "Show" : "Hide";
                return `${state} totals for "${action.target}"`;
            } catch {
                return `Toggle totals for ${action.target}`;
            }
        // Data Manipulation
        case "insertRows":
            try {
                const opts = JSON.parse(action.data);
                const count = opts.count || 1;
                return `Insert ${count} row(s) at ${action.target}`;
            } catch {
                return `Insert rows at ${action.target}`;
            }
        case "insertColumns":
            try {
                const opts = JSON.parse(action.data);
                const count = opts.count || 1;
                return `Insert ${count} column(s) at ${action.target}`;
            } catch {
                return `Insert columns at ${action.target}`;
            }
        case "deleteRows":
            return `Delete rows at ${action.target}`;
        case "deleteColumns":
            return `Delete columns at ${action.target}`;
        case "mergeCells":
            return `Merge cells ${action.target}`;
        case "unmergeCells":
            return `Unmerge cells ${action.target}`;
        case "findReplace":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Find: "${opts.find || opts.search}"`];
                if (opts.replace !== undefined) parts.push(`Replace: "${opts.replace}"`);
                if (opts.matchCase) parts.push("Match case");
                if (opts.matchEntire) parts.push("Match entire cell");
                return parts.join("\n");
            } catch {
                return action.data || "Find and replace";
            }
        case "textToColumns":
            try {
                const opts = JSON.parse(action.data);
                const delim = opts.delimiter || "comma";
                return `Split ${action.target} by ${delim}`;
            } catch {
                return `Split text to columns at ${action.target}`;
            }
        // PivotTable Operations
        case "createPivotTable":
            try {
                const opts = JSON.parse(action.data);
                const parts = [];
                if (opts.source) parts.push(`Source: ${opts.source}`);
                if (opts.destination) parts.push(`Destination: ${opts.destination}`);
                if (opts.name) parts.push(`Name: ${opts.name}`);
                return parts.join("\n") || `Create PivotTable at ${action.target}`;
            } catch {
                return `Create PivotTable at ${action.target}`;
            }
        case "addPivotField":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Field: ${opts.field || opts.fieldName}`];
                if (opts.area) parts.push(`Area: ${opts.area}`);
                if (opts.aggregation) parts.push(`Function: ${opts.aggregation}`);
                return parts.join("\n");
            } catch {
                return `Add field to ${action.target}`;
            }
        case "configurePivotLayout":
            try {
                const opts = JSON.parse(action.data);
                return `Configure "${action.target}" layout: ${opts.layout || "compact"}`;
            } catch {
                return `Configure ${action.target} layout`;
            }
        case "refreshPivotTable":
            return `Refresh PivotTable "${action.target}"`;
        case "deletePivotTable":
            return `Delete PivotTable "${action.target}"`;
        // Slicer Operations
        case "createSlicer":
            try {
                const opts = JSON.parse(action.data);
                const parts = [];
                if (opts.sourceName) parts.push(`Source: ${opts.sourceName}`);
                if (opts.field) parts.push(`Field: ${opts.field}`);
                if (opts.style) parts.push(`Style: ${opts.style}`);
                return parts.join("\n") || "Create slicer";
            } catch {
                return `Create slicer for ${action.target}`;
            }
        case "configureSlicer":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Slicer: ${action.target}`];
                if (opts.style) parts.push(`Style: ${opts.style}`);
                if (opts.selectedItems) parts.push(`Selected: ${opts.selectedItems.join(", ")}`);
                return parts.join("\n");
            } catch {
                return `Configure slicer ${action.target}`;
            }
        case "connectSlicerToTable":
            try {
                const opts = JSON.parse(action.data);
                return `Connect "${action.target}" to table "${opts.tableName}"`;
            } catch {
                return `Connect slicer to table`;
            }
        case "connectSlicerToPivot":
            try {
                const opts = JSON.parse(action.data);
                return `Connect "${action.target}" to PivotTable "${opts.pivotTableName}"`;
            } catch {
                return `Connect slicer to PivotTable`;
            }
        case "deleteSlicer":
            return `Delete slicer "${action.target}"`;
        // Named Range Operations
        case "createNamedRange":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Name: ${opts.name}`];
                if (opts.scope) parts.push(`Scope: ${opts.scope}`);
                if (opts.comment) parts.push(`Comment: ${opts.comment}`);
                return parts.join("\n");
            } catch {
                return `Create named range at ${action.target}`;
            }
        case "deleteNamedRange":
            return `Delete named range "${action.target}"`;
        case "updateNamedRange":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Update: ${action.target}`];
                if (opts.refersTo) parts.push(`New range: ${opts.refersTo}`);
                return parts.join("\n");
            } catch {
                return `Update named range ${action.target}`;
            }
        case "listNamedRanges":
            try {
                const opts = JSON.parse(action.data);
                return `List named ranges (scope: ${opts.scope || "all"})`;
            } catch {
                return "List all named ranges";
            }
        // Protection Operations
        case "protectWorksheet":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Protect sheet "${action.target}"`];
                if (opts.password) parts.push("Password: ****");
                return parts.join("\n");
            } catch {
                return `Protect sheet "${action.target}"`;
            }
        case "unprotectWorksheet":
            return `Unprotect sheet "${action.target}"`;
        case "protectRange":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Protect range ${action.target}`];
                if (opts.title) parts.push(`Title: ${opts.title}`);
                return parts.join("\n");
            } catch {
                return `Protect range ${action.target}`;
            }
        case "unprotectRange":
            return `Unprotect range ${action.target}`;
        case "protectWorkbook":
            return "Protect workbook structure";
        case "unprotectWorkbook":
            return "Unprotect workbook";
        // Shape Operations
        case "insertShape":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Type: ${opts.shapeType || opts.type || "Rectangle"}`];
                if (opts.position) parts.push(`Position: ${JSON.stringify(opts.position)}`);
                return parts.join("\n");
            } catch {
                return `Insert shape at ${action.target}`;
            }
        case "insertImage":
            try {
                const opts = JSON.parse(action.data);
                const size = opts.base64 ? `~${Math.round(opts.base64.length / 1024)}KB` : "image";
                return `Insert ${size} at ${action.target}`;
            } catch {
                return `Insert image at ${action.target}`;
            }
        case "insertTextBox":
            try {
                const opts = JSON.parse(action.data);
                const preview = opts.text ? opts.text.substring(0, 30) : "";
                return `Insert text box: "${preview}${opts.text && opts.text.length > 30 ? "..." : ""}"`;
            } catch {
                return `Insert text box at ${action.target}`;
            }
        case "formatShape":
            try {
                const opts = JSON.parse(action.data);
                const parts = [`Shape: ${action.target}`];
                if (opts.fill) parts.push(`Fill: ${opts.fill}`);
                if (opts.line) parts.push(`Line: ${opts.line}`);
                return parts.join("\n");
            } catch {
                return `Format shape ${action.target}`;
            }
        case "deleteShape":
            return `Delete shape "${action.target}"`;
        case "groupShapes":
            try {
                const opts = JSON.parse(action.data);
                const shapes = opts.shapes ? opts.shapes.join(", ") : action.target;
                return `Group shapes: ${shapes}`;
            } catch {
                return `Group shapes`;
            }
        case "arrangeShapes":
            try {
                const opts = JSON.parse(action.data);
                return `${opts.action || "Arrange"} "${action.target}"`;
            } catch {
                return `Arrange shape ${action.target}`;
            }
        case "ungroupShapes":
            return `Ungroup "${action.target}"`;
        // Comment Operations
        case "addComment":
            try {
                const opts = JSON.parse(action.data);
                const preview = opts.content ? opts.content.substring(0, 50) : "";
                const parts = [`Cell: ${action.target}`];
                if (opts.author) parts.push(`Author: ${opts.author}`);
                if (preview) parts.push(`"${preview}${opts.content && opts.content.length > 50 ? "..." : ""}"`);
                return parts.join("\n");
            } catch {
                return `Add comment at ${action.target}`;
            }
        case "addNote":
            try {
                const opts = JSON.parse(action.data);
                const preview = opts.text ? opts.text.substring(0, 50) : "";
                return `Add note at ${action.target}: "${preview}${opts.text && opts.text.length > 50 ? "..." : ""}"`;
            } catch {
                return `Add note at ${action.target}`;
            }
        case "editComment":
            try {
                const opts = JSON.parse(action.data);
                const preview = opts.content ? opts.content.substring(0, 50) : "";
                return `Edit comment at ${action.target}: "${preview}..."`;
            } catch {
                return `Edit comment at ${action.target}`;
            }
        case "editNote":
            try {
                const opts = JSON.parse(action.data);
                const preview = opts.text ? opts.text.substring(0, 50) : "";
                return `Edit note at ${action.target}: "${preview}..."`;
            } catch {
                return `Edit note at ${action.target}`;
            }
        case "deleteComment":
            return `Delete comment at ${action.target}`;
        case "deleteNote":
            return `Delete note at ${action.target}`;
        case "replyToComment":
            try {
                const opts = JSON.parse(action.data);
                const preview = opts.content ? opts.content.substring(0, 50) : "";
                return `Reply at ${action.target}: "${preview}..."`;
            } catch {
                return `Reply to comment at ${action.target}`;
            }
        case "resolveComment":
            return `Resolve comment at ${action.target}`;
        // Page Setup Operations
        case "setPageSetup":
            try {
                const opts = JSON.parse(action.data);
                const parts = [];
                if (opts.orientation) parts.push(`Orientation: ${opts.orientation}`);
                if (opts.paperSize) parts.push(`Paper: ${opts.paperSize}`);
                if (opts.scaling) parts.push(`Scale: ${opts.scaling}%`);
                return parts.join("\n") || "Configure page setup";
            } catch {
                return action.data || "Configure page setup";
            }
        case "setPageMargins":
            try {
                const opts = JSON.parse(action.data);
                const parts = [];
                if (opts.top) parts.push(`Top: ${opts.top}"`);
                if (opts.bottom) parts.push(`Bottom: ${opts.bottom}"`);
                if (opts.left) parts.push(`Left: ${opts.left}"`);
                if (opts.right) parts.push(`Right: ${opts.right}"`);
                return parts.join(", ") || "Set page margins";
            } catch {
                return action.data || "Set page margins";
            }
        case "setPageOrientation":
            try {
                const opts = JSON.parse(action.data);
                return `Set orientation: ${opts.orientation || "Portrait"}`;
            } catch {
                return action.data || "Set page orientation";
            }
        case "setPrintArea":
            return `Set print area: ${action.target}`;
        case "setHeaderFooter":
            try {
                const opts = JSON.parse(action.data);
                const parts = [];
                if (opts.header) parts.push(`Header: ${opts.header}`);
                if (opts.footer) parts.push(`Footer: ${opts.footer}`);
                return parts.join("\n") || "Set header/footer";
            } catch {
                return action.data || "Set header/footer";
            }
        case "setPageBreaks":
            try {
                const opts = JSON.parse(action.data);
                const type = opts.type || "row";
                return `Add ${type} break at ${action.target}`;
            } catch {
                return `Set page break at ${action.target}`;
            }
        // Hyperlink Operations
        case "addHyperlink":
            try {
                const parsed = JSON.parse(action.data || "{}");
                if (parsed.url) {
                    return `Add web link to <strong>${parsed.url}</strong>${parsed.displayText ? ` (display: "${parsed.displayText}")` : ""}`;
                } else if (parsed.email) {
                    return `Add email link to <strong>${parsed.email}</strong>${parsed.displayText ? ` (display: "${parsed.displayText}")` : ""}`;
                } else if (parsed.documentReference) {
                    return `Add internal link to <strong>${parsed.documentReference}</strong>${parsed.displayText ? ` (display: "${parsed.displayText}")` : ""}`;
                }
                return "Add hyperlink";
            } catch {
                return "Add hyperlink";
            }
        case "removeHyperlink":
            return "Remove hyperlink from cells";
        case "editHyperlink":
            try {
                const parsed = JSON.parse(action.data || "{}");
                const updates = [];
                if (parsed.url) updates.push(`URL: ${parsed.url}`);
                if (parsed.email) updates.push(`Email: ${parsed.email}`);
                if (parsed.documentReference) updates.push(`Reference: ${parsed.documentReference}`);
                if (parsed.displayText) updates.push(`Display: "${parsed.displayText}"`);
                if (parsed.tooltip) updates.push(`Tooltip: "${parsed.tooltip}"`);
                return updates.length > 0 ? `Update hyperlink (${updates.join(", ")})` : "Edit hyperlink";
            } catch {
                return "Edit hyperlink";
            }
        default:
            return action.data || "No details";
    }
}

/**
 * Filters actions based on selection state
 * @param {Object[]} actions - All pending actions
 * @param {boolean[]} selections - Selection state for each action
 * @returns {Object[]} Only the selected actions
 */
function filterSelectedActions(actions, selections) {
    if (!actions || !selections) return [];
    return actions.filter((_, index) => selections[index] === true);
}

/**
 * Checks if any actions are selected
 * @param {boolean[]} selections - Selection state array
 * @returns {boolean} True if at least one action is selected
 */
function hasSelectedActions(selections) {
    if (!selections || selections.length === 0) return false;
    return selections.some(s => s === true);
}

/**
 * Renders a single action preview item as HTML string
 * @param {Object} action - The action to render
 * @param {number} index - Index in the actions array
 * @param {boolean} isExpanded - Whether to show full details
 * @param {boolean} isSelected - Whether the action is selected
 * @param {boolean} hasWarning - Whether to show warning indicator
 * @returns {string} HTML string for the preview item
 */
function renderPreviewItem(action, index, isExpanded, isSelected, hasWarning) {
    const icon = getActionIcon(action.type);
    const summary = getActionSummary(action);
    const details = getActionDetails(action);
    const expandedClass = isExpanded ? 'expanded' : '';
    const warningClass = hasWarning ? 'warning' : '';
    
    return `
        <div class="preview-item ${expandedClass} ${warningClass}" data-index="${index}">
            <input type="checkbox" class="preview-checkbox" ${isSelected ? 'checked' : ''} data-index="${index}">
            <div class="preview-icon ${action.type}">${icon}</div>
            <div class="preview-content">
                <div class="preview-summary">
                    ${summary}
                    <span class="preview-target">${action.target}</span>
                    ${hasWarning ? '<svg class="preview-warning" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2L1 21h22L12 2zm0 3.5L19.5 19h-15L12 5.5zM11 10v4h2v-4h-2zm0 6v2h2v-2h-2z"/></svg>' : ''}
                </div>
                <div class="preview-details">${details}</div>
            </div>
            <div class="preview-expand">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg>
            </div>
        </div>
    `;
}

/**
 * Renders the complete preview panel HTML
 * @param {Object[]} actions - Array of pending actions
 * @param {boolean[]} selections - Selection state for each action
 * @param {number} expandedIndex - Index of expanded action (-1 if none)
 * @param {number[]} warningIndices - Indices of actions with warnings
 * @returns {string} HTML string for all preview items
 */
function renderPreviewList(actions, selections, expandedIndex, warningIndices = []) {
    if (!actions || actions.length === 0) return '';
    
    return actions.map((action, index) => {
        const isExpanded = index === expandedIndex;
        const isSelected = selections[index] !== false;
        const hasWarning = warningIndices.includes(index);
        return renderPreviewItem(action, index, isExpanded, isSelected, hasWarning);
    }).join('');
}

// ES Module exports
export {
    ACTION_TYPES,
    getActionIcon,
    getActionSummary,
    getActionDetails,
    filterSelectedActions,
    hasSelectedActions,
    renderPreviewItem,
    renderPreviewList
};

// CommonJS exports for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        ACTION_TYPES,
        getActionIcon,
        getActionSummary,
        getActionDetails,
        filterSelectedActions,
        hasSelectedActions,
        renderPreviewItem,
        renderPreviewList
    };
}
