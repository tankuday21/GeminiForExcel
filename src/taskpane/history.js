/**
 * History Module for Undo Functionality
 * Manages action history and undo data
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
 * 
 * Note: History functions (createHistoryEntry, addToHistory, removeFromHistory,
 * getLatestEntry, hasHistory) are type-agnostic and work with all action types
 * without requiring type-specific logic.
 */

const MAX_ENTRIES = 20;

/**
 * Human-readable labels for all 87 action types
 * Shared constant to avoid recreation on each renderHistoryEntry call
 */
const TYPE_LABELS = {
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
 * @typedef {Object} HistoryEntry
 * @property {string} id - Unique identifier
 * @property {string} type - Action type
 * @property {string} target - Target range address
 * @property {number} timestamp - Unix timestamp
 * @property {Object} undoData - Data to restore previous state
 */

/**
 * Generates a unique ID for history entries
 * @returns {string} Unique ID
 */
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

/**
 * Creates a new history entry
 * @param {Object} action - The action that was applied
 * @param {Object} undoData - The captured undo data
 * @returns {HistoryEntry} The created history entry
 */
function createHistoryEntry(action, undoData) {
    return {
        id: generateId(),
        type: action.type,
        target: action.target,
        timestamp: Date.now(),
        undoData: undoData
    };
}

/**
 * Adds an entry to the history array (prepends to front)
 * @param {HistoryEntry[]} entries - Current history entries
 * @param {HistoryEntry} entry - Entry to add
 * @param {number} maxEntries - Maximum entries to retain
 * @returns {HistoryEntry[]} Updated history entries
 */
function addToHistory(entries, entry, maxEntries = MAX_ENTRIES) {
    const newEntries = [entry, ...entries];
    // Enforce max limit by removing oldest entries
    if (newEntries.length > maxEntries) {
        return newEntries.slice(0, maxEntries);
    }
    return newEntries;
}

/**
 * Removes the most recent entry from history
 * @param {HistoryEntry[]} entries - Current history entries
 * @returns {{ entries: HistoryEntry[], removed: HistoryEntry|null }} Updated entries and removed entry
 */
function removeFromHistory(entries) {
    if (!entries || entries.length === 0) {
        return { entries: [], removed: null };
    }
    const [removed, ...remaining] = entries;
    return { entries: remaining, removed };
}

/**
 * Gets the most recent entry without removing it
 * @param {HistoryEntry[]} entries - Current history entries
 * @returns {HistoryEntry|null} Most recent entry or null
 */
function getLatestEntry(entries) {
    if (!entries || entries.length === 0) return null;
    return entries[0];
}

/**
 * Checks if history has any entries
 * @param {HistoryEntry[]} entries - Current history entries
 * @returns {boolean} True if history has entries
 */
function hasHistory(entries) {
    if (!entries || entries.length === 0) return false;
    return true;
}

/**
 * Formats a timestamp as relative time
 * @param {number} timestamp - Unix timestamp
 * @returns {string} Formatted relative time (e.g., "2 min ago")
 */
function formatRelativeTime(timestamp) {
    const now = Date.now();
    const diff = now - timestamp;
    
    const seconds = Math.floor(diff / 1000);
    const minutes = Math.floor(seconds / 60);
    const hours = Math.floor(minutes / 60);
    const days = Math.floor(hours / 24);
    
    if (seconds < 60) return 'just now';
    if (minutes < 60) return `${minutes} min ago`;
    if (hours < 24) return `${hours} hr ago`;
    if (days === 1) return 'yesterday';
    return `${days} days ago`;
}

/**
 * Renders a single history entry as HTML
 * @param {HistoryEntry} entry - The entry to render
 * @param {Function} getIcon - Function to get icon for action type
 * @returns {string} HTML string
 */
function renderHistoryEntry(entry, getIcon) {
    const icon = getIcon ? getIcon(entry.type) : '';
    const timeStr = formatRelativeTime(entry.timestamp);
    const label = TYPE_LABELS[entry.type] || entry.type;
    
    return `
        <div class="history-entry" data-id="${entry.id}">
            <div class="history-icon ${entry.type}">${icon}</div>
            <div class="history-content">
                <span class="history-label">${label}</span>
                <span class="history-target">${entry.target}</span>
            </div>
            <span class="history-time">${timeStr}</span>
        </div>
    `;
}

/**
 * Renders the history panel content
 * @param {HistoryEntry[]} entries - History entries to display
 * @param {Function} getIcon - Function to get icon for action type
 * @returns {string} HTML string
 */
function renderHistoryList(entries, getIcon) {
    if (!entries || entries.length === 0) {
        return '<div class="history-empty">No actions yet</div>';
    }
    
    return entries.map(entry => renderHistoryEntry(entry, getIcon)).join('');
}

// ES Module exports
export {
    MAX_ENTRIES,
    generateId,
    createHistoryEntry,
    addToHistory,
    removeFromHistory,
    getLatestEntry,
    hasHistory,
    formatRelativeTime,
    renderHistoryEntry,
    renderHistoryList
};

// CommonJS exports for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        MAX_ENTRIES,
        generateId,
        createHistoryEntry,
        addToHistory,
        removeFromHistory,
        getLatestEntry,
        hasHistory,
        formatRelativeTime,
        renderHistoryEntry,
        renderHistoryList
    };
}
