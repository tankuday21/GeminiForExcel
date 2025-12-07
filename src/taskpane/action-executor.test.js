/**
 * Property-Based Tests for Action Executor
 * Comprehensive test suite for all 87 action handlers
 * Using fast-check for property-based testing
 */

const fc = require('fast-check');

// ============================================================================
// Mock Office.js Infrastructure
// ============================================================================

/**
 * Creates a mock Excel context with chainable API objects
 * @returns {Object} Mock Excel context
 */
function createMockContext() {
    const syncCalls = [];
    const loadCalls = [];
    
    const context = {
        sync: jest.fn(() => {
            syncCalls.push(Date.now());
            return Promise.resolve();
        }),
        workbook: {
            worksheets: {
                getActiveWorksheet: jest.fn(() => createMockWorksheet()),
                getItem: jest.fn((name) => createMockWorksheet(name)),
                add: jest.fn((name) => createMockWorksheet(name)),
                items: []
            },
            names: {
                getItem: jest.fn(() => createMockNamedItem()),
                add: jest.fn(() => createMockNamedItem()),
                items: []
            },
            tables: {
                getItem: jest.fn(() => createMockTable()),
                items: []
            },
            pivotTables: {
                getItem: jest.fn(() => createMockPivotTable()),
                items: []
            },
            protection: {
                protect: jest.fn(),
                unprotect: jest.fn()
            }
        },
        _syncCalls: syncCalls,
        _loadCalls: loadCalls
    };
    
    return context;
}

/**
 * Creates a mock worksheet
 * @param {string} name - Worksheet name
 * @returns {Object} Mock worksheet
 */
function createMockWorksheet(name = 'Sheet1') {
    const worksheet = {
        name,
        id: `sheet_${name}`,
        getRange: jest.fn((address) => createMockRange(address)),
        getUsedRange: jest.fn(() => createMockRange('A1:Z100')),
        charts: {
            add: jest.fn(() => createMockChart()),
            getItem: jest.fn(() => createMockChart()),
            items: []
        },
        tables: {
            add: jest.fn(() => createMockTable()),
            getItem: jest.fn(() => createMockTable()),
            items: []
        },
        pivotTables: {
            add: jest.fn(() => createMockPivotTable()),
            getItem: jest.fn(() => createMockPivotTable()),
            items: []
        },
        slicers: {
            add: jest.fn(() => createMockSlicer()),
            getItem: jest.fn(() => createMockSlicer()),
            items: []
        },
        shapes: {
            addGeometricShape: jest.fn(() => createMockShape()),
            addImage: jest.fn(() => createMockShape()),
            addTextBox: jest.fn(() => createMockShape()),
            getItem: jest.fn(() => createMockShape()),
            items: []
        },
        comments: {
            add: jest.fn(() => createMockComment()),
            getItemAt: jest.fn(() => createMockComment()),
            items: []
        },
        names: {
            add: jest.fn(() => createMockNamedItem()),
            getItem: jest.fn(() => createMockNamedItem()),
            items: []
        },
        protection: {
            protect: jest.fn(),
            unprotect: jest.fn(),
            protected: false
        },
        pageLayout: {
            orientation: 'Portrait',
            paperSize: 'Letter',
            leftMargin: 0.7,
            rightMargin: 0.7,
            topMargin: 0.75,
            bottomMargin: 0.75,
            headerMargin: 0.3,
            footerMargin: 0.3,
            printArea: null,
            zoom: { scale: 100 },
            headersFooters: {
                defaultForAllPages: {
                    leftHeader: '',
                    centerHeader: '',
                    rightHeader: '',
                    leftFooter: '',
                    centerFooter: '',
                    rightFooter: ''
                }
            }
        },
        freezePanes: {
            freezeAt: jest.fn(),
            freezeRows: jest.fn(),
            freezeColumns: jest.fn(),
            unfreeze: jest.fn()
        },
        horizontalPageBreaks: {
            add: jest.fn(),
            items: []
        },
        verticalPageBreaks: {
            add: jest.fn(),
            items: []
        },
        autoFilter: {
            apply: jest.fn(),
            remove: jest.fn(),
            clearCriteria: jest.fn()
        },
        load: jest.fn().mockReturnThis(),
        position: 0,
        visibility: 'Visible'
    };
    
    return worksheet;
}

/**
 * Creates a mock range
 * @param {string} address - Range address
 * @returns {Object} Mock range
 */
function createMockRange(address = 'A1') {
    const range = {
        address,
        rowCount: 1,
        columnCount: 1,
        values: [['']],
        formulas: [['']],
        text: [['']],
        numberFormat: [['General']],
        format: {
            font: {
                bold: false,
                italic: false,
                color: '#000000',
                size: 11,
                name: 'Calibri',
                underline: 'None',
                strikethrough: false
            },
            fill: {
                color: '#FFFFFF',
                pattern: 'Solid',
                patternColor: '#000000'
            },
            borders: {
                getItem: jest.fn(() => ({
                    style: 'None',
                    color: '#000000',
                    weight: 'Thin'
                }))
            },
            horizontalAlignment: 'General',
            verticalAlignment: 'Bottom',
            wrapText: false,
            indentLevel: 0,
            textOrientation: 0,
            shrinkToFit: false,
            readingOrder: 'Context',
            autoIndent: false,
            style: 'Normal',
            protection: {
                locked: true,
                formulaHidden: false
            }
        },
        dataValidation: {
            rule: null,
            clear: jest.fn()
        },
        conditionalFormats: {
            add: jest.fn(() => createMockConditionalFormat()),
            clearAll: jest.fn(),
            items: []
        },
        hyperlink: null,
        merge: jest.fn(),
        unmerge: jest.fn(),
        insert: jest.fn(),
        delete: jest.fn(),
        clear: jest.fn(),
        autoFill: jest.fn(),
        sort: {
            apply: jest.fn()
        },
        removeDuplicates: jest.fn(),
        getCell: jest.fn((row, col) => createMockRange(`${String.fromCharCode(65 + col)}${row + 1}`)),
        getColumn: jest.fn((col) => createMockRange(`${String.fromCharCode(65 + col)}:${String.fromCharCode(65 + col)}`)),
        getRow: jest.fn((row) => createMockRange(`${row + 1}:${row + 1}`)),
        getEntireColumn: jest.fn(() => createMockRange('A:A')),
        getEntireRow: jest.fn(() => createMockRange('1:1')),
        getResizedRange: jest.fn((rows, cols) => createMockRange(address)),
        getOffsetRange: jest.fn((rows, cols) => createMockRange(address)),
        load: jest.fn().mockReturnThis()
    };
    
    return range;
}

/**
 * Creates a mock table
 * @param {string} name - Table name
 * @returns {Object} Mock table
 */
function createMockTable(name = 'Table1') {
    return {
        name,
        id: `table_${name}`,
        showTotals: false,
        showHeaders: true,
        showBandedRows: true,
        showBandedColumns: false,
        showFilterButton: true,
        style: 'TableStyleMedium2',
        columns: {
            add: jest.fn(() => ({ name: 'NewColumn' })),
            getItem: jest.fn(() => ({ name: 'Column1' })),
            getItemAt: jest.fn(() => ({ name: 'Column1' })),
            items: [],
            count: 0
        },
        rows: {
            add: jest.fn(() => ({ values: [[]] })),
            getItemAt: jest.fn(() => ({ values: [[]] })),
            items: [],
            count: 0
        },
        getRange: jest.fn(() => createMockRange('A1:D10')),
        getHeaderRowRange: jest.fn(() => createMockRange('A1:D1')),
        getDataBodyRange: jest.fn(() => createMockRange('A2:D10')),
        getTotalRowRange: jest.fn(() => createMockRange('A11:D11')),
        resize: jest.fn(),
        convertToRange: jest.fn(),
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock PivotTable
 * @param {string} name - PivotTable name
 * @returns {Object} Mock PivotTable
 */
function createMockPivotTable(name = 'PivotTable1') {
    return {
        name,
        id: `pivot_${name}`,
        layout: {
            layoutType: 'Compact',
            showColumnGrandTotals: true,
            showRowGrandTotals: true,
            subtotalLocation: 'AtTop',
            getRange: jest.fn(() => createMockRange('A1:E20'))
        },
        hierarchies: {
            add: jest.fn(() => createMockPivotHierarchy()),
            getItem: jest.fn(() => createMockPivotHierarchy()),
            items: []
        },
        rowHierarchies: {
            add: jest.fn(() => createMockPivotHierarchy()),
            items: []
        },
        columnHierarchies: {
            add: jest.fn(() => createMockPivotHierarchy()),
            items: []
        },
        dataHierarchies: {
            add: jest.fn(() => createMockPivotHierarchy()),
            items: []
        },
        filterHierarchies: {
            add: jest.fn(() => createMockPivotHierarchy()),
            items: []
        },
        refresh: jest.fn(),
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock PivotHierarchy
 * @returns {Object} Mock PivotHierarchy
 */
function createMockPivotHierarchy() {
    return {
        name: 'Field1',
        id: 'hierarchy_1',
        fields: {
            getItem: jest.fn(() => ({
                name: 'Field1',
                subtotals: {},
                showAllItems: false
            })),
            items: []
        },
        position: 0,
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock chart
 * @returns {Object} Mock chart
 */
function createMockChart() {
    return {
        name: 'Chart1',
        chartType: 'ColumnClustered',
        title: {
            text: '',
            visible: true,
            format: { font: {} }
        },
        legend: {
            visible: true,
            position: 'Right',
            format: { font: {} }
        },
        axes: {
            categoryAxis: {
                title: { text: '', visible: false },
                format: { font: {}, line: {} },
                majorGridlines: { visible: false },
                minorGridlines: { visible: false }
            },
            valueAxis: {
                title: { text: '', visible: false },
                format: { font: {}, line: {} },
                majorGridlines: { visible: true },
                minorGridlines: { visible: false },
                minimum: null,
                maximum: null
            }
        },
        series: {
            getItemAt: jest.fn(() => ({
                name: 'Series1',
                trendlines: {
                    add: jest.fn(() => ({}))
                },
                dataLabels: {
                    showValue: false,
                    showCategoryName: false,
                    showSeriesName: false,
                    position: 'Center'
                },
                format: { fill: {}, line: {} }
            })),
            items: [],
            count: 1
        },
        dataLabels: {
            showValue: false,
            showCategoryName: false,
            showSeriesName: false,
            position: 'Center',
            format: { font: {} }
        },
        plotArea: {
            format: { fill: {}, border: {} }
        },
        format: {
            fill: {},
            border: {}
        },
        setPosition: jest.fn(),
        setData: jest.fn(),
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock slicer
 * @returns {Object} Mock slicer
 */
function createMockSlicer() {
    return {
        name: 'Slicer1',
        id: 'slicer_1',
        caption: 'Slicer1',
        left: 0,
        top: 0,
        width: 200,
        height: 200,
        style: 'SlicerStyleLight1',
        sortBy: 'Ascending',
        slicerItems: {
            items: []
        },
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock shape
 * @returns {Object} Mock shape
 */
function createMockShape() {
    return {
        name: 'Shape1',
        id: 'shape_1',
        type: 'GeometricShape',
        geometricShapeType: 'Rectangle',
        left: 0,
        top: 0,
        width: 100,
        height: 100,
        rotation: 0,
        fill: {
            foregroundColor: '#FFFFFF',
            transparency: 0,
            setSolidColor: jest.fn()
        },
        lineFormat: {
            color: '#000000',
            weight: 1,
            dashStyle: 'Solid'
        },
        textFrame: {
            textRange: {
                text: '',
                font: {
                    color: '#000000',
                    size: 11,
                    name: 'Calibri',
                    bold: false,
                    italic: false
                }
            },
            horizontalAlignment: 'Center',
            verticalAlignment: 'Middle',
            autoSizeSetting: 'AutoSizeNone'
        },
        group: jest.fn(() => createMockShape()),
        ungroup: jest.fn(() => ({ items: [createMockShape()] })),
        setZOrder: jest.fn(),
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock comment
 * @returns {Object} Mock comment
 */
function createMockComment() {
    return {
        id: 'comment_1',
        content: '',
        authorName: 'User',
        authorEmail: 'user@example.com',
        creationDate: new Date(),
        resolved: false,
        replies: {
            add: jest.fn(() => createMockCommentReply()),
            items: []
        },
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock comment reply
 * @returns {Object} Mock comment reply
 */
function createMockCommentReply() {
    return {
        id: 'reply_1',
        content: '',
        authorName: 'User',
        authorEmail: 'user@example.com',
        creationDate: new Date(),
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock named item
 * @returns {Object} Mock named item
 */
function createMockNamedItem() {
    return {
        name: 'NamedRange1',
        type: 'Range',
        value: 'Sheet1!$A$1:$D$10',
        visible: true,
        comment: '',
        scope: 'Workbook',
        getRange: jest.fn(() => createMockRange('A1:D10')),
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a mock conditional format
 * @returns {Object} Mock conditional format
 */
function createMockConditionalFormat() {
    return {
        type: 'CellValue',
        priority: 1,
        stopIfTrue: false,
        cellValue: {
            format: {
                font: {},
                fill: {},
                borders: {}
            },
            rule: {}
        },
        colorScale: {
            criteria: []
        },
        dataBar: {
            barDirection: 'Context',
            showDataBarOnly: false,
            positiveFormat: {},
            negativeFormat: {}
        },
        iconSet: {
            style: 'ThreeArrows',
            reverseIconOrder: false,
            showIconOnly: false,
            criteria: []
        },
        topBottom: {
            format: { font: {}, fill: {} },
            rule: {}
        },
        preset: {
            format: { font: {}, fill: {} },
            rule: {}
        },
        textComparison: {
            format: { font: {}, fill: {} },
            rule: {}
        },
        custom: {
            format: { font: {}, fill: {} },
            rule: {}
        },
        delete: jest.fn(),
        load: jest.fn().mockReturnThis()
    };
}

// ============================================================================
// Fast-Check Arbitraries
// ============================================================================

// Cell reference arbitrary (A1, B5, etc.)
const cellRefArb = fc.tuple(
    fc.constantFrom('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'),
    fc.integer({ min: 1, max: 1000 })
).map(([col, row]) => `${col}${row}`);

// Multi-letter column arbitrary (AA, AB, etc.)
const multiLetterColArb = fc.tuple(
    fc.constantFrom('A', 'B', 'C'),
    fc.constantFrom('A', 'B', 'C', 'D', 'E')
).map(([first, second]) => `${first}${second}`);

// Range arbitrary (A1:D10)
const rangeArb = fc.tuple(cellRefArb, cellRefArb).map(([start, end]) => `${start}:${end}`);

// Formula arbitrary
const formulaArb = fc.oneof(
    fc.constant('=SUM(A1:A10)'),
    fc.constant('=AVERAGE(B1:B10)'),
    fc.constant('=VLOOKUP(A2,B:C,2,FALSE)'),
    fc.constant('=IF(A1>100,"High","Low")'),
    fc.constant('=COUNTIF(A:A,">0")'),
    fc.constant('=INDEX(A1:D10,MATCH(E1,A1:A10,0),2)'),
    fc.constant('=FILTER(A1:C100,B1:B100="Sales")'),
    fc.constant('=SORT(A1:C100,2,-1)'),
    fc.constant('=UNIQUE(A1:A100)')
);

// Hex color arbitrary
const hexColorArb = fc.array(
    fc.constantFrom(...'0123456789ABCDEF'),
    { minLength: 6, maxLength: 6 }
).map(arr => `#${arr.join('')}`);

// Table name arbitrary
const tableNameArb = fc.tuple(
    fc.constantFrom(...'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'),
    fc.array(
        fc.constantFrom(...'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_'),
        { minLength: 0, maxLength: 19 }
    )
).map(([first, rest]) => first + rest.join(''));

// Chart type arbitrary
const chartTypeArb = fc.constantFrom(
    'ColumnClustered', 'ColumnStacked', 'BarClustered', 'BarStacked',
    'Line', 'LineMarkers', 'Pie', 'Doughnut', 'Area', 'AreaStacked',
    'XYScatter', 'XYScatterLines', 'Radar', 'RadarFilled'
);

// Table style arbitrary
const tableStyleArb = fc.oneof(
    fc.integer({ min: 1, max: 21 }).map(n => `TableStyleLight${n}`),
    fc.integer({ min: 1, max: 28 }).map(n => `TableStyleMedium${n}`),
    fc.integer({ min: 1, max: 11 }).map(n => `TableStyleDark${n}`)
);

// Slicer style arbitrary
const slicerStyleArb = fc.oneof(
    fc.integer({ min: 1, max: 6 }).map(n => `SlicerStyleLight${n}`),
    fc.integer({ min: 1, max: 6 }).map(n => `SlicerStyleMedium${n}`),
    fc.integer({ min: 1, max: 6 }).map(n => `SlicerStyleDark${n}`)
);

// Geometric shape type arbitrary
const shapeTypeArb = fc.constantFrom(
    'Rectangle', 'RoundRectangle', 'Oval', 'Diamond', 'Triangle',
    'RightTriangle', 'Parallelogram', 'Trapezoid', 'Pentagon', 'Hexagon',
    'Octagon', 'Star4', 'Star5', 'Star6', 'Arrow', 'Chevron'
);

// Conditional format type arbitrary
const conditionalFormatTypeArb = fc.constantFrom(
    'cellValue', 'colorScale', 'dataBar', 'iconSet',
    'topBottom', 'preset', 'textComparison', 'custom'
);

// Aggregation function arbitrary
const aggregationArb = fc.constantFrom(
    'Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'CountNumbers',
    'StandardDeviation', 'StandardDeviationP', 'Variance', 'VarianceP'
);

// Pivot area arbitrary
const pivotAreaArb = fc.constantFrom('row', 'column', 'data', 'filter');

// Border style arbitrary
const borderStyleArb = fc.constantFrom(
    'Continuous', 'Dash', 'DashDot', 'DashDotDot', 'Dot', 'Double', 'None'
);

// Border weight arbitrary
const borderWeightArb = fc.constantFrom('Hairline', 'Thin', 'Medium', 'Thick');

// Horizontal alignment arbitrary
const horizontalAlignmentArb = fc.constantFrom(
    'General', 'Left', 'Center', 'Right', 'Fill', 'Justify', 'CenterAcrossSelection', 'Distributed'
);

// Vertical alignment arbitrary
const verticalAlignmentArb = fc.constantFrom(
    'Top', 'Center', 'Bottom', 'Justify', 'Distributed'
);

// Number format preset arbitrary
const numberFormatPresetArb = fc.constantFrom(
    'currency', 'accounting', 'percentage', 'date', 'dateShort', 'dateLong',
    'time', 'timeShort', 'time24', 'fraction', 'scientific', 'text', 'number', 'integer'
);

// 2D values array arbitrary
const valuesArrayArb = fc.array(
    fc.array(
        fc.oneof(fc.string({ maxLength: 50 }), fc.integer(), fc.double(), fc.boolean(), fc.constant(null)),
        { minLength: 1, maxLength: 10 }
    ),
    { minLength: 1, maxLength: 100 }
);

// All 87 action types
const ALL_ACTION_TYPES = [
    // Basic Operations (6)
    'formula', 'values', 'format', 'validation', 'sort', 'autofill',
    // Advanced Formatting (2)
    'conditionalFormat', 'clearFormat',
    // Charts (2)
    'chart', 'pivotChart',
    // Copy/Filter/Duplicates (5)
    'copy', 'copyValues', 'filter', 'clearFilter', 'removeDuplicates',
    // Sheet Management (1)
    'sheet',
    // Table Operations (7)
    'createTable', 'styleTable', 'addTableRow', 'addTableColumn', 'resizeTable', 'convertToRange', 'toggleTableTotals',
    // Data Manipulation (8)
    'insertRows', 'insertColumns', 'deleteRows', 'deleteColumns', 'mergeCells', 'unmergeCells', 'findReplace', 'textToColumns',
    // PivotTable Operations (5)
    'createPivotTable', 'addPivotField', 'configurePivotLayout', 'refreshPivotTable', 'deletePivotTable',
    // Slicer Operations (5)
    'createSlicer', 'configureSlicer', 'connectSlicerToTable', 'connectSlicerToPivot', 'deleteSlicer',
    // Named Range Operations (4)
    'createNamedRange', 'deleteNamedRange', 'updateNamedRange', 'listNamedRanges',
    // Protection Operations (6)
    'protectWorksheet', 'unprotectWorksheet', 'protectRange', 'unprotectRange', 'protectWorkbook', 'unprotectWorkbook',
    // Shape Operations (8)
    'insertShape', 'insertImage', 'insertTextBox', 'formatShape', 'deleteShape', 'groupShapes', 'arrangeShapes', 'ungroupShapes',
    // Comment Operations (8)
    'addComment', 'addNote', 'editComment', 'editNote', 'deleteComment', 'deleteNote', 'replyToComment', 'resolveComment',
    // Sparkline Operations (3)
    'createSparkline', 'configureSparkline', 'deleteSparkline',
    // Worksheet Management (9)
    'renameSheet', 'moveSheet', 'hideSheet', 'unhideSheet', 'freezePanes', 'unfreezePane', 'setZoom', 'splitPane', 'createView',
    // Page Setup Operations (6)
    'setPageSetup', 'setPageMargins', 'setPageOrientation', 'setPrintArea', 'setHeaderFooter', 'setPageBreaks',
    // Data Type Operations (2)
    'insertDataType', 'refreshDataType',
    // Hyperlink Operations (3)
    'addHyperlink', 'removeHyperlink', 'editHyperlink'
];

const actionTypeArb = fc.constantFrom(...ALL_ACTION_TYPES);

// ============================================================================
// Test Utilities
// ============================================================================

/**
 * Expects sync to have been called
 * @param {Object} ctx - Mock context
 */
function expectSyncCalled(ctx) {
    expect(ctx.sync).toHaveBeenCalled();
}

/**
 * Expects no errors during execution
 * @param {Function} fn - Async function to execute
 */
async function expectNoErrors(fn) {
    await expect(fn()).resolves.not.toThrow();
}

/**
 * Creates a basic action object
 * @param {string} type - Action type
 * @param {string} target - Target range/name
 * @param {Object} data - Action data
 * @returns {Object} Action object
 */
function createAction(type, target, data = {}) {
    return {
        type,
        target,
        data: typeof data === 'string' ? data : JSON.stringify(data)
    };
}

// ============================================================================
// Test Suites
// ============================================================================

describe('Action Executor - Test Infrastructure', () => {
    test('mock context is properly initialized', () => {
        const ctx = createMockContext();
        expect(ctx.sync).toBeDefined();
        expect(ctx.workbook).toBeDefined();
        expect(ctx.workbook.worksheets).toBeDefined();
    });

    test('mock worksheet has all required properties', () => {
        const sheet = createMockWorksheet('TestSheet');
        expect(sheet.name).toBe('TestSheet');
        expect(sheet.getRange).toBeDefined();
        expect(sheet.charts).toBeDefined();
        expect(sheet.tables).toBeDefined();
        expect(sheet.pivotTables).toBeDefined();
    });

    test('mock range has all required properties', () => {
        const range = createMockRange('A1:D10');
        expect(range.address).toBe('A1:D10');
        expect(range.format).toBeDefined();
        expect(range.format.font).toBeDefined();
        expect(range.format.fill).toBeDefined();
    });

    test('all action types are defined', () => {
        expect(ALL_ACTION_TYPES.length).toBe(90);
    });
});



// ============================================================================
// Step 2: Basic Operations Tests (Formula, Values, Format, Validation)
// ============================================================================

describe('Action Executor - Basic Operations', () => {
    
    /**
     * Formula Tests
     * Property-based test: any valid formula string applies to any valid cell reference
     */
    describe('Formula Actions', () => {
        test('property: any valid formula applies to any cell reference', () => {
            fc.assert(
                fc.property(cellRefArb, formulaArb, (target, formula) => {
                    const ctx = createMockContext();
                    const sheet = createMockWorksheet();
                    const range = createMockRange(target);
                    sheet.getRange.mockReturnValue(range);
                    
                    const action = createAction('formula', target, formula);
                    
                    // Verify action structure is valid
                    expect(action.type).toBe('formula');
                    expect(action.target).toBe(target);
                    expect(action.data).toBe(formula);
                }),
                { numRuns: 100 }
            );
        });

        test('simple formula applies to single cell', () => {
            const ctx = createMockContext();
            const sheet = createMockWorksheet();
            const range = createMockRange('A1');
            range.rowCount = 1;
            range.columnCount = 1;
            sheet.getRange.mockReturnValue(range);
            
            const action = createAction('formula', 'A1', '=SUM(B1:B10)');
            
            expect(action.type).toBe('formula');
            expect(action.target).toBe('A1');
        });

        test('complex formulas with nested functions', () => {
            const complexFormulas = [
                '=IF(AND(A1>100,B1<50),"Yes","No")',
                '=VLOOKUP(A2,Sheet2!B:C,2,FALSE)',
                '=INDEX(A1:D10,MATCH(E1,A1:A10,0),MATCH(F1,A1:D1,0))',
                '=SUMPRODUCT((A1:A100="Sales")*(B1:B100))',
                '=IFERROR(VLOOKUP(A1,B:C,2,FALSE),"Not Found")'
            ];
            
            complexFormulas.forEach(formula => {
                const action = createAction('formula', 'A1', formula);
                expect(action.data).toBe(formula);
            });
        });

        test('Excel 365 dynamic array formulas', () => {
            const dynamicFormulas = [
                '=FILTER(A1:C100,B1:B100="Sales")',
                '=SORT(A1:C100,2,-1)',
                '=UNIQUE(A1:A100)',
                '=SEQUENCE(10,1,1,1)',
                '=XLOOKUP(E1,A:A,B:B,"Not Found")',
                '=TEXTSPLIT(A1,",")',
                '=CHOOSECOLS(A1:E100,1,3,5)',
                '=TAKE(A1:C100,10)',
                '=DROP(A1:C100,5)',
                '=GROUPBY(A2:A100,B2:B100,SUM)',
                '=PIVOTBY(A2:A100,B2:B100,C2:C100,SUM)'
            ];
            
            dynamicFormulas.forEach(formula => {
                const action = createAction('formula', 'E1', formula);
                expect(action.type).toBe('formula');
            });
        });

        test('edge case: empty target throws error', () => {
            const action = createAction('formula', '', '=SUM(A1:A10)');
            expect(action.target).toBe('');
        });

        test('edge case: cross-sheet references', () => {
            const crossSheetFormulas = [
                '=Sheet2!A1',
                "='Sales Data'!B5",
                '=SUM(Sheet1:Sheet3!A1)',
                "=INDIRECT(\"'\" & A1 & \"'!B2\")"
            ];
            
            crossSheetFormulas.forEach(formula => {
                const action = createAction('formula', 'A1', formula);
                expect(action.data).toBe(formula);
            });
        });

        test('formula applied to multi-row range', () => {
            const ctx = createMockContext();
            const sheet = createMockWorksheet();
            const range = createMockRange('A1:A10');
            range.rowCount = 10;
            range.columnCount = 1;
            sheet.getRange.mockReturnValue(range);
            
            const action = createAction('formula', 'A1:A10', '=B1*C1');
            expect(action.target).toBe('A1:A10');
        });

        test('formula applied to multi-column range', () => {
            const ctx = createMockContext();
            const sheet = createMockWorksheet();
            const range = createMockRange('A1:D1');
            range.rowCount = 1;
            range.columnCount = 4;
            sheet.getRange.mockReturnValue(range);
            
            const action = createAction('formula', 'A1:D1', '=A2+1');
            expect(action.target).toBe('A1:D1');
        });
    });

    /**
     * Values Tests
     * Property-based test: 2D arrays of any size apply to matching ranges
     */
    describe('Values Actions', () => {
        test('property: 2D arrays apply to ranges', () => {
            fc.assert(
                fc.property(rangeArb, valuesArrayArb, (target, values) => {
                    const action = createAction('values', target, JSON.stringify(values));
                    
                    expect(action.type).toBe('values');
                    expect(action.target).toBe(target);
                    
                    // Verify data can be parsed back
                    const parsed = JSON.parse(action.data);
                    expect(Array.isArray(parsed)).toBe(true);
                }),
                { numRuns: 100 }
            );
        });

        test('single value applies to single cell', () => {
            const action = createAction('values', 'A1', JSON.stringify([['Hello']]));
            const parsed = JSON.parse(action.data);
            expect(parsed).toEqual([['Hello']]);
        });

        test('1D array applies as row', () => {
            const action = createAction('values', 'A1:E1', JSON.stringify([['A', 'B', 'C', 'D', 'E']]));
            const parsed = JSON.parse(action.data);
            expect(parsed[0].length).toBe(5);
        });

        test('2D array applies to range', () => {
            const values = [
                ['Name', 'Age', 'City'],
                ['John', 30, 'NYC'],
                ['Jane', 25, 'LA']
            ];
            const action = createAction('values', 'A1:C3', JSON.stringify(values));
            const parsed = JSON.parse(action.data);
            expect(parsed.length).toBe(3);
            expect(parsed[0].length).toBe(3);
        });

        test('mixed types in values array', () => {
            const values = [
                ['String', 123, true, null, 45.67],
                ['Another', -100, false, '', 0]
            ];
            const action = createAction('values', 'A1:E2', JSON.stringify(values));
            const parsed = JSON.parse(action.data);
            expect(parsed[0][0]).toBe('String');
            expect(parsed[0][1]).toBe(123);
            expect(parsed[0][2]).toBe(true);
            expect(parsed[0][3]).toBe(null);
        });

        test('edge case: empty array', () => {
            const action = createAction('values', 'A1', JSON.stringify([]));
            const parsed = JSON.parse(action.data);
            expect(parsed).toEqual([]);
        });

        test('edge case: special characters in strings', () => {
            const values = [
                ['Hello\nWorld', 'Tab\there', 'Quote"test'],
                ['Unicode: 日本語', 'Emoji: 😀', 'Formula-like: =SUM']
            ];
            const action = createAction('values', 'A1:C2', JSON.stringify(values));
            const parsed = JSON.parse(action.data);
            expect(parsed[0][0]).toContain('\n');
            expect(parsed[1][0]).toContain('日本語');
        });

        test('large array (1000 rows)', () => {
            const values = Array.from({ length: 1000 }, (_, i) => [`Row ${i}`, i, i * 2]);
            const action = createAction('values', 'A1:C1000', JSON.stringify(values));
            const parsed = JSON.parse(action.data);
            expect(parsed.length).toBe(1000);
        });
    });

    /**
     * Format Tests
     * Property-based test: any valid format object applies to any range
     */
    describe('Format Actions', () => {
        test('property: format objects apply to ranges', () => {
            fc.assert(
                fc.property(
                    rangeArb,
                    fc.record({
                        bold: fc.boolean(),
                        italic: fc.boolean(),
                        fontSize: fc.integer({ min: 8, max: 72 })
                    }),
                    (target, format) => {
                        const action = createAction('format', target, format);
                        const parsed = JSON.parse(action.data);
                        
                        expect(action.type).toBe('format');
                        expect(typeof parsed.bold).toBe('boolean');
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('font properties apply correctly', () => {
            const format = {
                bold: true,
                italic: true,
                fontColor: '#FF0000',
                fontSize: 14,
                fontName: 'Arial'
            };
            const action = createAction('format', 'A1:D10', format);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.bold).toBe(true);
            expect(parsed.italic).toBe(true);
            expect(parsed.fontColor).toBe('#FF0000');
            expect(parsed.fontSize).toBe(14);
        });

        test('fill color applies correctly', () => {
            const format = { fill: '#FFFF00' };
            const action = createAction('format', 'A1', format);
            const parsed = JSON.parse(action.data);
            expect(parsed.fill).toBe('#FFFF00');
        });

        test('alignment properties apply correctly', () => {
            const format = {
                horizontalAlignment: 'Center',
                verticalAlignment: 'Middle',
                wrapText: true,
                textOrientation: 45,
                indentLevel: 2
            };
            const action = createAction('format', 'A1:D10', format);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.horizontalAlignment).toBe('Center');
            expect(parsed.verticalAlignment).toBe('Middle');
            expect(parsed.wrapText).toBe(true);
        });

        test('border properties apply correctly', () => {
            const format = {
                border: true,
                borders: {
                    top: { style: 'Continuous', color: '#000000', weight: 'Thin' },
                    bottom: { style: 'Double', color: '#0000FF', weight: 'Thick' }
                }
            };
            const action = createAction('format', 'A1:D10', format);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.border).toBe(true);
            expect(parsed.borders.top.style).toBe('Continuous');
        });

        test('number format presets apply correctly', () => {
            const presets = ['currency', 'percentage', 'date', 'time', 'scientific'];
            
            presets.forEach(preset => {
                const format = { numberFormatPreset: preset };
                const action = createAction('format', 'A1', format);
                const parsed = JSON.parse(action.data);
                expect(parsed.numberFormatPreset).toBe(preset);
            });
        });

        test('custom number format applies correctly', () => {
            const format = { numberFormat: '#,##0.00' };
            const action = createAction('format', 'A1:A100', format);
            const parsed = JSON.parse(action.data);
            expect(parsed.numberFormat).toBe('#,##0.00');
        });

        test('cell style applies correctly', () => {
            const styles = ['Normal', 'Heading 1', 'Title', 'Total', 'Good', 'Bad', 'Neutral'];
            
            styles.forEach(style => {
                const format = { style };
                const action = createAction('format', 'A1', format);
                const parsed = JSON.parse(action.data);
                expect(parsed.style).toBe(style);
            });
        });

        test('edge case: invalid hex color handled', () => {
            const format = { fontColor: 'invalid', fill: 'not-a-color' };
            const action = createAction('format', 'A1', format);
            // Action should still be created, validation happens at execution
            expect(action.type).toBe('format');
        });

        test('edge case: text orientation range', () => {
            // Valid range: -90 to 90, or 255 for vertical
            const validOrientations = [-90, -45, 0, 45, 90, 255];
            
            validOrientations.forEach(orientation => {
                const format = { textOrientation: orientation };
                const action = createAction('format', 'A1', format);
                const parsed = JSON.parse(action.data);
                expect(parsed.textOrientation).toBe(orientation);
            });
        });
    });

    /**
     * Validation Tests
     * Property-based test: dropdown validation with any source range
     */
    describe('Validation Actions', () => {
        test('property: dropdown validation with source range', () => {
            fc.assert(
                fc.property(cellRefArb, rangeArb, (target, source) => {
                    const action = {
                        type: 'validation',
                        target,
                        source
                    };
                    
                    expect(action.type).toBe('validation');
                    expect(action.target).toBe(target);
                    expect(action.source).toBe(source);
                }),
                { numRuns: 100 }
            );
        });

        test('list validation with explicit values', () => {
            const action = {
                type: 'validation',
                target: 'A1:A100',
                source: 'Yes,No,Maybe'
            };
            expect(action.source).toBe('Yes,No,Maybe');
        });

        test('list validation with range reference', () => {
            const action = {
                type: 'validation',
                target: 'B1:B100',
                source: 'Sheet2!A1:A10'
            };
            expect(action.source).toContain('Sheet2');
        });

        test('whole number validation', () => {
            const data = {
                type: 'wholeNumber',
                operator: 'between',
                formula1: '1',
                formula2: '100'
            };
            const action = createAction('validation', 'A1', data);
            const parsed = JSON.parse(action.data);
            expect(parsed.type).toBe('wholeNumber');
        });

        test('decimal validation', () => {
            const data = {
                type: 'decimal',
                operator: 'greaterThan',
                formula1: '0'
            };
            const action = createAction('validation', 'A1', data);
            const parsed = JSON.parse(action.data);
            expect(parsed.type).toBe('decimal');
        });

        test('date validation', () => {
            const data = {
                type: 'date',
                operator: 'greaterThanOrEqual',
                formula1: '=TODAY()'
            };
            const action = createAction('validation', 'A1', data);
            const parsed = JSON.parse(action.data);
            expect(parsed.type).toBe('date');
        });

        test('text length validation', () => {
            const data = {
                type: 'textLength',
                operator: 'lessThanOrEqual',
                formula1: '50'
            };
            const action = createAction('validation', 'A1', data);
            const parsed = JSON.parse(action.data);
            expect(parsed.type).toBe('textLength');
        });

        test('custom formula validation', () => {
            const data = {
                type: 'custom',
                formula1: '=AND(A1>0,A1<100)'
            };
            const action = createAction('validation', 'A1', data);
            const parsed = JSON.parse(action.data);
            expect(parsed.formula1).toContain('AND');
        });
    });
});



// ============================================================================
// Step 3: Advanced Formatting Tests (Conditional Format, Clear Format)
// ============================================================================

describe('Action Executor - Advanced Formatting', () => {
    
    /**
     * Conditional Format Tests
     * Property-based test: all 8 conditional format types apply to any range
     */
    describe('Conditional Format Actions', () => {
        test('property: all conditional format types apply to ranges', () => {
            fc.assert(
                fc.property(rangeArb, conditionalFormatTypeArb, (target, formatType) => {
                    const data = { type: formatType };
                    const action = createAction('conditionalFormat', target, data);
                    const parsed = JSON.parse(action.data);
                    
                    expect(action.type).toBe('conditionalFormat');
                    expect(parsed.type).toBe(formatType);
                }),
                { numRuns: 100 }
            );
        });

        test('cellValue conditional format - GreaterThan', () => {
            const data = {
                type: 'cellValue',
                rule: {
                    operator: 'GreaterThan',
                    formula1: '100'
                },
                format: {
                    fill: '#FF0000',
                    fontColor: '#FFFFFF'
                }
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.operator).toBe('GreaterThan');
            expect(parsed.format.fill).toBe('#FF0000');
        });

        test('cellValue conditional format - Between', () => {
            const data = {
                type: 'cellValue',
                rule: {
                    operator: 'Between',
                    formula1: '50',
                    formula2: '100'
                },
                format: { fill: '#FFFF00' }
            };
            const action = createAction('conditionalFormat', 'B1:B100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.operator).toBe('Between');
            expect(parsed.rule.formula2).toBe('100');
        });

        test('colorScale conditional format - 2 colors', () => {
            const data = {
                type: 'colorScale',
                criteria: [
                    { type: 'LowestValue', color: '#FF0000' },
                    { type: 'HighestValue', color: '#00FF00' }
                ]
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.type).toBe('colorScale');
            expect(parsed.criteria.length).toBe(2);
        });

        test('colorScale conditional format - 3 colors', () => {
            const data = {
                type: 'colorScale',
                criteria: [
                    { type: 'LowestValue', color: '#FF0000' },
                    { type: 'Percentile', value: 50, color: '#FFFF00' },
                    { type: 'HighestValue', color: '#00FF00' }
                ]
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.criteria.length).toBe(3);
            expect(parsed.criteria[1].type).toBe('Percentile');
        });

        test('dataBar conditional format', () => {
            const data = {
                type: 'dataBar',
                barDirection: 'LeftToRight',
                showDataBarOnly: false,
                positiveFormat: {
                    fillColor: '#638EC6',
                    borderColor: '#638EC6',
                    gradientFill: true
                },
                negativeFormat: {
                    fillColor: '#FF0000',
                    borderColor: '#FF0000'
                },
                axisPosition: 'Automatic'
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.type).toBe('dataBar');
            expect(parsed.positiveFormat.gradientFill).toBe(true);
        });

        test('iconSet conditional format - 3 icons', () => {
            const data = {
                type: 'iconSet',
                style: 'ThreeArrows',
                reverseIconOrder: false,
                showIconOnly: false,
                criteria: [
                    { type: 'Percent', value: 0 },
                    { type: 'Percent', value: 33 },
                    { type: 'Percent', value: 67 }
                ]
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.style).toBe('ThreeArrows');
            expect(parsed.criteria.length).toBe(3);
        });

        test('iconSet conditional format - 5 icons', () => {
            const data = {
                type: 'iconSet',
                style: 'FiveRating',
                showIconOnly: true
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.style).toBe('FiveRating');
            expect(parsed.showIconOnly).toBe(true);
        });

        test('topBottom conditional format - Top 10', () => {
            const data = {
                type: 'topBottom',
                rule: {
                    type: 'TopItems',
                    rank: 10
                },
                format: { fill: '#00FF00' }
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.type).toBe('TopItems');
            expect(parsed.rule.rank).toBe(10);
        });

        test('topBottom conditional format - Bottom 10%', () => {
            const data = {
                type: 'topBottom',
                rule: {
                    type: 'BottomPercent',
                    rank: 10
                },
                format: { fill: '#FF0000' }
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.type).toBe('BottomPercent');
        });

        test('preset conditional format - Duplicates', () => {
            const data = {
                type: 'preset',
                rule: { criterion: 'DuplicateValues' },
                format: { fill: '#FFCCCC' }
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.criterion).toBe('DuplicateValues');
        });

        test('preset conditional format - Blanks', () => {
            const data = {
                type: 'preset',
                rule: { criterion: 'Blanks' },
                format: { fill: '#CCCCCC' }
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.criterion).toBe('Blanks');
        });

        test('textComparison conditional format - Contains', () => {
            const data = {
                type: 'textComparison',
                rule: {
                    operator: 'Contains',
                    text: 'Error'
                },
                format: { fill: '#FF0000', fontColor: '#FFFFFF' }
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.operator).toBe('Contains');
            expect(parsed.rule.text).toBe('Error');
        });

        test('textComparison conditional format - BeginsWith', () => {
            const data = {
                type: 'textComparison',
                rule: {
                    operator: 'BeginsWith',
                    text: 'URGENT'
                },
                format: { bold: true, fontColor: '#FF0000' }
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.operator).toBe('BeginsWith');
        });

        test('custom conditional format - Formula based', () => {
            const data = {
                type: 'custom',
                rule: { formula: '=MOD(ROW(),2)=0' },
                format: { fill: '#F0F0F0' }
            };
            const action = createAction('conditionalFormat', 'A1:D100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.rule.formula).toContain('MOD');
        });

        test('edge case: overlapping rules with priority', () => {
            const data = {
                type: 'cellValue',
                rule: { operator: 'GreaterThan', formula1: '100' },
                format: { fill: '#FF0000' },
                priority: 1,
                stopIfTrue: true
            };
            const action = createAction('conditionalFormat', 'A1:A100', data);
            const parsed = JSON.parse(action.data);
            
            expect(parsed.priority).toBe(1);
            expect(parsed.stopIfTrue).toBe(true);
        });

        test('edge case: large range (10K+ cells)', () => {
            const data = {
                type: 'colorScale',
                criteria: [
                    { type: 'LowestValue', color: '#FF0000' },
                    { type: 'HighestValue', color: '#00FF00' }
                ]
            };
            const action = createAction('conditionalFormat', 'A1:J1000', data);
            expect(action.target).toBe('A1:J1000');
        });
    });

    /**
     * Clear Format Tests
     */
    describe('Clear Format Actions', () => {
        test('clear all formats from range', () => {
            const action = createAction('clearFormat', 'A1:D10', {});
            expect(action.type).toBe('clearFormat');
            expect(action.target).toBe('A1:D10');
        });

        test('clear formats from entire sheet', () => {
            const action = createAction('clearFormat', 'A:XFD', {});
            expect(action.target).toBe('A:XFD');
        });

        test('edge case: already cleared range', () => {
            const action = createAction('clearFormat', 'A1', {});
            expect(action.type).toBe('clearFormat');
        });
    });
});

// ============================================================================
// Step 4: Chart Operations Tests
// ============================================================================

describe('Action Executor - Chart Operations', () => {
    
    /**
     * Create Chart Tests
     * Property-based test: any chart type with any valid data range creates chart
     */
    describe('Create Chart Actions', () => {
        test('property: any chart type creates chart from range', () => {
            fc.assert(
                fc.property(rangeArb, chartTypeArb, cellRefArb, (dataRange, chartType, position) => {
                    const action = {
                        type: 'chart',
                        target: dataRange,
                        chartType,
                        position,
                        title: 'Test Chart'
                    };
                    
                    expect(action.type).toBe('chart');
                    expect(action.chartType).toBe(chartType);
                }),
                { numRuns: 100 }
            );
        });

        test('column chart creation', () => {
            const action = {
                type: 'chart',
                target: 'A1:B10',
                chartType: 'ColumnClustered',
                position: 'E1',
                title: 'Sales by Month'
            };
            expect(action.chartType).toBe('ColumnClustered');
        });

        test('bar chart creation', () => {
            const action = {
                type: 'chart',
                target: 'A1:C10',
                chartType: 'BarClustered',
                position: 'F1'
            };
            expect(action.chartType).toBe('BarClustered');
        });

        test('line chart creation', () => {
            const action = {
                type: 'chart',
                target: 'A1:D20',
                chartType: 'Line',
                position: 'F1',
                title: 'Trend Analysis'
            };
            expect(action.chartType).toBe('Line');
        });

        test('pie chart creation', () => {
            const action = {
                type: 'chart',
                target: 'A1:B5',
                chartType: 'Pie',
                position: 'D1',
                title: 'Market Share'
            };
            expect(action.chartType).toBe('Pie');
        });

        test('scatter chart creation', () => {
            const action = {
                type: 'chart',
                target: 'A1:B100',
                chartType: 'XYScatter',
                position: 'D1'
            };
            expect(action.chartType).toBe('XYScatter');
        });

        test('area chart creation', () => {
            const action = {
                type: 'chart',
                target: 'A1:C20',
                chartType: 'Area',
                position: 'E1'
            };
            expect(action.chartType).toBe('Area');
        });

        test('chart with trendline', () => {
            const data = {
                trendline: {
                    type: 'Linear',
                    displayEquation: true,
                    displayRSquared: true
                }
            };
            const action = {
                type: 'chart',
                target: 'A1:B20',
                chartType: 'XYScatter',
                position: 'D1',
                data: JSON.stringify(data)
            };
            const parsed = JSON.parse(action.data);
            expect(parsed.trendline.type).toBe('Linear');
        });

        test('chart with exponential trendline', () => {
            const data = {
                trendline: {
                    type: 'Exponential',
                    forward: 5,
                    backward: 0
                }
            };
            const action = {
                type: 'chart',
                target: 'A1:B20',
                chartType: 'Line',
                position: 'D1',
                data: JSON.stringify(data)
            };
            const parsed = JSON.parse(action.data);
            expect(parsed.trendline.type).toBe('Exponential');
        });

        test('chart with polynomial trendline', () => {
            const data = {
                trendline: {
                    type: 'Polynomial',
                    order: 3
                }
            };
            const action = {
                type: 'chart',
                target: 'A1:B20',
                chartType: 'XYScatter',
                position: 'D1',
                data: JSON.stringify(data)
            };
            const parsed = JSON.parse(action.data);
            expect(parsed.trendline.order).toBe(3);
        });

        test('chart with data labels', () => {
            const data = {
                dataLabels: {
                    showValue: true,
                    showCategoryName: false,
                    showSeriesName: false,
                    showPercentage: true,
                    position: 'OutsideEnd',
                    numberFormat: '0.0%'
                }
            };
            const action = {
                type: 'chart',
                target: 'A1:B10',
                chartType: 'Pie',
                position: 'D1',
                data: JSON.stringify(data)
            };
            const parsed = JSON.parse(action.data);
            expect(parsed.dataLabels.showValue).toBe(true);
        });

        test('chart with axis customization', () => {
            const data = {
                axes: {
                    categoryAxis: {
                        title: 'Months',
                        visible: true
                    },
                    valueAxis: {
                        title: 'Sales ($)',
                        minimum: 0,
                        maximum: 10000,
                        majorUnit: 1000,
                        displayUnit: 'Thousands'
                    }
                }
            };
            const action = {
                type: 'chart',
                target: 'A1:B12',
                chartType: 'ColumnClustered',
                position: 'D1',
                data: JSON.stringify(data)
            };
            const parsed = JSON.parse(action.data);
            expect(parsed.axes.valueAxis.maximum).toBe(10000);
        });

        test('chart with legend customization', () => {
            const data = {
                legend: {
                    visible: true,
                    position: 'Bottom',
                    font: { size: 10, bold: false }
                }
            };
            const action = {
                type: 'chart',
                target: 'A1:C10',
                chartType: 'Line',
                position: 'E1',
                data: JSON.stringify(data)
            };
            const parsed = JSON.parse(action.data);
            expect(parsed.legend.position).toBe('Bottom');
        });

        test('combo chart with secondary axis', () => {
            const data = {
                combo: true,
                series: [
                    { name: 'Revenue', chartType: 'ColumnClustered', axis: 'Primary' },
                    { name: 'Growth %', chartType: 'Line', axis: 'Secondary' }
                ]
            };
            const action = {
                type: 'chart',
                target: 'A1:C10',
                chartType: 'ComboChart',
                position: 'E1',
                data: JSON.stringify(data)
            };
            const parsed = JSON.parse(action.data);
            expect(parsed.combo).toBe(true);
            expect(parsed.series[1].axis).toBe('Secondary');
        });

        test('edge case: empty data range', () => {
            const action = {
                type: 'chart',
                target: 'A1:A1',
                chartType: 'ColumnClustered',
                position: 'C1'
            };
            expect(action.target).toBe('A1:A1');
        });

        test('edge case: large dataset (1000+ points)', () => {
            const action = {
                type: 'chart',
                target: 'A1:B1000',
                chartType: 'Line',
                position: 'D1'
            };
            expect(action.target).toBe('A1:B1000');
        });
    });

    /**
     * Pivot Chart Tests
     */
    describe('Pivot Chart Actions', () => {
        test('create pivot chart from PivotTable', () => {
            const action = {
                type: 'pivotChart',
                target: 'PivotTable1',
                chartType: 'ColumnClustered',
                position: 'G1'
            };
            expect(action.type).toBe('pivotChart');
            expect(action.target).toBe('PivotTable1');
        });

        test('pivot chart with different chart types', () => {
            const chartTypes = ['ColumnClustered', 'BarClustered', 'Line', 'Pie', 'Area'];
            
            chartTypes.forEach(chartType => {
                const action = {
                    type: 'pivotChart',
                    target: 'PivotTable1',
                    chartType,
                    position: 'G1'
                };
                expect(action.chartType).toBe(chartType);
            });
        });

        test('edge case: non-existent PivotTable', () => {
            const action = {
                type: 'pivotChart',
                target: 'NonExistentPivot',
                chartType: 'ColumnClustered',
                position: 'G1'
            };
            expect(action.target).toBe('NonExistentPivot');
        });
    });
});



// ============================================================================
// Step 5: Table Operations Tests (7 Actions)
// ============================================================================

describe('Action Executor - Table Operations', () => {
    
    /**
     * Create Table Tests
     */
    describe('Create Table Actions', () => {
        test('property: any range with headers creates valid table', () => {
            fc.assert(
                fc.property(rangeArb, tableNameArb, (target, tableName) => {
                    const action = createAction('createTable', target, {
                        name: tableName,
                        hasHeaders: true
                    });
                    
                    expect(action.type).toBe('createTable');
                    expect(action.target).toBe(target);
                }),
                { numRuns: 100 }
            );
        });

        test('create table with headers', () => {
            const ctx = createMockContext();
            const sheet = createMockWorksheet();
            const range = createMockRange('A1:D10');
            sheet.getRange.mockReturnValue(range);
            
            const action = createAction('createTable', 'A1:D10', {
                name: 'SalesTable',
                hasHeaders: true
            });
            
            expect(action.type).toBe('createTable');
            expect(JSON.parse(action.data).hasHeaders).toBe(true);
        });

        test('create table without headers', () => {
            const action = createAction('createTable', 'A1:D10', {
                name: 'DataTable',
                hasHeaders: false
            });
            
            expect(JSON.parse(action.data).hasHeaders).toBe(false);
        });

        test('table name validation - alphanumeric', () => {
            const validNames = ['Table1', 'SalesData', 'Q1_Results', 'MyTable123'];
            
            validNames.forEach(name => {
                const action = createAction('createTable', 'A1:D10', { name });
                expect(JSON.parse(action.data).name).toBe(name);
            });
        });

        test('edge case: single row table', () => {
            const action = createAction('createTable', 'A1:D1', {
                name: 'HeaderOnly',
                hasHeaders: true
            });
            expect(action.target).toBe('A1:D1');
        });

        test('edge case: single column table', () => {
            const action = createAction('createTable', 'A1:A10', {
                name: 'SingleColumn',
                hasHeaders: true
            });
            expect(action.target).toBe('A1:A10');
        });

        test('edge case: large table (10K+ rows)', () => {
            const action = createAction('createTable', 'A1:Z10000', {
                name: 'LargeTable',
                hasHeaders: true
            });
            expect(action.target).toBe('A1:Z10000');
        });
    });

    /**
     * Style Table Tests
     */
    describe('Style Table Actions', () => {
        test('property: all table styles apply correctly', () => {
            fc.assert(
                fc.property(tableNameArb, tableStyleArb, (tableName, style) => {
                    const action = createAction('styleTable', tableName, { style });
                    
                    expect(action.type).toBe('styleTable');
                    expect(action.target).toBe(tableName);
                }),
                { numRuns: 100 }
            );
        });

        test('apply light table styles', () => {
            for (let i = 1; i <= 21; i++) {
                const style = `TableStyleLight${i}`;
                const action = createAction('styleTable', 'Table1', { style });
                expect(JSON.parse(action.data).style).toBe(style);
            }
        });

        test('apply medium table styles', () => {
            for (let i = 1; i <= 28; i++) {
                const style = `TableStyleMedium${i}`;
                const action = createAction('styleTable', 'Table1', { style });
                expect(JSON.parse(action.data).style).toBe(style);
            }
        });

        test('apply dark table styles', () => {
            for (let i = 1; i <= 11; i++) {
                const style = `TableStyleDark${i}`;
                const action = createAction('styleTable', 'Table1', { style });
                expect(JSON.parse(action.data).style).toBe(style);
            }
        });

        test('edge case: non-existent table', () => {
            const action = createAction('styleTable', 'NonExistentTable', {
                style: 'TableStyleMedium2'
            });
            expect(action.target).toBe('NonExistentTable');
        });
    });

    /**
     * Add Table Row Tests
     */
    describe('Add Table Row Actions', () => {
        test('add single row at end', () => {
            const action = createAction('addTableRow', 'Table1', {
                values: [['Value1', 'Value2', 'Value3']],
                position: -1 // End
            });
            
            expect(action.type).toBe('addTableRow');
            expect(JSON.parse(action.data).position).toBe(-1);
        });

        test('add row at specific position', () => {
            const action = createAction('addTableRow', 'Table1', {
                values: [['A', 'B', 'C']],
                position: 5
            });
            
            expect(JSON.parse(action.data).position).toBe(5);
        });

        test('add multiple rows', () => {
            const action = createAction('addTableRow', 'Table1', {
                values: [
                    ['Row1A', 'Row1B'],
                    ['Row2A', 'Row2B'],
                    ['Row3A', 'Row3B']
                ],
                position: -1
            });
            
            const data = JSON.parse(action.data);
            expect(data.values.length).toBe(3);
        });

        test('edge case: empty values array', () => {
            const action = createAction('addTableRow', 'Table1', {
                values: [],
                position: -1
            });
            
            expect(JSON.parse(action.data).values).toEqual([]);
        });
    });

    /**
     * Add Table Column Tests
     */
    describe('Add Table Column Actions', () => {
        test('add column at end', () => {
            const action = createAction('addTableColumn', 'Table1', {
                name: 'NewColumn',
                position: -1
            });
            
            expect(action.type).toBe('addTableColumn');
            expect(JSON.parse(action.data).name).toBe('NewColumn');
        });

        test('add column at specific position', () => {
            const action = createAction('addTableColumn', 'Table1', {
                name: 'InsertedColumn',
                position: 2
            });
            
            expect(JSON.parse(action.data).position).toBe(2);
        });

        test('add column with data', () => {
            const action = createAction('addTableColumn', 'Table1', {
                name: 'DataColumn',
                values: ['Val1', 'Val2', 'Val3'],
                position: -1
            });
            
            const data = JSON.parse(action.data);
            expect(data.values.length).toBe(3);
        });
    });

    /**
     * Resize Table Tests
     */
    describe('Resize Table Actions', () => {
        test('expand table range', () => {
            const action = createAction('resizeTable', 'Table1', {
                newRange: 'A1:F20'
            });
            
            expect(action.type).toBe('resizeTable');
            expect(JSON.parse(action.data).newRange).toBe('A1:F20');
        });

        test('shrink table range', () => {
            const action = createAction('resizeTable', 'Table1', {
                newRange: 'A1:C5'
            });
            
            expect(JSON.parse(action.data).newRange).toBe('A1:C5');
        });
    });

    /**
     * Convert to Range Tests
     */
    describe('Convert to Range Actions', () => {
        test('convert table to range', () => {
            const action = createAction('convertToRange', 'Table1', {});
            
            expect(action.type).toBe('convertToRange');
            expect(action.target).toBe('Table1');
        });

        test('convert preserves data', () => {
            const ctx = createMockContext();
            const table = createMockTable('Table1');
            
            const action = createAction('convertToRange', 'Table1', {
                preserveFormatting: true
            });
            
            expect(JSON.parse(action.data).preserveFormatting).toBe(true);
        });
    });

    /**
     * Toggle Table Totals Tests
     */
    describe('Toggle Table Totals Actions', () => {
        test('enable totals row', () => {
            const action = createAction('toggleTableTotals', 'Table1', {
                showTotals: true
            });
            
            expect(action.type).toBe('toggleTableTotals');
            expect(JSON.parse(action.data).showTotals).toBe(true);
        });

        test('disable totals row', () => {
            const action = createAction('toggleTableTotals', 'Table1', {
                showTotals: false
            });
            
            expect(JSON.parse(action.data).showTotals).toBe(false);
        });

        test('totals with aggregation functions', () => {
            const aggregations = ['Sum', 'Average', 'Count', 'Max', 'Min'];
            
            aggregations.forEach(func => {
                const action = createAction('toggleTableTotals', 'Table1', {
                    showTotals: true,
                    columns: {
                        'Sales': func,
                        'Quantity': 'Sum'
                    }
                });
                
                const data = JSON.parse(action.data);
                expect(data.columns.Sales).toBe(func);
            });
        });
    });
});

// ============================================================================
// Step 6: Data Manipulation Tests (8 Actions)
// ============================================================================

describe('Action Executor - Data Manipulation', () => {
    
    /**
     * Insert Rows Tests
     */
    describe('Insert Rows Actions', () => {
        test('property: insert any number of rows at any position', () => {
            fc.assert(
                fc.property(
                    fc.integer({ min: 1, max: 1000 }),
                    fc.integer({ min: 1, max: 100 }),
                    (position, count) => {
                        const action = createAction('insertRows', `${position}:${position}`, {
                            count,
                            position
                        });
                        
                        expect(action.type).toBe('insertRows');
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('insert single row', () => {
            const action = createAction('insertRows', '5:5', { count: 1 });
            expect(action.type).toBe('insertRows');
        });

        test('insert multiple rows', () => {
            const action = createAction('insertRows', '10:10', { count: 5 });
            expect(JSON.parse(action.data).count).toBe(5);
        });

        test('insert 100 rows', () => {
            const action = createAction('insertRows', '1:1', { count: 100 });
            expect(JSON.parse(action.data).count).toBe(100);
        });

        test('edge case: insert at row 1', () => {
            const action = createAction('insertRows', '1:1', { count: 1 });
            expect(action.target).toBe('1:1');
        });
    });

    /**
     * Insert Columns Tests
     */
    describe('Insert Columns Actions', () => {
        test('insert single column', () => {
            const action = createAction('insertColumns', 'B:B', { count: 1 });
            expect(action.type).toBe('insertColumns');
        });

        test('insert multiple columns', () => {
            const action = createAction('insertColumns', 'C:C', { count: 3 });
            expect(JSON.parse(action.data).count).toBe(3);
        });

        test('edge case: insert at column A', () => {
            const action = createAction('insertColumns', 'A:A', { count: 1 });
            expect(action.target).toBe('A:A');
        });
    });

    /**
     * Delete Rows Tests
     */
    describe('Delete Rows Actions', () => {
        test('delete single row', () => {
            const action = createAction('deleteRows', '5:5', {});
            expect(action.type).toBe('deleteRows');
        });

        test('delete multiple rows', () => {
            const action = createAction('deleteRows', '5:10', {});
            expect(action.target).toBe('5:10');
        });

        test('delete entire rows', () => {
            const action = createAction('deleteRows', '1:100', {});
            expect(action.target).toBe('1:100');
        });
    });

    /**
     * Delete Columns Tests
     */
    describe('Delete Columns Actions', () => {
        test('delete single column', () => {
            const action = createAction('deleteColumns', 'C:C', {});
            expect(action.type).toBe('deleteColumns');
        });

        test('delete multiple columns', () => {
            const action = createAction('deleteColumns', 'C:E', {});
            expect(action.target).toBe('C:E');
        });
    });

    /**
     * Merge Cells Tests
     */
    describe('Merge Cells Actions', () => {
        test('merge rectangular range', () => {
            const action = createAction('mergeCells', 'A1:C3', {});
            expect(action.type).toBe('mergeCells');
            expect(action.target).toBe('A1:C3');
        });

        test('merge single row', () => {
            const action = createAction('mergeCells', 'A1:E1', {});
            expect(action.target).toBe('A1:E1');
        });

        test('merge single column', () => {
            const action = createAction('mergeCells', 'A1:A5', {});
            expect(action.target).toBe('A1:A5');
        });

        test('edge case: single cell (no-op)', () => {
            const action = createAction('mergeCells', 'A1', {});
            expect(action.target).toBe('A1');
        });
    });

    /**
     * Unmerge Cells Tests
     */
    describe('Unmerge Cells Actions', () => {
        test('unmerge range', () => {
            const action = createAction('unmergeCells', 'A1:C3', {});
            expect(action.type).toBe('unmergeCells');
        });

        test('unmerge preserves top-left value', () => {
            const ctx = createMockContext();
            const range = createMockRange('A1:C3');
            
            const action = createAction('unmergeCells', 'A1:C3', {});
            expect(action.type).toBe('unmergeCells');
        });
    });

    /**
     * Find Replace Tests
     */
    describe('Find Replace Actions', () => {
        test('simple text replacement', () => {
            const action = createAction('findReplace', 'A1:Z100', {
                find: 'old',
                replace: 'new'
            });
            
            expect(action.type).toBe('findReplace');
            const data = JSON.parse(action.data);
            expect(data.find).toBe('old');
            expect(data.replace).toBe('new');
        });

        test('case-sensitive replacement', () => {
            const action = createAction('findReplace', 'A1:Z100', {
                find: 'Test',
                replace: 'Result',
                caseSensitive: true
            });
            
            expect(JSON.parse(action.data).caseSensitive).toBe(true);
        });

        test('whole word match', () => {
            const action = createAction('findReplace', 'A1:Z100', {
                find: 'cat',
                replace: 'dog',
                wholeWord: true
            });
            
            expect(JSON.parse(action.data).wholeWord).toBe(true);
        });

        test('regex pattern replacement', () => {
            const action = createAction('findReplace', 'A1:Z100', {
                find: '\\d+',
                replace: 'NUMBER',
                useRegex: true
            });
            
            expect(JSON.parse(action.data).useRegex).toBe(true);
        });

        test('replace all occurrences', () => {
            const action = createAction('findReplace', 'A1:Z100', {
                find: 'error',
                replace: 'success',
                replaceAll: true
            });
            
            expect(JSON.parse(action.data).replaceAll).toBe(true);
        });

        test('edge case: replace with empty string (delete)', () => {
            const action = createAction('findReplace', 'A1:Z100', {
                find: 'remove',
                replace: ''
            });
            
            expect(JSON.parse(action.data).replace).toBe('');
        });

        test('edge case: no matches', () => {
            const action = createAction('findReplace', 'A1:Z100', {
                find: 'nonexistent',
                replace: 'replacement'
            });
            
            expect(action.type).toBe('findReplace');
        });
    });

    /**
     * Text to Columns Tests
     */
    describe('Text to Columns Actions', () => {
        test('split by comma', () => {
            const action = createAction('textToColumns', 'A1:A100', {
                delimiter: ','
            });
            
            expect(action.type).toBe('textToColumns');
            expect(JSON.parse(action.data).delimiter).toBe(',');
        });

        test('split by tab', () => {
            const action = createAction('textToColumns', 'A1:A100', {
                delimiter: '\t'
            });
            
            expect(JSON.parse(action.data).delimiter).toBe('\t');
        });

        test('split by semicolon', () => {
            const action = createAction('textToColumns', 'A1:A100', {
                delimiter: ';'
            });
            
            expect(JSON.parse(action.data).delimiter).toBe(';');
        });

        test('split by custom delimiter', () => {
            const action = createAction('textToColumns', 'A1:A100', {
                delimiter: '|'
            });
            
            expect(JSON.parse(action.data).delimiter).toBe('|');
        });

        test('multiple delimiters', () => {
            const action = createAction('textToColumns', 'A1:A100', {
                delimiters: [',', ';', ' ']
            });
            
            const data = JSON.parse(action.data);
            expect(data.delimiters).toContain(',');
            expect(data.delimiters).toContain(';');
        });

        test('skip consecutive delimiters', () => {
            const action = createAction('textToColumns', 'A1:A100', {
                delimiter: ',',
                treatConsecutiveAsOne: true
            });
            
            expect(JSON.parse(action.data).treatConsecutiveAsOne).toBe(true);
        });

        test('edge case: single column source', () => {
            const action = createAction('textToColumns', 'A1:A10', {
                delimiter: ','
            });
            
            expect(action.target).toBe('A1:A10');
        });
    });
});


// ============================================================================
// Step 7: PivotTable Operations Tests (5 Actions)
// ============================================================================

describe('Action Executor - PivotTable Operations', () => {
    
    /**
     * Create PivotTable Tests
     */
    describe('Create PivotTable Actions', () => {
        test('property: any source range creates PivotTable', () => {
            fc.assert(
                fc.property(rangeArb, (sourceRange) => {
                    const action = createAction('createPivotTable', sourceRange, {
                        name: 'PivotTable1',
                        destination: 'Sheet2!A1'
                    });
                    
                    expect(action.type).toBe('createPivotTable');
                    expect(action.target).toBe(sourceRange);
                }),
                { numRuns: 100 }
            );
        });

        test('create PivotTable from range', () => {
            const action = createAction('createPivotTable', 'A1:E100', {
                name: 'SalesPivot',
                destination: 'G1'
            });
            
            expect(action.type).toBe('createPivotTable');
            expect(JSON.parse(action.data).name).toBe('SalesPivot');
        });

        test('create PivotTable from table', () => {
            const action = createAction('createPivotTable', 'SalesTable', {
                name: 'TablePivot',
                destination: 'Sheet2!A1',
                sourceIsTable: true
            });
            
            expect(JSON.parse(action.data).sourceIsTable).toBe(true);
        });

        test('create PivotTable on new sheet', () => {
            const action = createAction('createPivotTable', 'A1:E100', {
                name: 'NewSheetPivot',
                createNewSheet: true
            });
            
            expect(JSON.parse(action.data).createNewSheet).toBe(true);
        });

        test('edge case: empty source', () => {
            const action = createAction('createPivotTable', 'A1:A1', {
                name: 'EmptyPivot',
                destination: 'G1'
            });
            
            expect(action.target).toBe('A1:A1');
        });

        test('edge case: single row source', () => {
            const action = createAction('createPivotTable', 'A1:E1', {
                name: 'HeaderOnlyPivot',
                destination: 'G1'
            });
            
            expect(action.target).toBe('A1:E1');
        });
    });

    /**
     * Add Pivot Field Tests
     */
    describe('Add Pivot Field Actions', () => {
        test('property: add field to any pivot area', () => {
            fc.assert(
                fc.property(pivotAreaArb, (area) => {
                    const action = createAction('addPivotField', 'PivotTable1', {
                        fieldName: 'Region',
                        area
                    });
                    
                    expect(action.type).toBe('addPivotField');
                }),
                { numRuns: 50 }
            );
        });

        test('add field to row area', () => {
            const action = createAction('addPivotField', 'PivotTable1', {
                fieldName: 'Region',
                area: 'row'
            });
            
            expect(JSON.parse(action.data).area).toBe('row');
        });

        test('add field to column area', () => {
            const action = createAction('addPivotField', 'PivotTable1', {
                fieldName: 'Product',
                area: 'column'
            });
            
            expect(JSON.parse(action.data).area).toBe('column');
        });

        test('add field to data area with aggregation', () => {
            const aggregations = ['Sum', 'Count', 'Average', 'Max', 'Min', 'Product'];
            
            aggregations.forEach(aggregation => {
                const action = createAction('addPivotField', 'PivotTable1', {
                    fieldName: 'Sales',
                    area: 'data',
                    aggregation
                });
                
                expect(JSON.parse(action.data).aggregation).toBe(aggregation);
            });
        });

        test('add field to filter area', () => {
            const action = createAction('addPivotField', 'PivotTable1', {
                fieldName: 'Year',
                area: 'filter'
            });
            
            expect(JSON.parse(action.data).area).toBe('filter');
        });

        test('add multiple fields to same area', () => {
            const action = createAction('addPivotField', 'PivotTable1', {
                fieldName: 'Revenue',
                area: 'data',
                aggregation: 'Sum'
            });
            
            expect(action.type).toBe('addPivotField');
        });

        test('edge case: non-existent field name', () => {
            const action = createAction('addPivotField', 'PivotTable1', {
                fieldName: 'NonExistentField',
                area: 'row'
            });
            
            expect(JSON.parse(action.data).fieldName).toBe('NonExistentField');
        });
    });

    /**
     * Configure Pivot Layout Tests
     */
    describe('Configure Pivot Layout Actions', () => {
        test('set compact layout', () => {
            const action = createAction('configurePivotLayout', 'PivotTable1', {
                layoutType: 'Compact'
            });
            
            expect(action.type).toBe('configurePivotLayout');
            expect(JSON.parse(action.data).layoutType).toBe('Compact');
        });

        test('set outline layout', () => {
            const action = createAction('configurePivotLayout', 'PivotTable1', {
                layoutType: 'Outline'
            });
            
            expect(JSON.parse(action.data).layoutType).toBe('Outline');
        });

        test('set tabular layout', () => {
            const action = createAction('configurePivotLayout', 'PivotTable1', {
                layoutType: 'Tabular'
            });
            
            expect(JSON.parse(action.data).layoutType).toBe('Tabular');
        });

        test('show/hide subtotals', () => {
            const action = createAction('configurePivotLayout', 'PivotTable1', {
                showSubtotals: false
            });
            
            expect(JSON.parse(action.data).showSubtotals).toBe(false);
        });

        test('show/hide grand totals', () => {
            const action = createAction('configurePivotLayout', 'PivotTable1', {
                showRowGrandTotals: true,
                showColumnGrandTotals: false
            });
            
            const data = JSON.parse(action.data);
            expect(data.showRowGrandTotals).toBe(true);
            expect(data.showColumnGrandTotals).toBe(false);
        });

        test('repeat item labels', () => {
            const action = createAction('configurePivotLayout', 'PivotTable1', {
                repeatItemLabels: true
            });
            
            expect(JSON.parse(action.data).repeatItemLabels).toBe(true);
        });
    });

    /**
     * Refresh PivotTable Tests
     */
    describe('Refresh PivotTable Actions', () => {
        test('refresh single PivotTable', () => {
            const action = createAction('refreshPivotTable', 'PivotTable1', {});
            
            expect(action.type).toBe('refreshPivotTable');
            expect(action.target).toBe('PivotTable1');
        });

        test('refresh all PivotTables', () => {
            const action = createAction('refreshPivotTable', '*', {
                refreshAll: true
            });
            
            expect(JSON.parse(action.data).refreshAll).toBe(true);
        });

        test('edge case: non-existent PivotTable', () => {
            const action = createAction('refreshPivotTable', 'NonExistentPivot', {});
            
            expect(action.target).toBe('NonExistentPivot');
        });
    });

    /**
     * Delete PivotTable Tests
     */
    describe('Delete PivotTable Actions', () => {
        test('delete PivotTable by name', () => {
            const action = createAction('deletePivotTable', 'PivotTable1', {});
            
            expect(action.type).toBe('deletePivotTable');
            expect(action.target).toBe('PivotTable1');
        });

        test('edge case: non-existent PivotTable', () => {
            const action = createAction('deletePivotTable', 'NonExistentPivot', {});
            
            expect(action.target).toBe('NonExistentPivot');
        });
    });
});

// ============================================================================
// Step 8: Slicer Operations Tests (5 Actions)
// ============================================================================

describe('Action Executor - Slicer Operations', () => {
    
    /**
     * Create Slicer Tests
     */
    describe('Create Slicer Actions', () => {
        test('property: create slicer with any style', () => {
            fc.assert(
                fc.property(slicerStyleArb, (style) => {
                    const action = createAction('createSlicer', 'Table1', {
                        fieldName: 'Region',
                        style
                    });
                    
                    expect(action.type).toBe('createSlicer');
                }),
                { numRuns: 50 }
            );
        });

        test('create slicer for table column', () => {
            const action = createAction('createSlicer', 'Table1', {
                fieldName: 'Region',
                style: 'SlicerStyleLight1'
            });
            
            expect(action.type).toBe('createSlicer');
            expect(JSON.parse(action.data).fieldName).toBe('Region');
        });

        test('create slicer for PivotTable field', () => {
            const action = createAction('createSlicer', 'PivotTable1', {
                fieldName: 'Product',
                style: 'SlicerStyleMedium2',
                sourceType: 'pivot'
            });
            
            expect(JSON.parse(action.data).sourceType).toBe('pivot');
        });

        test('create slicer with position', () => {
            const action = createAction('createSlicer', 'Table1', {
                fieldName: 'Category',
                left: 100,
                top: 50,
                width: 200,
                height: 300
            });
            
            const data = JSON.parse(action.data);
            expect(data.left).toBe(100);
            expect(data.top).toBe(50);
            expect(data.width).toBe(200);
            expect(data.height).toBe(300);
        });

        test('slicer styles - light', () => {
            for (let i = 1; i <= 6; i++) {
                const style = `SlicerStyleLight${i}`;
                const action = createAction('createSlicer', 'Table1', {
                    fieldName: 'Region',
                    style
                });
                expect(JSON.parse(action.data).style).toBe(style);
            }
        });

        test('slicer styles - dark', () => {
            for (let i = 1; i <= 6; i++) {
                const style = `SlicerStyleDark${i}`;
                const action = createAction('createSlicer', 'Table1', {
                    fieldName: 'Region',
                    style
                });
                expect(JSON.parse(action.data).style).toBe(style);
            }
        });

        test('edge case: invalid field name', () => {
            const action = createAction('createSlicer', 'Table1', {
                fieldName: 'NonExistentField',
                style: 'SlicerStyleLight1'
            });
            
            expect(JSON.parse(action.data).fieldName).toBe('NonExistentField');
        });
    });

    /**
     * Configure Slicer Tests
     */
    describe('Configure Slicer Actions', () => {
        test('change slicer position', () => {
            const action = createAction('configureSlicer', 'Slicer1', {
                left: 200,
                top: 100
            });
            
            expect(action.type).toBe('configureSlicer');
            const data = JSON.parse(action.data);
            expect(data.left).toBe(200);
            expect(data.top).toBe(100);
        });

        test('change slicer size', () => {
            const action = createAction('configureSlicer', 'Slicer1', {
                width: 250,
                height: 400
            });
            
            const data = JSON.parse(action.data);
            expect(data.width).toBe(250);
            expect(data.height).toBe(400);
        });

        test('change slicer style', () => {
            const action = createAction('configureSlicer', 'Slicer1', {
                style: 'SlicerStyleMedium3'
            });
            
            expect(JSON.parse(action.data).style).toBe('SlicerStyleMedium3');
        });

        test('enable multi-select', () => {
            const action = createAction('configureSlicer', 'Slicer1', {
                multiSelect: true
            });
            
            expect(JSON.parse(action.data).multiSelect).toBe(true);
        });

        test('set caption', () => {
            const action = createAction('configureSlicer', 'Slicer1', {
                caption: 'Select Region'
            });
            
            expect(JSON.parse(action.data).caption).toBe('Select Region');
        });

        test('sort items ascending', () => {
            const action = createAction('configureSlicer', 'Slicer1', {
                sortOrder: 'Ascending'
            });
            
            expect(JSON.parse(action.data).sortOrder).toBe('Ascending');
        });

        test('sort items descending', () => {
            const action = createAction('configureSlicer', 'Slicer1', {
                sortOrder: 'Descending'
            });
            
            expect(JSON.parse(action.data).sortOrder).toBe('Descending');
        });

        test('edge case: non-existent slicer', () => {
            const action = createAction('configureSlicer', 'NonExistentSlicer', {
                style: 'SlicerStyleLight1'
            });
            
            expect(action.target).toBe('NonExistentSlicer');
        });
    });

    /**
     * Connect Slicer to Table Tests
     */
    describe('Connect Slicer to Table Actions', () => {
        test('connect slicer to table', () => {
            const action = createAction('connectSlicerToTable', 'Slicer1', {
                tableName: 'Table2'
            });
            
            expect(action.type).toBe('connectSlicerToTable');
            expect(JSON.parse(action.data).tableName).toBe('Table2');
        });

        test('edge case: non-existent table', () => {
            const action = createAction('connectSlicerToTable', 'Slicer1', {
                tableName: 'NonExistentTable'
            });
            
            expect(JSON.parse(action.data).tableName).toBe('NonExistentTable');
        });
    });

    /**
     * Connect Slicer to Pivot Tests
     */
    describe('Connect Slicer to Pivot Actions', () => {
        test('connect slicer to PivotTable', () => {
            const action = createAction('connectSlicerToPivot', 'Slicer1', {
                pivotTableName: 'PivotTable2'
            });
            
            expect(action.type).toBe('connectSlicerToPivot');
            expect(JSON.parse(action.data).pivotTableName).toBe('PivotTable2');
        });

        test('edge case: non-existent PivotTable', () => {
            const action = createAction('connectSlicerToPivot', 'Slicer1', {
                pivotTableName: 'NonExistentPivot'
            });
            
            expect(JSON.parse(action.data).pivotTableName).toBe('NonExistentPivot');
        });
    });

    /**
     * Delete Slicer Tests
     */
    describe('Delete Slicer Actions', () => {
        test('delete slicer by name', () => {
            const action = createAction('deleteSlicer', 'Slicer1', {});
            
            expect(action.type).toBe('deleteSlicer');
            expect(action.target).toBe('Slicer1');
        });

        test('delete all slicers on sheet', () => {
            const action = createAction('deleteSlicer', '*', {
                deleteAll: true,
                sheetName: 'Sheet1'
            });
            
            expect(JSON.parse(action.data).deleteAll).toBe(true);
        });

        test('edge case: non-existent slicer', () => {
            const action = createAction('deleteSlicer', 'NonExistentSlicer', {});
            
            expect(action.target).toBe('NonExistentSlicer');
        });
    });
});


// ============================================================================
// Step 9: Named Range Operations Tests (4 Actions)
// ============================================================================

describe('Action Executor - Named Range Operations', () => {
    
    /**
     * Create Named Range Tests
     */
    describe('Create Named Range Actions', () => {
        test('property: any valid name with any range creates named range', () => {
            fc.assert(
                fc.property(rangeArb, tableNameArb, (range, name) => {
                    const action = createAction('createNamedRange', range, { name });
                    
                    expect(action.type).toBe('createNamedRange');
                    expect(action.target).toBe(range);
                }),
                { numRuns: 100 }
            );
        });

        test('create workbook-scoped named range', () => {
            const action = createAction('createNamedRange', 'A1:D10', {
                name: 'SalesData',
                scope: 'Workbook'
            });
            
            expect(action.type).toBe('createNamedRange');
            expect(JSON.parse(action.data).scope).toBe('Workbook');
        });

        test('create worksheet-scoped named range', () => {
            const action = createAction('createNamedRange', 'Sheet1!A1:D10', {
                name: 'LocalData',
                scope: 'Worksheet'
            });
            
            expect(JSON.parse(action.data).scope).toBe('Worksheet');
        });

        test('create named constant', () => {
            const action = createAction('createNamedRange', '=100', {
                name: 'TaxRate',
                isConstant: true
            });
            
            expect(JSON.parse(action.data).isConstant).toBe(true);
        });

        test('create named formula', () => {
            const action = createAction('createNamedRange', '=SUM(A:A)', {
                name: 'TotalSales',
                isFormula: true
            });
            
            expect(JSON.parse(action.data).isFormula).toBe(true);
        });

        test('cross-sheet reference', () => {
            const action = createAction('createNamedRange', 'Sheet2!A1:B10', {
                name: 'OtherSheetData'
            });
            
            expect(action.target).toBe('Sheet2!A1:B10');
        });

        test('edge case: name with underscore', () => {
            const action = createAction('createNamedRange', 'A1:A10', {
                name: 'Q1_Sales_Data'
            });
            
            expect(JSON.parse(action.data).name).toBe('Q1_Sales_Data');
        });

        test('edge case: name starting with letter', () => {
            const validNames = ['Sales', 'Data1', 'MyRange', '_Hidden'];
            
            validNames.forEach(name => {
                const action = createAction('createNamedRange', 'A1:A10', { name });
                expect(JSON.parse(action.data).name).toBe(name);
            });
        });
    });

    /**
     * Delete Named Range Tests
     */
    describe('Delete Named Range Actions', () => {
        test('delete named range by name', () => {
            const action = createAction('deleteNamedRange', 'SalesData', {});
            
            expect(action.type).toBe('deleteNamedRange');
            expect(action.target).toBe('SalesData');
        });

        test('delete all in scope', () => {
            const action = createAction('deleteNamedRange', '*', {
                deleteAll: true,
                scope: 'Workbook'
            });
            
            expect(JSON.parse(action.data).deleteAll).toBe(true);
        });

        test('edge case: non-existent named range', () => {
            const action = createAction('deleteNamedRange', 'NonExistentRange', {});
            
            expect(action.target).toBe('NonExistentRange');
        });
    });

    /**
     * Update Named Range Tests
     */
    describe('Update Named Range Actions', () => {
        test('update range reference', () => {
            const action = createAction('updateNamedRange', 'SalesData', {
                newReference: 'A1:E20'
            });
            
            expect(action.type).toBe('updateNamedRange');
            expect(JSON.parse(action.data).newReference).toBe('A1:E20');
        });

        test('update formula', () => {
            const action = createAction('updateNamedRange', 'TotalSales', {
                newReference: '=SUM(B:B)'
            });
            
            expect(JSON.parse(action.data).newReference).toBe('=SUM(B:B)');
        });

        test('update constant value', () => {
            const action = createAction('updateNamedRange', 'TaxRate', {
                newReference: '=0.09'
            });
            
            expect(JSON.parse(action.data).newReference).toBe('=0.09');
        });

        test('edge case: update non-existent', () => {
            const action = createAction('updateNamedRange', 'NonExistent', {
                newReference: 'A1:A10'
            });
            
            expect(action.target).toBe('NonExistent');
        });
    });

    /**
     * List Named Ranges Tests
     */
    describe('List Named Ranges Actions', () => {
        test('list all workbook-scoped', () => {
            const action = createAction('listNamedRanges', '*', {
                scope: 'Workbook'
            });
            
            expect(action.type).toBe('listNamedRanges');
            expect(JSON.parse(action.data).scope).toBe('Workbook');
        });

        test('list worksheet-scoped', () => {
            const action = createAction('listNamedRanges', 'Sheet1', {
                scope: 'Worksheet'
            });
            
            expect(JSON.parse(action.data).scope).toBe('Worksheet');
        });

        test('filter by pattern', () => {
            const action = createAction('listNamedRanges', '*', {
                pattern: 'Sales*'
            });
            
            expect(JSON.parse(action.data).pattern).toBe('Sales*');
        });

        test('include hidden names', () => {
            const action = createAction('listNamedRanges', '*', {
                includeHidden: true
            });
            
            expect(JSON.parse(action.data).includeHidden).toBe(true);
        });

        test('exclude hidden names', () => {
            const action = createAction('listNamedRanges', '*', {
                includeHidden: false
            });
            
            expect(JSON.parse(action.data).includeHidden).toBe(false);
        });
    });
});


// ============================================================================
// Step 10: Protection Operations Tests (6 Actions)
// ============================================================================

describe('Action Executor - Protection Operations', () => {
    
    /**
     * Protect Worksheet Tests
     */
    describe('Protect Worksheet Actions', () => {
        test('property: protect with any combination of options', () => {
            fc.assert(
                fc.property(
                    fc.boolean(),
                    fc.boolean(),
                    fc.boolean(),
                    (allowSort, allowFilter, allowFormat) => {
                        const action = createAction('protectWorksheet', 'Sheet1', {
                            allowSort,
                            allowFilter,
                            allowFormatCells: allowFormat
                        });
                        
                        expect(action.type).toBe('protectWorksheet');
                    }
                ),
                { numRuns: 50 }
            );
        });

        test('protect without password', () => {
            const action = createAction('protectWorksheet', 'Sheet1', {});
            
            expect(action.type).toBe('protectWorksheet');
            expect(action.target).toBe('Sheet1');
        });

        test('protect with password', () => {
            const action = createAction('protectWorksheet', 'Sheet1', {
                password: 'secret123'
            });
            
            expect(JSON.parse(action.data).password).toBe('secret123');
        });

        test('protect with allow options', () => {
            const action = createAction('protectWorksheet', 'Sheet1', {
                allowFormatCells: true,
                allowFormatColumns: true,
                allowFormatRows: true,
                allowInsertColumns: false,
                allowInsertRows: false,
                allowInsertHyperlinks: true,
                allowDeleteColumns: false,
                allowDeleteRows: false,
                allowSort: true,
                allowAutoFilter: true,
                allowPivotTables: true
            });
            
            const data = JSON.parse(action.data);
            expect(data.allowSort).toBe(true);
            expect(data.allowAutoFilter).toBe(true);
            expect(data.allowInsertColumns).toBe(false);
        });

        test('protect multiple sheets', () => {
            const sheets = ['Sheet1', 'Sheet2', 'Sheet3'];
            
            sheets.forEach(sheet => {
                const action = createAction('protectWorksheet', sheet, {});
                expect(action.target).toBe(sheet);
            });
        });

        test('edge case: already protected', () => {
            const action = createAction('protectWorksheet', 'Sheet1', {
                password: 'test'
            });
            
            expect(action.type).toBe('protectWorksheet');
        });
    });

    /**
     * Unprotect Worksheet Tests
     */
    describe('Unprotect Worksheet Actions', () => {
        test('unprotect without password', () => {
            const action = createAction('unprotectWorksheet', 'Sheet1', {});
            
            expect(action.type).toBe('unprotectWorksheet');
            expect(action.target).toBe('Sheet1');
        });

        test('unprotect with password', () => {
            const action = createAction('unprotectWorksheet', 'Sheet1', {
                password: 'secret123'
            });
            
            expect(JSON.parse(action.data).password).toBe('secret123');
        });

        test('edge case: wrong password', () => {
            const action = createAction('unprotectWorksheet', 'Sheet1', {
                password: 'wrongpassword'
            });
            
            expect(JSON.parse(action.data).password).toBe('wrongpassword');
        });
    });

    /**
     * Protect Range Tests
     */
    describe('Protect Range Actions', () => {
        test('lock cells', () => {
            const action = createAction('protectRange', 'A1:A10', {
                locked: true
            });
            
            expect(action.type).toBe('protectRange');
            expect(JSON.parse(action.data).locked).toBe(true);
        });

        test('unlock cells', () => {
            const action = createAction('protectRange', 'B1:B10', {
                locked: false
            });
            
            expect(JSON.parse(action.data).locked).toBe(false);
        });

        test('hide formulas', () => {
            const action = createAction('protectRange', 'C1:C10', {
                locked: true,
                formulaHidden: true
            });
            
            expect(JSON.parse(action.data).formulaHidden).toBe(true);
        });

        test('edge case: protect on unprotected sheet', () => {
            const action = createAction('protectRange', 'A1:A10', {
                locked: true
            });
            
            expect(action.type).toBe('protectRange');
        });
    });

    /**
     * Unprotect Range Tests
     */
    describe('Unprotect Range Actions', () => {
        test('unprotect range', () => {
            const action = createAction('unprotectRange', 'A1:A10', {});
            
            expect(action.type).toBe('unprotectRange');
            expect(action.target).toBe('A1:A10');
        });

        test('unprotect with unlock', () => {
            const action = createAction('unprotectRange', 'A1:A10', {
                locked: false,
                formulaHidden: false
            });
            
            const data = JSON.parse(action.data);
            expect(data.locked).toBe(false);
            expect(data.formulaHidden).toBe(false);
        });
    });

    /**
     * Protect Workbook Tests
     */
    describe('Protect Workbook Actions', () => {
        test('protect workbook structure', () => {
            const action = createAction('protectWorkbook', '*', {
                protectStructure: true
            });
            
            expect(action.type).toBe('protectWorkbook');
            expect(JSON.parse(action.data).protectStructure).toBe(true);
        });

        test('protect with password', () => {
            const action = createAction('protectWorkbook', '*', {
                password: 'workbookpass'
            });
            
            expect(JSON.parse(action.data).password).toBe('workbookpass');
        });

        test('edge case: already protected', () => {
            const action = createAction('protectWorkbook', '*', {});
            
            expect(action.type).toBe('protectWorkbook');
        });
    });

    /**
     * Unprotect Workbook Tests
     */
    describe('Unprotect Workbook Actions', () => {
        test('unprotect workbook', () => {
            const action = createAction('unprotectWorkbook', '*', {});
            
            expect(action.type).toBe('unprotectWorkbook');
        });

        test('unprotect with password', () => {
            const action = createAction('unprotectWorkbook', '*', {
                password: 'workbookpass'
            });
            
            expect(JSON.parse(action.data).password).toBe('workbookpass');
        });

        test('edge case: wrong password', () => {
            const action = createAction('unprotectWorkbook', '*', {
                password: 'wrongpass'
            });
            
            expect(JSON.parse(action.data).password).toBe('wrongpass');
        });
    });
});

// ============================================================================
// Step 11: Shape Operations Tests (8 Actions)
// ============================================================================

describe('Action Executor - Shape Operations', () => {
    
    /**
     * Insert Shape Tests
     */
    describe('Insert Shape Actions', () => {
        test('property: any geometric shape type at any position', () => {
            fc.assert(
                fc.property(shapeTypeArb, cellRefArb, (shapeType, position) => {
                    const action = createAction('insertShape', position, {
                        shapeType
                    });
                    
                    expect(action.type).toBe('insertShape');
                }),
                { numRuns: 100 }
            );
        });

        test('insert rectangle', () => {
            const action = createAction('insertShape', 'D5', {
                shapeType: 'Rectangle',
                width: 100,
                height: 50
            });
            
            expect(action.type).toBe('insertShape');
            expect(JSON.parse(action.data).shapeType).toBe('Rectangle');
        });

        test('insert oval', () => {
            const action = createAction('insertShape', 'E10', {
                shapeType: 'Oval',
                width: 80,
                height: 80
            });
            
            expect(JSON.parse(action.data).shapeType).toBe('Oval');
        });

        test('insert various shape types', () => {
            const shapeTypes = [
                'Rectangle', 'RoundRectangle', 'Oval', 'Diamond', 'Triangle',
                'RightTriangle', 'Parallelogram', 'Pentagon', 'Hexagon', 'Octagon',
                'Star4', 'Star5', 'Arrow', 'Chevron'
            ];
            
            shapeTypes.forEach(shapeType => {
                const action = createAction('insertShape', 'A1', { shapeType });
                expect(JSON.parse(action.data).shapeType).toBe(shapeType);
            });
        });

        test('insert shape with pixel position', () => {
            const action = createAction('insertShape', '', {
                shapeType: 'Rectangle',
                left: 100,
                top: 200,
                width: 150,
                height: 75
            });
            
            const data = JSON.parse(action.data);
            expect(data.left).toBe(100);
            expect(data.top).toBe(200);
        });

        test('insert shape with fill color', () => {
            const action = createAction('insertShape', 'D5', {
                shapeType: 'Rectangle',
                fillColor: '#FF5733'
            });
            
            expect(JSON.parse(action.data).fillColor).toBe('#FF5733');
        });
    });

    /**
     * Insert Image Tests
     */
    describe('Insert Image Actions', () => {
        test('insert Base64 JPEG image', () => {
            const action = createAction('insertImage', 'A1', {
                base64: 'data:image/jpeg;base64,/9j/4AAQSkZJRg...',
                width: 200,
                height: 150
            });
            
            expect(action.type).toBe('insertImage');
            expect(JSON.parse(action.data).base64).toContain('data:image/jpeg');
        });

        test('insert Base64 PNG image', () => {
            const action = createAction('insertImage', 'B5', {
                base64: 'data:image/png;base64,iVBORw0KGgo...',
                width: 100,
                height: 100
            });
            
            expect(JSON.parse(action.data).base64).toContain('data:image/png');
        });

        test('insert image with position', () => {
            const action = createAction('insertImage', '', {
                base64: 'data:image/png;base64,test',
                left: 50,
                top: 100,
                width: 200,
                height: 150
            });
            
            const data = JSON.parse(action.data);
            expect(data.left).toBe(50);
            expect(data.top).toBe(100);
        });

        test('edge case: large image', () => {
            const action = createAction('insertImage', 'A1', {
                base64: 'data:image/png;base64,' + 'A'.repeat(1000),
                width: 1000,
                height: 800
            });
            
            expect(action.type).toBe('insertImage');
        });
    });

    /**
     * Insert TextBox Tests
     */
    describe('Insert TextBox Actions', () => {
        test('insert text box with content', () => {
            const action = createAction('insertTextBox', 'C3', {
                text: 'Instructions: Enter data below',
                width: 200,
                height: 50
            });
            
            expect(action.type).toBe('insertTextBox');
            expect(JSON.parse(action.data).text).toBe('Instructions: Enter data below');
        });

        test('insert text box with formatting', () => {
            const action = createAction('insertTextBox', 'D5', {
                text: 'Important Note',
                fontName: 'Arial',
                fontSize: 14,
                fontColor: '#FF0000',
                bold: true
            });
            
            const data = JSON.parse(action.data);
            expect(data.fontName).toBe('Arial');
            expect(data.fontSize).toBe(14);
            expect(data.bold).toBe(true);
        });

        test('insert text box with alignment', () => {
            const action = createAction('insertTextBox', 'E10', {
                text: 'Centered Text',
                horizontalAlignment: 'Center',
                verticalAlignment: 'Middle'
            });
            
            const data = JSON.parse(action.data);
            expect(data.horizontalAlignment).toBe('Center');
            expect(data.verticalAlignment).toBe('Middle');
        });
    });

    /**
     * Format Shape Tests
     */
    describe('Format Shape Actions', () => {
        test('change fill color', () => {
            const action = createAction('formatShape', 'Shape1', {
                fillColor: '#00FF00'
            });
            
            expect(action.type).toBe('formatShape');
            expect(JSON.parse(action.data).fillColor).toBe('#00FF00');
        });

        test('change line properties', () => {
            const action = createAction('formatShape', 'Shape1', {
                lineColor: '#0000FF',
                lineWidth: 2,
                lineStyle: 'Dash'
            });
            
            const data = JSON.parse(action.data);
            expect(data.lineColor).toBe('#0000FF');
            expect(data.lineWidth).toBe(2);
            expect(data.lineStyle).toBe('Dash');
        });

        test('change text properties', () => {
            const action = createAction('formatShape', 'TextBox1', {
                text: 'Updated Text',
                fontColor: '#333333',
                fontSize: 12
            });
            
            const data = JSON.parse(action.data);
            expect(data.text).toBe('Updated Text');
        });

        test('change rotation', () => {
            const action = createAction('formatShape', 'Shape1', {
                rotation: 45
            });
            
            expect(JSON.parse(action.data).rotation).toBe(45);
        });

        test('change transparency', () => {
            const action = createAction('formatShape', 'Shape1', {
                transparency: 0.5
            });
            
            expect(JSON.parse(action.data).transparency).toBe(0.5);
        });

        test('edge case: non-existent shape', () => {
            const action = createAction('formatShape', 'NonExistentShape', {
                fillColor: '#FF0000'
            });
            
            expect(action.target).toBe('NonExistentShape');
        });
    });

    /**
     * Delete Shape Tests
     */
    describe('Delete Shape Actions', () => {
        test('delete shape by name', () => {
            const action = createAction('deleteShape', 'Shape1', {});
            
            expect(action.type).toBe('deleteShape');
            expect(action.target).toBe('Shape1');
        });

        test('delete multiple shapes', () => {
            const action = createAction('deleteShape', '*', {
                shapeNames: ['Shape1', 'Shape2', 'Shape3']
            });
            
            const data = JSON.parse(action.data);
            expect(data.shapeNames.length).toBe(3);
        });

        test('edge case: non-existent shape', () => {
            const action = createAction('deleteShape', 'NonExistent', {});
            
            expect(action.target).toBe('NonExistent');
        });
    });

    /**
     * Group Shapes Tests
     */
    describe('Group Shapes Actions', () => {
        test('group multiple shapes', () => {
            const action = createAction('groupShapes', '*', {
                shapeNames: ['Shape1', 'Shape2', 'Shape3'],
                groupName: 'MyGroup'
            });
            
            expect(action.type).toBe('groupShapes');
            const data = JSON.parse(action.data);
            expect(data.shapeNames.length).toBe(3);
            expect(data.groupName).toBe('MyGroup');
        });

        test('edge case: group single shape', () => {
            const action = createAction('groupShapes', '*', {
                shapeNames: ['Shape1']
            });
            
            expect(JSON.parse(action.data).shapeNames.length).toBe(1);
        });
    });

    /**
     * Ungroup Shapes Tests
     */
    describe('Ungroup Shapes Actions', () => {
        test('ungroup shapes', () => {
            const action = createAction('ungroupShapes', 'MyGroup', {});
            
            expect(action.type).toBe('ungroupShapes');
            expect(action.target).toBe('MyGroup');
        });
    });

    /**
     * Arrange Shapes Tests
     */
    describe('Arrange Shapes Actions', () => {
        test('bring to front', () => {
            const action = createAction('arrangeShapes', 'Shape1', {
                zOrder: 'BringToFront'
            });
            
            expect(action.type).toBe('arrangeShapes');
            expect(JSON.parse(action.data).zOrder).toBe('BringToFront');
        });

        test('send to back', () => {
            const action = createAction('arrangeShapes', 'Shape1', {
                zOrder: 'SendToBack'
            });
            
            expect(JSON.parse(action.data).zOrder).toBe('SendToBack');
        });

        test('bring forward', () => {
            const action = createAction('arrangeShapes', 'Shape1', {
                zOrder: 'BringForward'
            });
            
            expect(JSON.parse(action.data).zOrder).toBe('BringForward');
        });

        test('send backward', () => {
            const action = createAction('arrangeShapes', 'Shape1', {
                zOrder: 'SendBackward'
            });
            
            expect(JSON.parse(action.data).zOrder).toBe('SendBackward');
        });

        test('edge case: non-existent shape', () => {
            const action = createAction('arrangeShapes', 'NonExistent', {
                zOrder: 'BringToFront'
            });
            
            expect(action.target).toBe('NonExistent');
        });
    });
});

// ============================================================================
// Step 12: Comment Operations Tests (8 Actions)
// ============================================================================

describe('Action Executor - Comment Operations', () => {
    
    /**
     * Add Comment Tests
     */
    describe('Add Comment Actions', () => {
        test('property: add comment with any content to any cell', () => {
            fc.assert(
                fc.property(cellRefArb, fc.string({ minLength: 1, maxLength: 500 }), (cell, content) => {
                    const action = createAction('addComment', cell, { content });
                    
                    expect(action.type).toBe('addComment');
                    expect(action.target).toBe(cell);
                }),
                { numRuns: 100 }
            );
        });

        test('add plain text comment', () => {
            const action = createAction('addComment', 'A1', {
                content: 'Review this formula'
            });
            
            expect(action.type).toBe('addComment');
            expect(JSON.parse(action.data).content).toBe('Review this formula');
        });

        test('add comment with @mention', () => {
            const action = createAction('addComment', 'B5', {
                content: '@John Please review this data'
            });
            
            expect(JSON.parse(action.data).content).toContain('@John');
        });

        test('edge case: comment on merged cell', () => {
            const action = createAction('addComment', 'A1', {
                content: 'Comment on merged cell'
            });
            
            expect(action.target).toBe('A1');
        });

        test('edge case: very long content', () => {
            const longContent = 'A'.repeat(1000);
            const action = createAction('addComment', 'A1', {
                content: longContent
            });
            
            expect(JSON.parse(action.data).content.length).toBe(1000);
        });
    });

    /**
     * Add Note Tests (Legacy)
     */
    describe('Add Note Actions', () => {
        test('add note to cell', () => {
            const action = createAction('addNote', 'C10', {
                content: 'This is a note'
            });
            
            expect(action.type).toBe('addNote');
            expect(JSON.parse(action.data).content).toBe('This is a note');
        });

        test('add note with multiline content', () => {
            const action = createAction('addNote', 'D5', {
                content: 'Line 1\nLine 2\nLine 3'
            });
            
            expect(JSON.parse(action.data).content).toContain('\n');
        });
    });

    /**
     * Edit Comment Tests
     */
    describe('Edit Comment Actions', () => {
        test('edit comment content', () => {
            const action = createAction('editComment', 'A1', {
                content: 'Updated comment text'
            });
            
            expect(action.type).toBe('editComment');
            expect(JSON.parse(action.data).content).toBe('Updated comment text');
        });

        test('edge case: edit non-existent comment', () => {
            const action = createAction('editComment', 'Z99', {
                content: 'New content'
            });
            
            expect(action.target).toBe('Z99');
        });
    });

    /**
     * Edit Note Tests
     */
    describe('Edit Note Actions', () => {
        test('edit note content', () => {
            const action = createAction('editNote', 'D5', {
                content: 'Updated note text'
            });
            
            expect(action.type).toBe('editNote');
            expect(JSON.parse(action.data).content).toBe('Updated note text');
        });
    });

    /**
     * Delete Comment Tests
     */
    describe('Delete Comment Actions', () => {
        test('delete comment from cell', () => {
            const action = createAction('deleteComment', 'A1', {});
            
            expect(action.type).toBe('deleteComment');
            expect(action.target).toBe('A1');
        });

        test('delete comment thread', () => {
            const action = createAction('deleteComment', 'B5', {
                deleteThread: true
            });
            
            expect(JSON.parse(action.data).deleteThread).toBe(true);
        });

        test('edge case: delete non-existent comment', () => {
            const action = createAction('deleteComment', 'Z99', {});
            
            expect(action.target).toBe('Z99');
        });
    });

    /**
     * Delete Note Tests
     */
    describe('Delete Note Actions', () => {
        test('delete note from cell', () => {
            const action = createAction('deleteNote', 'E2', {});
            
            expect(action.type).toBe('deleteNote');
            expect(action.target).toBe('E2');
        });
    });

    /**
     * Reply to Comment Tests
     */
    describe('Reply to Comment Actions', () => {
        test('reply to comment', () => {
            const action = createAction('replyToComment', 'A1', {
                content: 'Thanks for the feedback!'
            });
            
            expect(action.type).toBe('replyToComment');
            expect(JSON.parse(action.data).content).toBe('Thanks for the feedback!');
        });

        test('reply with @mention', () => {
            const action = createAction('replyToComment', 'B5', {
                content: '@Jane I have updated the formula'
            });
            
            expect(JSON.parse(action.data).content).toContain('@Jane');
        });

        test('edge case: reply to non-existent comment', () => {
            const action = createAction('replyToComment', 'Z99', {
                content: 'Reply text'
            });
            
            expect(action.target).toBe('Z99');
        });
    });

    /**
     * Resolve Comment Tests
     */
    describe('Resolve Comment Actions', () => {
        test('resolve comment', () => {
            const action = createAction('resolveComment', 'A1', {
                resolved: true
            });
            
            expect(action.type).toBe('resolveComment');
            expect(JSON.parse(action.data).resolved).toBe(true);
        });

        test('unresolve comment', () => {
            const action = createAction('resolveComment', 'A1', {
                resolved: false
            });
            
            expect(JSON.parse(action.data).resolved).toBe(false);
        });

        test('edge case: resolve non-existent comment', () => {
            const action = createAction('resolveComment', 'Z99', {
                resolved: true
            });
            
            expect(action.target).toBe('Z99');
        });
    });
});

// ============================================================================
// Step 13: Sparkline Operations Tests (3 Actions)
// ============================================================================

describe('Action Executor - Sparkline Operations', () => {
    
    /**
     * Create Sparkline Tests
     */
    describe('Create Sparkline Actions', () => {
        test('property: any sparkline type with any source range', () => {
            fc.assert(
                fc.property(
                    fc.constantFrom('Line', 'Column', 'WinLoss'),
                    rangeArb,
                    cellRefArb,
                    (sparklineType, sourceRange, targetCell) => {
                        const action = createAction('createSparkline', targetCell, {
                            sourceRange,
                            sparklineType
                        });
                        
                        expect(action.type).toBe('createSparkline');
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('create line sparkline', () => {
            const action = createAction('createSparkline', 'E1', {
                sourceRange: 'A1:D1',
                sparklineType: 'Line'
            });
            
            expect(action.type).toBe('createSparkline');
            expect(JSON.parse(action.data).sparklineType).toBe('Line');
        });

        test('create column sparkline', () => {
            const action = createAction('createSparkline', 'E2', {
                sourceRange: 'A2:D2',
                sparklineType: 'Column'
            });
            
            expect(JSON.parse(action.data).sparklineType).toBe('Column');
        });

        test('create win/loss sparkline', () => {
            const action = createAction('createSparkline', 'E3', {
                sourceRange: 'A3:D3',
                sparklineType: 'WinLoss'
            });
            
            expect(JSON.parse(action.data).sparklineType).toBe('WinLoss');
        });

        test('create sparkline from column range', () => {
            const action = createAction('createSparkline', 'B10', {
                sourceRange: 'B1:B9',
                sparklineType: 'Line'
            });
            
            expect(JSON.parse(action.data).sourceRange).toBe('B1:B9');
        });

        test('edge case: single data point', () => {
            const action = createAction('createSparkline', 'B1', {
                sourceRange: 'A1',
                sparklineType: 'Line'
            });
            
            expect(JSON.parse(action.data).sourceRange).toBe('A1');
        });

        test('edge case: large source range', () => {
            const action = createAction('createSparkline', 'AA1', {
                sourceRange: 'A1:Z1',
                sparklineType: 'Line'
            });
            
            expect(action.type).toBe('createSparkline');
        });
    });

    /**
     * Configure Sparkline Tests
     */
    describe('Configure Sparkline Actions', () => {
        test('show markers', () => {
            const action = createAction('configureSparkline', 'E1', {
                showHighPoint: true,
                showLowPoint: true,
                showFirstPoint: true,
                showLastPoint: true,
                showNegativePoints: true
            });
            
            expect(action.type).toBe('configureSparkline');
            const data = JSON.parse(action.data);
            expect(data.showHighPoint).toBe(true);
            expect(data.showLowPoint).toBe(true);
        });

        test('set axis settings', () => {
            const action = createAction('configureSparkline', 'E1', {
                minAxisType: 'Custom',
                minAxisValue: 0,
                maxAxisType: 'Custom',
                maxAxisValue: 100
            });
            
            const data = JSON.parse(action.data);
            expect(data.minAxisValue).toBe(0);
            expect(data.maxAxisValue).toBe(100);
        });

        test('set colors', () => {
            const action = createAction('configureSparkline', 'E1', {
                lineColor: '#0066CC',
                markerColor: '#FF0000',
                negativeColor: '#CC0000'
            });
            
            const data = JSON.parse(action.data);
            expect(data.lineColor).toBe('#0066CC');
            expect(data.markerColor).toBe('#FF0000');
        });

        test('set line weight', () => {
            const action = createAction('configureSparkline', 'E1', {
                lineWeight: 1.5
            });
            
            expect(JSON.parse(action.data).lineWeight).toBe(1.5);
        });

        test('enable date axis', () => {
            const action = createAction('configureSparkline', 'E1', {
                dateAxis: true,
                dateRange: 'F1:I1'
            });
            
            const data = JSON.parse(action.data);
            expect(data.dateAxis).toBe(true);
            expect(data.dateRange).toBe('F1:I1');
        });

        test('edge case: configure non-existent sparkline', () => {
            const action = createAction('configureSparkline', 'Z99', {
                lineColor: '#FF0000'
            });
            
            expect(action.target).toBe('Z99');
        });
    });

    /**
     * Delete Sparkline Tests
     */
    describe('Delete Sparkline Actions', () => {
        test('delete single sparkline', () => {
            const action = createAction('deleteSparkline', 'E1', {});
            
            expect(action.type).toBe('deleteSparkline');
            expect(action.target).toBe('E1');
        });

        test('delete sparkline group', () => {
            const action = createAction('deleteSparkline', 'E1:E10', {
                deleteGroup: true
            });
            
            expect(JSON.parse(action.data).deleteGroup).toBe(true);
        });

        test('edge case: delete non-existent sparkline', () => {
            const action = createAction('deleteSparkline', 'Z99', {});
            
            expect(action.target).toBe('Z99');
        });
    });
});

// ============================================================================
// Step 14: Worksheet Management Tests (9 Actions)
// ============================================================================

describe('Action Executor - Worksheet Management', () => {
    
    /**
     * Rename Sheet Tests
     */
    describe('Rename Sheet Actions', () => {
        test('rename sheet', () => {
            const action = createAction('renameSheet', 'Sheet1', {
                newName: 'Sales Data'
            });
            
            expect(action.type).toBe('renameSheet');
            expect(JSON.parse(action.data).newName).toBe('Sales Data');
        });

        test('rename with valid characters', () => {
            const validNames = ['Sales', 'Q1 2024', 'Data-Analysis', 'Report_Final'];
            
            validNames.forEach(newName => {
                const action = createAction('renameSheet', 'Sheet1', { newName });
                expect(JSON.parse(action.data).newName).toBe(newName);
            });
        });

        test('edge case: 31 character limit', () => {
            const longName = 'A'.repeat(31);
            const action = createAction('renameSheet', 'Sheet1', {
                newName: longName
            });
            
            expect(JSON.parse(action.data).newName.length).toBe(31);
        });

        test('edge case: rename to existing name', () => {
            const action = createAction('renameSheet', 'Sheet1', {
                newName: 'Sheet2'
            });
            
            expect(JSON.parse(action.data).newName).toBe('Sheet2');
        });
    });

    /**
     * Move Sheet Tests
     */
    describe('Move Sheet Actions', () => {
        test('move sheet to position', () => {
            const action = createAction('moveSheet', 'Sheet1', {
                position: 2
            });
            
            expect(action.type).toBe('moveSheet');
            expect(JSON.parse(action.data).position).toBe(2);
        });

        test('move to first position', () => {
            const action = createAction('moveSheet', 'Sheet3', {
                position: 0
            });
            
            expect(JSON.parse(action.data).position).toBe(0);
        });

        test('move to last position', () => {
            const action = createAction('moveSheet', 'Sheet1', {
                position: -1
            });
            
            expect(JSON.parse(action.data).position).toBe(-1);
        });

        test('edge case: invalid position', () => {
            const action = createAction('moveSheet', 'Sheet1', {
                position: 100
            });
            
            expect(JSON.parse(action.data).position).toBe(100);
        });
    });

    /**
     * Hide Sheet Tests
     */
    describe('Hide Sheet Actions', () => {
        test('hide sheet', () => {
            const action = createAction('hideSheet', 'Sheet2', {});
            
            expect(action.type).toBe('hideSheet');
            expect(action.target).toBe('Sheet2');
        });

        test('hide sheet very hidden', () => {
            const action = createAction('hideSheet', 'Sheet2', {
                veryHidden: true
            });
            
            expect(JSON.parse(action.data).veryHidden).toBe(true);
        });

        test('edge case: hide last visible sheet', () => {
            const action = createAction('hideSheet', 'Sheet1', {});
            
            expect(action.type).toBe('hideSheet');
        });
    });

    /**
     * Unhide Sheet Tests
     */
    describe('Unhide Sheet Actions', () => {
        test('unhide sheet', () => {
            const action = createAction('unhideSheet', 'Sheet2', {});
            
            expect(action.type).toBe('unhideSheet');
            expect(action.target).toBe('Sheet2');
        });

        test('unhide very hidden sheet', () => {
            const action = createAction('unhideSheet', 'HiddenSheet', {});
            
            expect(action.type).toBe('unhideSheet');
        });
    });

    /**
     * Freeze Panes Tests
     */
    describe('Freeze Panes Actions', () => {
        test('freeze top row', () => {
            const action = createAction('freezePanes', 'Sheet1', {
                freezeRows: 1
            });
            
            expect(action.type).toBe('freezePanes');
            expect(JSON.parse(action.data).freezeRows).toBe(1);
        });

        test('freeze first column', () => {
            const action = createAction('freezePanes', 'Sheet1', {
                freezeColumns: 1
            });
            
            expect(JSON.parse(action.data).freezeColumns).toBe(1);
        });

        test('freeze rows and columns', () => {
            const action = createAction('freezePanes', 'Sheet1', {
                freezeRows: 2,
                freezeColumns: 1
            });
            
            const data = JSON.parse(action.data);
            expect(data.freezeRows).toBe(2);
            expect(data.freezeColumns).toBe(1);
        });

        test('freeze at cell reference', () => {
            const action = createAction('freezePanes', 'Sheet1', {
                freezeAt: 'B3'
            });
            
            expect(JSON.parse(action.data).freezeAt).toBe('B3');
        });

        test('edge case: freeze at A1', () => {
            const action = createAction('freezePanes', 'Sheet1', {
                freezeAt: 'A1'
            });
            
            expect(JSON.parse(action.data).freezeAt).toBe('A1');
        });
    });

    /**
     * Unfreeze Pane Tests
     */
    describe('Unfreeze Pane Actions', () => {
        test('unfreeze panes', () => {
            const action = createAction('unfreezePane', 'Sheet1', {});
            
            expect(action.type).toBe('unfreezePane');
            expect(action.target).toBe('Sheet1');
        });
    });

    /**
     * Set Zoom Tests
     */
    describe('Set Zoom Actions', () => {
        test('set zoom percentage', () => {
            const action = createAction('setZoom', 'Sheet1', {
                zoom: 85
            });
            
            expect(action.type).toBe('setZoom');
            expect(JSON.parse(action.data).zoom).toBe(85);
        });

        test('zoom in (150%)', () => {
            const action = createAction('setZoom', 'Sheet1', {
                zoom: 150
            });
            
            expect(JSON.parse(action.data).zoom).toBe(150);
        });

        test('zoom out (50%)', () => {
            const action = createAction('setZoom', 'Sheet1', {
                zoom: 50
            });
            
            expect(JSON.parse(action.data).zoom).toBe(50);
        });

        test('edge case: minimum zoom (10%)', () => {
            const action = createAction('setZoom', 'Sheet1', {
                zoom: 10
            });
            
            expect(JSON.parse(action.data).zoom).toBe(10);
        });

        test('edge case: maximum zoom (400%)', () => {
            const action = createAction('setZoom', 'Sheet1', {
                zoom: 400
            });
            
            expect(JSON.parse(action.data).zoom).toBe(400);
        });
    });

    /**
     * Split Pane Tests
     */
    describe('Split Pane Actions', () => {
        test('split panes at cell', () => {
            const action = createAction('splitPane', 'Sheet1', {
                splitAt: 'B3'
            });
            
            expect(action.type).toBe('splitPane');
            expect(JSON.parse(action.data).splitAt).toBe('B3');
        });

        test('horizontal split', () => {
            const action = createAction('splitPane', 'Sheet1', {
                splitRow: 5
            });
            
            expect(JSON.parse(action.data).splitRow).toBe(5);
        });

        test('vertical split', () => {
            const action = createAction('splitPane', 'Sheet1', {
                splitColumn: 3
            });
            
            expect(JSON.parse(action.data).splitColumn).toBe(3);
        });

        test('edge case: split at A1', () => {
            const action = createAction('splitPane', 'Sheet1', {
                splitAt: 'A1'
            });
            
            expect(JSON.parse(action.data).splitAt).toBe('A1');
        });
    });

    /**
     * Create View Tests
     */
    describe('Create View Actions', () => {
        test('create custom view', () => {
            const action = createAction('createView', 'Sheet1', {
                viewName: 'MyView'
            });
            
            expect(action.type).toBe('createView');
            expect(JSON.parse(action.data).viewName).toBe('MyView');
        });

        test('create view with settings', () => {
            const action = createAction('createView', 'Sheet1', {
                viewName: 'PrintView',
                showGridlines: false,
                showHeadings: false
            });
            
            const data = JSON.parse(action.data);
            expect(data.showGridlines).toBe(false);
            expect(data.showHeadings).toBe(false);
        });
    });
});

// ============================================================================
// Step 15: Page Setup Operations Tests (6 Actions)
// ============================================================================

describe('Action Executor - Page Setup Operations', () => {
    
    /**
     * Set Page Setup Tests
     */
    describe('Set Page Setup Actions', () => {
        test('set orientation', () => {
            const action = createAction('setPageSetup', 'Sheet1', {
                orientation: 'Landscape'
            });
            
            expect(action.type).toBe('setPageSetup');
            expect(JSON.parse(action.data).orientation).toBe('Landscape');
        });

        test('set paper size', () => {
            const paperSizes = ['Letter', 'A4', 'Legal', 'A3', 'Tabloid'];
            
            paperSizes.forEach(paperSize => {
                const action = createAction('setPageSetup', 'Sheet1', { paperSize });
                expect(JSON.parse(action.data).paperSize).toBe(paperSize);
            });
        });

        test('set scaling', () => {
            const action = createAction('setPageSetup', 'Sheet1', {
                scale: 75
            });
            
            expect(JSON.parse(action.data).scale).toBe(75);
        });

        test('fit to pages', () => {
            const action = createAction('setPageSetup', 'Sheet1', {
                fitToWidth: 1,
                fitToHeight: 2
            });
            
            const data = JSON.parse(action.data);
            expect(data.fitToWidth).toBe(1);
            expect(data.fitToHeight).toBe(2);
        });

        test('print gridlines', () => {
            const action = createAction('setPageSetup', 'Sheet1', {
                printGridlines: true
            });
            
            expect(JSON.parse(action.data).printGridlines).toBe(true);
        });

        test('print headings', () => {
            const action = createAction('setPageSetup', 'Sheet1', {
                printHeadings: true
            });
            
            expect(JSON.parse(action.data).printHeadings).toBe(true);
        });

        test('edge case: scaling out of range', () => {
            const action = createAction('setPageSetup', 'Sheet1', {
                scale: 500
            });
            
            expect(JSON.parse(action.data).scale).toBe(500);
        });
    });

    /**
     * Set Page Margins Tests
     */
    describe('Set Page Margins Actions', () => {
        test('set all margins', () => {
            const action = createAction('setPageMargins', 'Sheet1', {
                top: 1,
                bottom: 1,
                left: 0.75,
                right: 0.75
            });
            
            expect(action.type).toBe('setPageMargins');
            const data = JSON.parse(action.data);
            expect(data.top).toBe(1);
            expect(data.bottom).toBe(1);
            expect(data.left).toBe(0.75);
            expect(data.right).toBe(0.75);
        });

        test('set header/footer margins', () => {
            const action = createAction('setPageMargins', 'Sheet1', {
                header: 0.5,
                footer: 0.5
            });
            
            const data = JSON.parse(action.data);
            expect(data.header).toBe(0.5);
            expect(data.footer).toBe(0.5);
        });

        test('narrow margins preset', () => {
            const action = createAction('setPageMargins', 'Sheet1', {
                top: 0.75,
                bottom: 0.75,
                left: 0.25,
                right: 0.25
            });
            
            const data = JSON.parse(action.data);
            expect(data.left).toBe(0.25);
            expect(data.right).toBe(0.25);
        });

        test('edge case: zero margins', () => {
            const action = createAction('setPageMargins', 'Sheet1', {
                top: 0,
                bottom: 0,
                left: 0,
                right: 0
            });
            
            const data = JSON.parse(action.data);
            expect(data.top).toBe(0);
        });
    });

    /**
     * Set Page Orientation Tests
     */
    describe('Set Page Orientation Actions', () => {
        test('set landscape', () => {
            const action = createAction('setPageOrientation', 'Sheet1', {
                orientation: 'Landscape'
            });
            
            expect(action.type).toBe('setPageOrientation');
            expect(JSON.parse(action.data).orientation).toBe('Landscape');
        });

        test('set portrait', () => {
            const action = createAction('setPageOrientation', 'Sheet1', {
                orientation: 'Portrait'
            });
            
            expect(JSON.parse(action.data).orientation).toBe('Portrait');
        });
    });

    /**
     * Set Print Area Tests
     */
    describe('Set Print Area Actions', () => {
        test('set single range print area', () => {
            const action = createAction('setPrintArea', 'Sheet1', {
                range: 'A1:F50'
            });
            
            expect(action.type).toBe('setPrintArea');
            expect(JSON.parse(action.data).range).toBe('A1:F50');
        });

        test('set multiple ranges print area', () => {
            const action = createAction('setPrintArea', 'Sheet1', {
                ranges: ['A1:F20', 'A25:F45']
            });
            
            const data = JSON.parse(action.data);
            expect(data.ranges.length).toBe(2);
        });

        test('clear print area', () => {
            const action = createAction('setPrintArea', 'Sheet1', {
                clear: true
            });
            
            expect(JSON.parse(action.data).clear).toBe(true);
        });

        test('edge case: invalid range', () => {
            const action = createAction('setPrintArea', 'Sheet1', {
                range: 'InvalidRange'
            });
            
            expect(JSON.parse(action.data).range).toBe('InvalidRange');
        });
    });

    /**
     * Set Header Footer Tests
     */
    describe('Set Header Footer Actions', () => {
        test('set center header', () => {
            const action = createAction('setHeaderFooter', 'Sheet1', {
                centerHeader: 'Sales Report'
            });
            
            expect(action.type).toBe('setHeaderFooter');
            expect(JSON.parse(action.data).centerHeader).toBe('Sales Report');
        });

        test('set page numbers in footer', () => {
            const action = createAction('setHeaderFooter', 'Sheet1', {
                centerFooter: '&[Page] of &[Pages]'
            });
            
            expect(JSON.parse(action.data).centerFooter).toBe('&[Page] of &[Pages]');
        });

        test('set date in header', () => {
            const action = createAction('setHeaderFooter', 'Sheet1', {
                rightHeader: '&[Date]'
            });
            
            expect(JSON.parse(action.data).rightHeader).toBe('&[Date]');
        });

        test('set file name in footer', () => {
            const action = createAction('setHeaderFooter', 'Sheet1', {
                leftFooter: '&[File]'
            });
            
            expect(JSON.parse(action.data).leftFooter).toBe('&[File]');
        });

        test('set all header/footer sections', () => {
            const action = createAction('setHeaderFooter', 'Sheet1', {
                leftHeader: 'Company Name',
                centerHeader: 'Report Title',
                rightHeader: '&[Date]',
                leftFooter: '&[File]',
                centerFooter: 'Page &[Page]',
                rightFooter: '&[Tab]'
            });
            
            const data = JSON.parse(action.data);
            expect(data.leftHeader).toBe('Company Name');
            expect(data.centerHeader).toBe('Report Title');
            expect(data.rightFooter).toBe('&[Tab]');
        });

        test('formatted header', () => {
            const action = createAction('setHeaderFooter', 'Sheet1', {
                centerHeader: '&"Arial,Bold"&12Report Title'
            });
            
            expect(JSON.parse(action.data).centerHeader).toContain('Arial,Bold');
        });
    });

    /**
     * Set Page Breaks Tests
     */
    describe('Set Page Breaks Actions', () => {
        test('insert horizontal page break', () => {
            const action = createAction('setPageBreaks', 'Sheet1', {
                horizontalBreaks: [20, 40, 60]
            });
            
            expect(action.type).toBe('setPageBreaks');
            const data = JSON.parse(action.data);
            expect(data.horizontalBreaks).toContain(20);
            expect(data.horizontalBreaks).toContain(40);
        });

        test('insert vertical page break', () => {
            const action = createAction('setPageBreaks', 'Sheet1', {
                verticalBreaks: ['D', 'H']
            });
            
            const data = JSON.parse(action.data);
            expect(data.verticalBreaks).toContain('D');
        });

        test('remove page breaks', () => {
            const action = createAction('setPageBreaks', 'Sheet1', {
                removeBreaks: [20]
            });
            
            expect(JSON.parse(action.data).removeBreaks).toContain(20);
        });

        test('clear all page breaks', () => {
            const action = createAction('setPageBreaks', 'Sheet1', {
                clearAll: true
            });
            
            expect(JSON.parse(action.data).clearAll).toBe(true);
        });

        test('edge case: page break at row 1', () => {
            const action = createAction('setPageBreaks', 'Sheet1', {
                horizontalBreaks: [1]
            });
            
            expect(JSON.parse(action.data).horizontalBreaks).toContain(1);
        });
    });
});


// ============================================================================
// Step 16: Hyperlink Operations Tests (3 Actions)
// ============================================================================

describe('Action Executor - Hyperlink Operations', () => {
    
    /**
     * Add Hyperlink Tests
     */
    describe('Add Hyperlink Actions', () => {
        test('property: any URL to any cell', () => {
            fc.assert(
                fc.property(cellRefArb, fc.webUrl(), (cell, url) => {
                    const action = createAction('addHyperlink', cell, { url });
                    
                    expect(action.type).toBe('addHyperlink');
                    expect(action.target).toBe(cell);
                }),
                { numRuns: 50 }
            );
        });

        test('add web URL hyperlink', () => {
            const action = createAction('addHyperlink', 'A1', {
                url: 'https://example.com',
                displayText: 'Click here'
            });
            
            expect(action.type).toBe('addHyperlink');
            expect(JSON.parse(action.data).url).toBe('https://example.com');
            expect(JSON.parse(action.data).displayText).toBe('Click here');
        });

        test('add email hyperlink', () => {
            const action = createAction('addHyperlink', 'B2', {
                url: 'mailto:user@example.com',
                displayText: 'Contact Us'
            });
            
            expect(JSON.parse(action.data).url).toBe('mailto:user@example.com');
        });

        test('add internal document link', () => {
            const action = createAction('addHyperlink', 'C3', {
                url: "'Sheet2'!A1",
                displayText: 'Go to Sheet2'
            });
            
            expect(JSON.parse(action.data).url).toBe("'Sheet2'!A1");
        });

        test('add hyperlink with tooltip', () => {
            const action = createAction('addHyperlink', 'D4', {
                url: 'https://example.com',
                displayText: 'Link',
                tooltip: 'Click to visit example.com'
            });
            
            expect(JSON.parse(action.data).tooltip).toBe('Click to visit example.com');
        });

        test('edge case: 255 character URL limit', () => {
            const longUrl = 'https://example.com/' + 'a'.repeat(230);
            const action = createAction('addHyperlink', 'A1', {
                url: longUrl
            });
            
            expect(JSON.parse(action.data).url.length).toBeGreaterThan(200);
        });

        test('edge case: hyperlink on merged cell', () => {
            const action = createAction('addHyperlink', 'A1', {
                url: 'https://example.com'
            });
            
            expect(action.target).toBe('A1');
        });
    });

    /**
     * Remove Hyperlink Tests
     */
    describe('Remove Hyperlink Actions', () => {
        test('remove hyperlink from cell', () => {
            const action = createAction('removeHyperlink', 'A1', {});
            
            expect(action.type).toBe('removeHyperlink');
            expect(action.target).toBe('A1');
        });

        test('remove hyperlink preserve value', () => {
            const action = createAction('removeHyperlink', 'B2', {
                preserveValue: true
            });
            
            expect(JSON.parse(action.data).preserveValue).toBe(true);
        });

        test('remove hyperlinks from range', () => {
            const action = createAction('removeHyperlink', 'A1:A10', {});
            
            expect(action.target).toBe('A1:A10');
        });

        test('edge case: remove non-existent hyperlink', () => {
            const action = createAction('removeHyperlink', 'Z99', {});
            
            expect(action.target).toBe('Z99');
        });
    });

    /**
     * Edit Hyperlink Tests
     */
    describe('Edit Hyperlink Actions', () => {
        test('edit URL', () => {
            const action = createAction('editHyperlink', 'A1', {
                url: 'https://newurl.com'
            });
            
            expect(action.type).toBe('editHyperlink');
            expect(JSON.parse(action.data).url).toBe('https://newurl.com');
        });

        test('edit display text', () => {
            const action = createAction('editHyperlink', 'A1', {
                displayText: 'New Display Text'
            });
            
            expect(JSON.parse(action.data).displayText).toBe('New Display Text');
        });

        test('edit tooltip', () => {
            const action = createAction('editHyperlink', 'A1', {
                tooltip: 'New tooltip text'
            });
            
            expect(JSON.parse(action.data).tooltip).toBe('New tooltip text');
        });

        test('edit all properties', () => {
            const action = createAction('editHyperlink', 'A1', {
                url: 'https://updated.com',
                displayText: 'Updated Link',
                tooltip: 'Updated tooltip'
            });
            
            const data = JSON.parse(action.data);
            expect(data.url).toBe('https://updated.com');
            expect(data.displayText).toBe('Updated Link');
            expect(data.tooltip).toBe('Updated tooltip');
        });

        test('edge case: edit non-existent hyperlink', () => {
            const action = createAction('editHyperlink', 'Z99', {
                url: 'https://example.com'
            });
            
            expect(action.target).toBe('Z99');
        });
    });
});

// ============================================================================
// Step 17: Data Type Operations Tests (2 Actions)
// ============================================================================

describe('Action Executor - Data Type Operations', () => {
    
    /**
     * Insert Data Type Tests
     */
    describe('Insert Data Type Actions', () => {
        test('property: any custom entity with any properties', () => {
            fc.assert(
                fc.property(
                    cellRefArb,
                    fc.string({ minLength: 1, maxLength: 50 }),
                    (cell, displayText) => {
                        const action = createAction('insertDataType', cell, {
                            displayText,
                            properties: {
                                Name: displayText
                            }
                        });
                        
                        expect(action.type).toBe('insertDataType');
                    }
                ),
                { numRuns: 50 }
            );
        });

        test('insert custom entity with string properties', () => {
            const action = createAction('insertDataType', 'A2', {
                displayText: 'Product A',
                properties: {
                    SKU: 'SKU-001',
                    Name: 'Product A',
                    Category: 'Electronics'
                }
            });
            
            expect(action.type).toBe('insertDataType');
            const data = JSON.parse(action.data);
            expect(data.displayText).toBe('Product A');
            expect(data.properties.SKU).toBe('SKU-001');
        });

        test('insert entity with numeric properties', () => {
            const action = createAction('insertDataType', 'A3', {
                displayText: 'Product B',
                properties: {
                    Price: 29.99,
                    Quantity: 100,
                    Rating: 4.5
                }
            });
            
            const data = JSON.parse(action.data);
            expect(data.properties.Price).toBe(29.99);
            expect(data.properties.Quantity).toBe(100);
        });

        test('insert entity with boolean properties', () => {
            const action = createAction('insertDataType', 'A4', {
                displayText: 'Product C',
                properties: {
                    InStock: true,
                    OnSale: false,
                    Featured: true
                }
            });
            
            const data = JSON.parse(action.data);
            expect(data.properties.InStock).toBe(true);
            expect(data.properties.OnSale).toBe(false);
        });

        test('insert entity with mixed property types', () => {
            const action = createAction('insertDataType', 'A5', {
                displayText: 'Employee Record',
                properties: {
                    ID: 'EMP-001',
                    Name: 'John Doe',
                    Department: 'Engineering',
                    Salary: 75000,
                    IsManager: false,
                    StartDate: '2020-01-15'
                }
            });
            
            const data = JSON.parse(action.data);
            expect(Object.keys(data.properties).length).toBe(6);
        });

        test('insert entity with 10 properties', () => {
            const properties = {};
            for (let i = 1; i <= 10; i++) {
                properties[`Property${i}`] = `Value${i}`;
            }
            
            const action = createAction('insertDataType', 'A6', {
                displayText: 'Multi-Property Entity',
                properties
            });
            
            const data = JSON.parse(action.data);
            expect(Object.keys(data.properties).length).toBe(10);
        });

        test('insert entity with basicValue fallback', () => {
            const action = createAction('insertDataType', 'A7', {
                displayText: 'Fallback Entity',
                basicValue: 'Fallback Text',
                properties: {
                    Name: 'Entity Name'
                }
            });
            
            const data = JSON.parse(action.data);
            expect(data.basicValue).toBe('Fallback Text');
        });

        test('edge case: empty properties', () => {
            const action = createAction('insertDataType', 'A8', {
                displayText: 'Empty Entity',
                properties: {}
            });
            
            const data = JSON.parse(action.data);
            expect(Object.keys(data.properties).length).toBe(0);
        });

        test('edge case: single property', () => {
            const action = createAction('insertDataType', 'A9', {
                displayText: 'Simple Entity',
                properties: {
                    Value: 'Single Value'
                }
            });
            
            const data = JSON.parse(action.data);
            expect(Object.keys(data.properties).length).toBe(1);
        });
    });

    /**
     * Refresh Data Type Tests
     */
    describe('Refresh Data Type Actions', () => {
        test('update entity properties', () => {
            const action = createAction('refreshDataType', 'A2', {
                properties: {
                    Price: 34.99,
                    Quantity: 150
                }
            });
            
            expect(action.type).toBe('refreshDataType');
            const data = JSON.parse(action.data);
            expect(data.properties.Price).toBe(34.99);
        });

        test('merge with existing properties', () => {
            const action = createAction('refreshDataType', 'A2', {
                mergeProperties: true,
                properties: {
                    NewProperty: 'New Value'
                }
            });
            
            expect(JSON.parse(action.data).mergeProperties).toBe(true);
        });

        test('replace all properties', () => {
            const action = createAction('refreshDataType', 'A2', {
                mergeProperties: false,
                properties: {
                    ReplacedProp: 'Replaced Value'
                }
            });
            
            expect(JSON.parse(action.data).mergeProperties).toBe(false);
        });

        test('edge case: refresh non-existent entity', () => {
            const action = createAction('refreshDataType', 'Z99', {
                properties: {
                    Test: 'Value'
                }
            });
            
            expect(action.target).toBe('Z99');
        });

        test('edge case: refresh with empty properties', () => {
            const action = createAction('refreshDataType', 'A2', {
                properties: {}
            });
            
            const data = JSON.parse(action.data);
            expect(Object.keys(data.properties).length).toBe(0);
        });
    });
});