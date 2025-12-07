/**
 * Performance Tests for Action Executor
 * Tests performance with large datasets and complex operations using the real executor
 * 
 * These tests import and call the real executeAction() function to measure
 * actual performance characteristics of the production code.
 */

import { executeAction } from './action-executor.js';

// ============================================================================
// Mock Excel Global Object
// ============================================================================

global.Excel = {
    ChartType: {
        columnClustered: 'ColumnClustered',
        line: 'Line',
        pie: 'Pie',
        bar: 'BarClustered',
        area: 'Area',
        scatter: 'XYScatter',
        combo: 'ComboColumnLine'
    },
    ConditionalFormatType: {
        colorScale: 'ColorScale',
        dataBar: 'DataBar',
        iconSet: 'IconSet',
        topBottom: 'TopBottom',
        presetCriteria: 'PresetCriteria',
        custom: 'Custom'
    },
    DataValidationType: {
        wholeNumber: 'WholeNumber',
        decimal: 'Decimal',
        list: 'List',
        date: 'Date',
        time: 'Time',
        textLength: 'TextLength',
        custom: 'Custom'
    },
    AutoFillType: {
        fillDefault: 'FillDefault',
        fillCopy: 'FillCopy',
        fillSeries: 'FillSeries',
        fillFormats: 'FillFormats',
        fillValues: 'FillValues'
    },
    run: jest.fn((callback) => {
        const ctx = { sync: jest.fn(() => Promise.resolve()) };
        return callback(ctx);
    })
};

// ============================================================================
// Performance Test Utilities
// ============================================================================

/**
 * Measures execution time of an async function
 * @param {Function} fn - Async function to measure
 * @returns {Object} Result with value and executionTime
 */
async function measureExecutionTime(fn) {
    const startTime = performance.now();
    const result = await fn();
    const endTime = performance.now();
    return {
        result,
        executionTime: endTime - startTime
    };
}

/**
 * Generates a large dataset for testing
 * @param {number} rows - Number of rows
 * @param {number} cols - Number of columns
 * @returns {Array} 2D array of test data
 */
function generateLargeDataset(rows, cols) {
    const data = [];
    for (let i = 0; i < rows; i++) {
        const row = [];
        for (let j = 0; j < cols; j++) {
            if (j % 3 === 0) {
                row.push(`Text_${i}_${j}`);
            } else if (j % 3 === 1) {
                row.push(Math.random() * 10000);
            } else {
                row.push(i % 2 === 0);
            }
        }
        data.push(row);
    }
    return data;
}

/**
 * Captures memory usage snapshot
 */
function memorySnapshot() {
    if (typeof process !== 'undefined' && process.memoryUsage) {
        const usage = process.memoryUsage();
        return {
            heapUsed: usage.heapUsed,
            heapTotal: usage.heapTotal,
            external: usage.external,
            rss: usage.rss
        };
    }
    return { heapUsed: 0, heapTotal: 0, external: 0, rss: 0 };
}

/**
 * Creates a performance-oriented mock context with realistic state tracking
 */
function createPerformanceMockContext() {
    const state = {
        worksheets: new Map([['Sheet1', createSheetState('Sheet1')]]),
        namedRanges: new Map(),
        ranges: new Map()
    };

    function createSheetState(name) {
        return {
            name,
            tables: new Map(),
            pivotTables: new Map(),
            charts: new Map(),
            slicers: new Map(),
            conditionalFormats: []
        };
    }

    const context = {
        sync: jest.fn(() => Promise.resolve()),
        workbook: {
            worksheets: {
                getActiveWorksheet: jest.fn(() => createPerformanceMockWorksheet(state.worksheets.get('Sheet1'), state)),
                getItem: jest.fn((name) => {
                    if (!state.worksheets.has(name)) {
                        state.worksheets.set(name, createSheetState(name));
                    }
                    return createPerformanceMockWorksheet(state.worksheets.get(name), state);
                }),
                getItemOrNullObject: jest.fn((name) => {
                    if (!state.worksheets.has(name)) {
                        return { isNullObject: true, load: jest.fn().mockReturnThis() };
                    }
                    return createPerformanceMockWorksheet(state.worksheets.get(name), state);
                }),
                add: jest.fn((name) => {
                    state.worksheets.set(name, createSheetState(name));
                    return createPerformanceMockWorksheet(state.worksheets.get(name), state);
                })
            },
            names: {
                add: jest.fn((name, ref) => {
                    state.namedRanges.set(name, { name, value: ref });
                    return { name, value: ref, load: jest.fn().mockReturnThis() };
                }),
                getItem: jest.fn((name) => state.namedRanges.get(name))
            },
            tables: {
                getItem: jest.fn((name) => {
                    for (const sheet of state.worksheets.values()) {
                        if (sheet.tables.has(name)) return sheet.tables.get(name);
                    }
                    throw new Error(`Table '${name}' not found`);
                })
            }
        },
        _state: state
    };
    return context;
}

/**
 * Creates a performance mock worksheet with state tracking
 */
function createPerformanceMockWorksheet(sheetState, globalState) {
    return {
        name: sheetState.name,
        getRange: jest.fn((address) => createPerformanceMockRange(address, globalState, sheetState)),
        getUsedRange: jest.fn(() => createPerformanceMockRange('A1:Z100', globalState, sheetState)),
        tables: {
            add: jest.fn((address, hasHeaders) => {
                const name = `Table${sheetState.tables.size + 1}`;
                const table = { name, address, style: 'TableStyleMedium2', load: jest.fn().mockReturnThis() };
                sheetState.tables.set(name, table);
                return table;
            }),
            getItem: jest.fn((name) => sheetState.tables.get(name))
        },
        charts: {
            add: jest.fn((type, source, seriesBy) => {
                const name = `Chart${sheetState.charts.size + 1}`;
                const chart = { name, chartType: type, load: jest.fn().mockReturnThis(), setPosition: jest.fn() };
                sheetState.charts.set(name, chart);
                return chart;
            })
        },
        pivotTables: {
            add: jest.fn((name, source, dest) => {
                const pivot = { name, load: jest.fn().mockReturnThis(), refresh: jest.fn() };
                sheetState.pivotTables.set(name, pivot);
                return pivot;
            })
        },
        load: jest.fn().mockReturnThis()
    };
}

/**
 * Creates a performance mock range with state tracking
 */
function createPerformanceMockRange(address, globalState, sheetState) {
    const rangeKey = `${sheetState.name}!${address}`;
    if (!globalState.ranges.has(rangeKey)) {
        globalState.ranges.set(rangeKey, {
            values: [['']],
            formulas: [['']],
            format: {
                font: { bold: false, italic: false, color: '#000000' },
                fill: { color: '#FFFFFF' },
                borders: { getItem: jest.fn(() => ({ style: 'None' })) }
            },
            conditionalFormats: { add: jest.fn((type) => ({ type })), clearAll: jest.fn(), items: [] },
            dataValidation: { rule: null, clear: jest.fn() }
        });
    }
    const rangeState = globalState.ranges.get(rangeKey);

    return {
        address,
        get values() { return rangeState.values; },
        set values(v) { rangeState.values = v; },
        get formulas() { return rangeState.formulas; },
        set formulas(f) { rangeState.formulas = f; },
        format: rangeState.format,
        conditionalFormats: rangeState.conditionalFormats,
        dataValidation: rangeState.dataValidation,
        load: jest.fn().mockReturnThis(),
        clear: jest.fn(),
        merge: jest.fn(),
        sort: { apply: jest.fn() },
        rowCount: 1,
        columnCount: 1
    };
}

/**
 * Creates an action object
 */
function createAction(type, target, data = {}) {
    return {
        type,
        target,
        data: typeof data === 'string' ? data : JSON.stringify(data)
    };
}

/**
 * Executes action through the REAL executor with timing
 * Calls the production action-executor.js code to measure actual performance
 */
async function executeActionWithTiming(ctx, action) {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    return measureExecutionTime(async () => {
        await executeAction(ctx, sheet, action);
    });
}

// ============================================================================
// Performance Benchmarks Configuration
// ============================================================================

const BENCHMARKS = {
    insertValues: {
        '10K rows': { rows: 10000, cols: 5, targetTime: 2000 },
        '50K rows': { rows: 50000, cols: 5, targetTime: 5000 },
        '100K rows': { rows: 100000, cols: 5, targetTime: 10000 }
    },
    formulas: {
        '100 cells': { count: 100, targetTime: 3000 },
        '500 cells': { count: 500, targetTime: 8000 },
        '1K cells': { count: 1000, targetTime: 15000 }
    },
    tables: {
        '10K rows': { rows: 10000, targetTime: 2000 },
        '50K rows': { rows: 50000, targetTime: 5000 }
    },
    charts: {
        '100 points': { points: 100, targetTime: 1000 },
        '1K points': { points: 1000, targetTime: 3000 },
        '5K points': { points: 5000, targetTime: 8000 }
    },
    conditionalFormat: {
        '1K cells': { cells: 1000, targetTime: 2000 },
        '10K cells': { cells: 10000, targetTime: 4000 },
        '50K cells': { cells: 50000, targetTime: 10000 }
    },
    sparklines: {
        '50 sparklines': { count: 50, targetTime: 3000 },
        '100 sparklines': { count: 100, targetTime: 5000 },
        '200 sparklines': { count: 200, targetTime: 10000 }
    },
    pivotTables: {
        '10K source rows': { rows: 10000, targetTime: 5000 },
        '50K source rows': { rows: 50000, targetTime: 15000 },
        '100K source rows': { rows: 100000, targetTime: 30000 }
    }
};

// ============================================================================
// Large Dataset Tests with Real Executor
// ============================================================================

describe('Performance Tests - Large Datasets', () => {
    
    describe('Insert Values Performance', () => {
        test('10K rows insert via executor completes within target time', async () => {
            const ctx = createPerformanceMockContext();
            const data = generateLargeDataset(10000, 5);
            const action = createAction('values', 'A1:E10000', JSON.stringify(data));
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.insertValues['10K rows'].targetTime);
            console.log(`10K rows insert via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('50K rows insert via executor completes within target time', async () => {
            const ctx = createPerformanceMockContext();
            const data = generateLargeDataset(50000, 5);
            const action = createAction('values', 'A1:E50000', JSON.stringify(data));
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.insertValues['50K rows'].targetTime);
            console.log(`50K rows insert via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('100K rows insert via executor completes within target time', async () => {
            const ctx = createPerformanceMockContext();
            const data = generateLargeDataset(100000, 5);
            const action = createAction('values', 'A1:E100000', JSON.stringify(data));
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.insertValues['100K rows'].targetTime);
            console.log(`100K rows insert via executor: ${executionTime.toFixed(2)}ms`);
        });
    });

    describe('Table Creation Performance', () => {
        test('create table from 10K rows via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('createTable', 'A1:E10000', { name: 'PerfTable10K', hasHeaders: true });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.tables['10K rows'].targetTime);
            console.log(`Table from 10K rows via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('create table from 50K rows via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('createTable', 'A1:E50000', { name: 'PerfTable50K', hasHeaders: true });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.tables['50K rows'].targetTime);
            console.log(`Table from 50K rows via executor: ${executionTime.toFixed(2)}ms`);
        });
    });
});

// ============================================================================
// Batch Operation Tests with Real Executor
// ============================================================================

describe('Performance Tests - Batch Operations', () => {
    
    describe('Formula Batch Performance', () => {
        test('apply 100 formulas via executor', async () => {
            const ctx = createPerformanceMockContext();
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            
            const { executionTime } = await measureExecutionTime(async () => {
                for (let i = 1; i <= 100; i++) {
                    const action = createAction('formula', `A${i}`, `=SUM(B${i}:Z${i})`);
                    await executeAction(ctx, sheet, action);
                }
            });
            
            expect(executionTime).toBeLessThan(BENCHMARKS.formulas['100 cells'].targetTime);
            console.log(`100 formulas via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('apply 500 formulas via executor', async () => {
            const ctx = createPerformanceMockContext();
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            
            const { executionTime } = await measureExecutionTime(async () => {
                for (let i = 1; i <= 500; i++) {
                    const action = createAction('formula', `A${i}`, `=SUM(B${i}:Z${i})`);
                    await executeAction(ctx, sheet, action);
                }
            });
            
            expect(executionTime).toBeLessThan(BENCHMARKS.formulas['500 cells'].targetTime);
            console.log(`500 formulas via executor: ${executionTime.toFixed(2)}ms`);
        });
    });

    describe('Chart Batch Performance', () => {
        test('create 10 charts via executor', async () => {
            const ctx = createPerformanceMockContext();
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            
            const { executionTime } = await measureExecutionTime(async () => {
                for (let i = 0; i < 10; i++) {
                    const action = createAction('chart', 'A1:B10', { chartType: 'ColumnClustered', position: `D${i * 20 + 1}` });
                    await executeAction(ctx, sheet, action);
                }
            });
            
            expect(executionTime).toBeLessThan(3000);
            console.log(`10 charts via executor: ${executionTime.toFixed(2)}ms`);
        });
    });
});


// ============================================================================
// Complex Operation Tests with Real Executor
// ============================================================================

describe('Performance Tests - Complex Operations', () => {
    
    describe('Conditional Format Performance', () => {
        test('color scale on 1K cells via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('conditionalFormat', 'A1:A1000', {
                type: 'colorScale',
                colorScale: {
                    minimum: { color: '#FF0000' },
                    maximum: { color: '#00FF00' }
                }
            });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.conditionalFormat['1K cells'].targetTime);
            console.log(`Color scale 1K cells via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('color scale on 10K cells via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('conditionalFormat', 'A1:A10000', {
                type: 'colorScale',
                colorScale: {
                    minimum: { color: '#FF0000' },
                    maximum: { color: '#00FF00' }
                }
            });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.conditionalFormat['10K cells'].targetTime);
            console.log(`Color scale 10K cells via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('color scale on 50K cells via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('conditionalFormat', 'A1:A50000', {
                type: 'colorScale',
                colorScale: {
                    minimum: { color: '#FF0000' },
                    maximum: { color: '#00FF00' }
                }
            });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.conditionalFormat['50K cells'].targetTime);
            console.log(`Color scale 50K cells via executor: ${executionTime.toFixed(2)}ms`);
        });
    });

    describe('PivotTable Performance', () => {
        test('PivotTable from 10K source rows via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('createPivotTable', 'A1:E10000', { name: 'PerfPivot10K', destination: 'G1' });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.pivotTables['10K source rows'].targetTime);
            console.log(`PivotTable 10K rows via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('PivotTable from 50K source rows via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('createPivotTable', 'A1:E50000', { name: 'PerfPivot50K', destination: 'G1' });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.pivotTables['50K source rows'].targetTime);
            console.log(`PivotTable 50K rows via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('PivotTable from 100K source rows via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('createPivotTable', 'A1:E100000', { name: 'PerfPivot100K', destination: 'G1' });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.pivotTables['100K source rows'].targetTime);
            console.log(`PivotTable 100K rows via executor: ${executionTime.toFixed(2)}ms`);
        });
    });

    describe('Chart with Large Dataset', () => {
        test('line chart with 1000 data points via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('chart', 'A1:B1000', { chartType: 'Line', position: 'D1' });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.charts['1K points'].targetTime);
            console.log(`Chart 1K points via executor: ${executionTime.toFixed(2)}ms`);
        });

        test('line chart with 5000 data points via executor', async () => {
            const ctx = createPerformanceMockContext();
            const action = createAction('chart', 'A1:B5000', { chartType: 'Line', position: 'D1' });
            
            const { executionTime } = await executeActionWithTiming(ctx, action);
            
            expect(executionTime).toBeLessThan(BENCHMARKS.charts['5K points'].targetTime);
            console.log(`Chart 5K points via executor: ${executionTime.toFixed(2)}ms`);
        });
    });
});

// ============================================================================
// Memory Tests with Real Executor
// ============================================================================

describe('Performance Tests - Memory Usage', () => {
    
    test('memory usage for 10K row insert via executor', async () => {
        const beforeMemory = memorySnapshot();
        
        const ctx = createPerformanceMockContext();
        const data = generateLargeDataset(10000, 5);
        const action = createAction('values', 'A1:E10000', JSON.stringify(data));
        await executeActionWithTiming(ctx, action);
        
        const afterMemory = memorySnapshot();
        const memoryIncrease = (afterMemory.heapUsed - beforeMemory.heapUsed) / (1024 * 1024);
        
        console.log(`Memory increase for 10K rows via executor: ${memoryIncrease.toFixed(2)}MB`);
        expect(memoryIncrease).toBeLessThan(100);
    });

    test('memory usage for 50K row insert via executor', async () => {
        const beforeMemory = memorySnapshot();
        
        const ctx = createPerformanceMockContext();
        const data = generateLargeDataset(50000, 5);
        const action = createAction('values', 'A1:E50000', JSON.stringify(data));
        await executeActionWithTiming(ctx, action);
        
        const afterMemory = memorySnapshot();
        const memoryIncrease = (afterMemory.heapUsed - beforeMemory.heapUsed) / (1024 * 1024);
        
        console.log(`Memory increase for 50K rows via executor: ${memoryIncrease.toFixed(2)}MB`);
        expect(memoryIncrease).toBeLessThan(500);
    });

    test('no memory leak on repeated operations via executor', async () => {
        const ctx = createPerformanceMockContext();
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        
        const initialMemory = memorySnapshot();
        
        // Perform 100 repeated operations through the executor
        for (let i = 0; i < 100; i++) {
            const action = createAction('values', `A${i + 1}`, JSON.stringify([[`Value${i}`]]));
            await executeAction(ctx, sheet, action);
        }
        
        if (global.gc) {
            global.gc();
        }
        
        const finalMemory = memorySnapshot();
        const memoryIncrease = (finalMemory.heapUsed - initialMemory.heapUsed) / (1024 * 1024);
        
        console.log(`Memory after 100 operations via executor: ${memoryIncrease.toFixed(2)}MB increase`);
        expect(memoryIncrease).toBeLessThan(50);
    });
});

// ============================================================================
// Concurrent Operation Tests with Real Executor
// ============================================================================

describe('Performance Tests - Concurrent Operations', () => {
    
    test('execute 10 actions simultaneously via executor', async () => {
        const ctx = createPerformanceMockContext();
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        
        const { executionTime } = await measureExecutionTime(async () => {
            const promises = [];
            for (let i = 0; i < 10; i++) {
                const action = createAction('values', `A${i + 1}`, JSON.stringify([[`Concurrent${i}`]]));
                promises.push(executeAction(ctx, sheet, action));
            }
            await Promise.all(promises);
        });
        
        console.log(`10 concurrent actions via executor: ${executionTime.toFixed(2)}ms`);
        expect(executionTime).toBeLessThan(2000);
    });

    test('batch operations efficiency via executor', async () => {
        const ctx = createPerformanceMockContext();
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        
        // Single large batch
        const { executionTime: batchTime } = await measureExecutionTime(async () => {
            const data = generateLargeDataset(100, 5);
            const action = createAction('values', 'A1:E100', JSON.stringify(data));
            await executeAction(ctx, sheet, action);
        });
        
        console.log(`Batch insert (100 rows, 1 action): ${batchTime.toFixed(2)}ms`);
        
        // Multiple small operations
        const { executionTime: multiTime } = await measureExecutionTime(async () => {
            for (let i = 0; i < 10; i++) {
                const action = createAction('values', `F${i + 1}`, JSON.stringify([[`Multi${i}`]]));
                await executeAction(ctx, sheet, action);
            }
        });
        
        console.log(`Multi insert (10 rows, 10 actions): ${multiTime.toFixed(2)}ms`);
    });
});


// ============================================================================
// Performance Report Generator
// ============================================================================

describe('Performance Tests - Summary Report', () => {
    
    test('generate performance summary', () => {
        const report = {
            timestamp: new Date().toISOString(),
            benchmarks: BENCHMARKS,
            recommendations: [
                'For datasets >50K rows, consider chunking operations',
                'Limit charts to <20 per sheet for optimal performance',
                'Use batch sync instead of individual syncs',
                'Avoid >100 sparklines per sheet',
                'Consider using tables for auto-expanding data ranges',
                'PivotTables with >100K source rows may be slow',
                'Conditional formatting on >50K cells impacts performance'
            ]
        };
        
        console.log('\n=== Performance Test Summary ===');
        console.log(JSON.stringify(report, null, 2));
        
        expect(report.benchmarks).toBeDefined();
        expect(report.recommendations.length).toBeGreaterThan(0);
    });
});
