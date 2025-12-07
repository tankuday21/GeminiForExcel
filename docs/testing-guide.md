# Testing Guide - Excel AI Copilot

Comprehensive guide for testing the Excel AI Copilot add-in, including test architecture, patterns, and best practices.

---

## Table of Contents

1. [Testing Philosophy](#testing-philosophy)
2. [Test Architecture](#test-architecture)
3. [Running Tests](#running-tests)
4. [Writing Tests](#writing-tests)
5. [Mock Office.js Setup](#mock-officejs-setup)
6. [Property-Based Testing](#property-based-testing)
7. [Test Coverage Goals](#test-coverage-goals)
8. [Integration Testing](#integration-testing)
9. [Performance Testing](#performance-testing)
10. [Debugging Tests](#debugging-tests)
11. [Contributing Tests](#contributing-tests)

---

## Testing Philosophy

### Why Comprehensive Testing Matters

The Excel AI Copilot executes actions directly on user spreadsheets. Bugs can:
- Corrupt user data
- Apply incorrect formulas
- Break existing workbook functionality
- Cause data loss

Comprehensive testing ensures:
- All 87 action handlers work correctly
- Edge cases are handled gracefully
- Performance meets user expectations
- Regressions are caught early

### Property-Based Testing Benefits

We use [fast-check](https://github.com/dubzzz/fast-check) for property-based testing:

1. **Discovers edge cases** - Generates thousands of test inputs automatically
2. **Reduces bias** - Tests inputs you wouldn't think to try
3. **Shrinking** - When a test fails, fast-check finds the minimal failing case
4. **Reproducibility** - Failed tests can be replayed with the same seed

### Test Pyramid

```
        /\
       /  \     E2E Tests (Manual)
      /----\    - Real Excel testing
     /      \   - User acceptance
    /--------\  Integration Tests
   /          \ - Multi-action workflows
  /------------\- State management
 /              \ Unit Tests
/----------------\- Individual handlers
                  - Edge cases
                  - Property tests
```

---

## Test Architecture

### Directory Structure

```
src/taskpane/
├── action-executor.js           # Main file under test
├── action-executor.test.js      # Unit tests (~3000 lines)
├── action-executor.integration.test.js  # Integration tests
├── action-executor.performance.test.js  # Performance tests
├── ai-engine.test.js            # AI engine tests
├── preview.test.js              # Preview component tests
└── history.test.js              # History component tests
```

### Test Categories

| Category | File | Purpose | Run Time |
|----------|------|---------|----------|
| Unit | `*.test.js` | Individual function testing | Fast (<30s) |
| Integration | `*.integration.test.js` | Multi-action workflows | Medium (<2min) |
| Performance | `*.performance.test.js` | Benchmarks, large data | Slow (<5min) |

### Test Organization

Tests are organized by feature category matching the action types:

```javascript
describe('Action Executor - Basic Operations', () => {
    describe('Formula Actions', () => { /* ... */ });
    describe('Values Actions', () => { /* ... */ });
    describe('Format Actions', () => { /* ... */ });
});

describe('Action Executor - Table Operations', () => {
    describe('Create Table Actions', () => { /* ... */ });
    describe('Style Table Actions', () => { /* ... */ });
    // ...
});
```

---

## Running Tests

### Basic Commands

```bash
# Run all tests
npm test

# Run unit tests only
npm run test:unit

# Run integration tests
npm run test:integration

# Run performance tests
npm run test:performance

# Run specific test file
npm run test:executor

# Run all test suites
npm run test:all

# Watch mode (re-run on changes)
npm run test:watch

# Generate coverage report
npm run test:coverage
```

### Test Scripts in package.json

```json
{
  "scripts": {
    "test": "jest",
    "test:unit": "jest --testPathPatterns=\"\\.test\\.js$\" --testPathIgnorePatterns=\"integration|performance\"",
    "test:integration": "jest --testPathPatterns=\"integration\\.test\\.js$\"",
    "test:performance": "jest --testPathPatterns=\"performance\\.test\\.js$\" --maxWorkers=1",
    "test:watch": "jest --watch",
    "test:coverage": "jest --coverage --coverageDirectory=coverage",
    "test:executor": "jest src/taskpane/action-executor.test.js",
    "test:all": "npm run test:unit && npm run test:integration && npm run test:performance"
  }
}
```

### Running Specific Tests

```bash
# Run tests matching a pattern
npx jest -t "Formula Actions"

# Run tests in a specific file
npx jest action-executor.test.js

# Run with verbose output
npx jest --verbose

# Run with coverage for specific file
npx jest --coverage --collectCoverageFrom="src/taskpane/action-executor.js"
```

---

## Writing Tests

### Test Structure

```javascript
describe('Feature Category', () => {
    // Setup before each test
    let ctx;
    
    beforeEach(() => {
        ctx = createMockContext();
    });
    
    describe('Specific Action', () => {
        // Property-based test
        test('property: description of invariant', () => {
            fc.assert(
                fc.property(arbitrary1, arbitrary2, (input1, input2) => {
                    // Test that property holds for all inputs
                    expect(/* condition */).toBe(true);
                }),
                { numRuns: 100 }
            );
        });
        
        // Unit test
        test('specific scenario description', () => {
            const action = createAction('type', 'target', { data });
            // Assertions
            expect(action.type).toBe('type');
        });
        
        // Edge case test
        test('edge case: description', () => {
            // Test boundary conditions
        });
    });
});
```

### Naming Conventions

- **Describe blocks**: Use feature/action names
- **Test names**: Start with what's being tested
- **Property tests**: Prefix with "property:"
- **Edge cases**: Prefix with "edge case:"

```javascript
describe('Action Executor - Table Operations', () => {
    describe('Create Table Actions', () => {
        test('property: any range with headers creates valid table', () => {});
        test('create table with headers', () => {});
        test('create table without headers', () => {});
        test('edge case: single row table', () => {});
        test('edge case: overlapping tables', () => {});
    });
});
```

### Assertion Patterns

```javascript
// Basic assertions
expect(result).toBe(expected);
expect(result).toEqual(expected);  // Deep equality
expect(result).toBeDefined();
expect(result).toBeNull();

// Numeric assertions
expect(value).toBeGreaterThan(0);
expect(value).toBeLessThan(100);
expect(value).toBeCloseTo(3.14, 2);

// String assertions
expect(str).toContain('substring');
expect(str).toMatch(/pattern/);

// Array assertions
expect(arr).toContain(item);
expect(arr).toHaveLength(5);

// Object assertions
expect(obj).toHaveProperty('key');
expect(obj).toMatchObject({ key: 'value' });

// Function assertions
expect(fn).toHaveBeenCalled();
expect(fn).toHaveBeenCalledWith(arg1, arg2);
expect(fn).toHaveBeenCalledTimes(3);

// Async assertions
await expect(asyncFn()).resolves.toBe(value);
await expect(asyncFn()).rejects.toThrow('error');
```

---

## Mock Office.js Setup

### Creating Mock Context

```javascript
function createMockContext() {
    const syncCalls = [];
    
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
            }
        },
        _syncCalls: syncCalls
    };
    
    return context;
}
```

### Creating Mock Worksheet

```javascript
function createMockWorksheet(name = 'Sheet1') {
    return {
        name,
        id: `sheet_${name}`,
        getRange: jest.fn((address) => createMockRange(address)),
        getUsedRange: jest.fn(() => createMockRange('A1:Z100')),
        tables: {
            add: jest.fn(() => createMockTable()),
            getItem: jest.fn(() => createMockTable()),
            items: []
        },
        charts: {
            add: jest.fn(() => createMockChart()),
            items: []
        },
        // ... other properties
        load: jest.fn().mockReturnThis()
    };
}
```

### Creating Mock Range

```javascript
function createMockRange(address = 'A1') {
    return {
        address,
        rowCount: 1,
        columnCount: 1,
        values: [['']],
        formulas: [['']],
        format: {
            font: { bold: false, italic: false, color: '#000000' },
            fill: { color: '#FFFFFF' },
            borders: { getItem: jest.fn(() => ({})) }
        },
        conditionalFormats: {
            add: jest.fn(() => ({})),
            clearAll: jest.fn()
        },
        load: jest.fn().mockReturnThis(),
        clear: jest.fn(),
        merge: jest.fn()
    };
}
```

### Stateful Mocks for Integration Tests

```javascript
function createStatefulMockContext() {
    const state = {
        worksheets: new Map([['Sheet1', { tables: new Map() }]]),
        namedRanges: new Map()
    };
    
    return {
        sync: jest.fn(() => Promise.resolve()),
        workbook: {
            worksheets: {
                getActiveWorksheet: jest.fn(() => {
                    const sheet = state.worksheets.get('Sheet1');
                    return createStatefulWorksheet(sheet, state);
                })
            }
        },
        _state: state  // Expose for assertions
    };
}
```

---

## Property-Based Testing

### Fast-Check Arbitraries

```javascript
const fc = require('fast-check');

// Cell reference (A1, B5, etc.)
const cellRefArb = fc.tuple(
    fc.constantFrom('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'),
    fc.integer({ min: 1, max: 1000 })
).map(([col, row]) => `${col}${row}`);

// Range (A1:D10)
const rangeArb = fc.tuple(cellRefArb, cellRefArb)
    .map(([start, end]) => `${start}:${end}`);

// Hex color (#FF0000)
const hexColorArb = fc.hexaString({ minLength: 6, maxLength: 6 })
    .map(hex => `#${hex.toUpperCase()}`);

// Table name (alphanumeric, starts with letter)
const tableNameArb = fc.stringOf(
    fc.constantFrom(...'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_'),
    { minLength: 1, maxLength: 20 }
).filter(s => /^[A-Za-z]/.test(s));

// Chart type
const chartTypeArb = fc.constantFrom(
    'ColumnClustered', 'BarClustered', 'Line', 'Pie', 'Area', 'Scatter'
);

// 2D values array
const valuesArrayArb = fc.array(
    fc.array(
        fc.oneof(fc.string(), fc.integer(), fc.double(), fc.boolean()),
        { minLength: 1, maxLength: 10 }
    ),
    { minLength: 1, maxLength: 100 }
);
```

### Property Test Examples

```javascript
// Property: Any valid formula applies to any cell
test('property: any valid formula applies to any cell reference', () => {
    fc.assert(
        fc.property(cellRefArb, formulaArb, (target, formula) => {
            const action = createAction('formula', target, formula);
            expect(action.type).toBe('formula');
            expect(action.target).toBe(target);
        }),
        { numRuns: 100 }
    );
});

// Property: Table styles are valid
test('property: all table styles apply correctly', () => {
    fc.assert(
        fc.property(tableNameArb, tableStyleArb, (name, style) => {
            const action = createAction('styleTable', name, { style });
            const data = JSON.parse(action.data);
            expect(data.style).toMatch(/^TableStyle(Light|Medium|Dark)\d+$/);
        }),
        { numRuns: 100 }
    );
});

// Property: Insert rows at any position
test('property: insert any number of rows at any position', () => {
    fc.assert(
        fc.property(
            fc.integer({ min: 1, max: 1000 }),
            fc.integer({ min: 1, max: 100 }),
            (position, count) => {
                const action = createAction('insertRows', `${position}:${position}`, { count });
                expect(JSON.parse(action.data).count).toBe(count);
            }
        ),
        { numRuns: 100 }
    );
});
```

### Configuring Property Tests

```javascript
fc.assert(
    fc.property(/* ... */),
    {
        numRuns: 100,           // Number of test cases
        seed: 12345,            // Reproducible seed
        verbose: true,          // Show all generated values
        endOnFailure: true,     // Stop on first failure
        skipAllAfterTimeLimit: 5000  // Timeout in ms
    }
);
```

---

## Test Coverage Goals

### Coverage Targets

| File | Current | Target | Priority |
|------|---------|--------|----------|
| action-executor.js | 0% | 85%+ | 🔴 Critical |
| ai-engine.js | ~80% | 95%+ | 🟡 High |
| preview.js | ~95% | 95%+ | 🟢 Good |
| history.js | ~95% | 95%+ | 🟢 Good |
| excel-data.js | Unknown | 80%+ | 🟡 High |

### Viewing Coverage Reports

```bash
# Generate coverage report
npm run test:coverage

# Open HTML report
open coverage/lcov-report/index.html
```

### Coverage Metrics

- **Line Coverage**: Percentage of lines executed
- **Branch Coverage**: Percentage of if/else branches taken
- **Function Coverage**: Percentage of functions called
- **Statement Coverage**: Percentage of statements executed

### Acceptable Coverage Gaps

Some code may be difficult to test:
- Error handling for rare Office.js errors
- Platform-specific code paths
- UI event handlers (tested manually)

Document gaps with comments:
```javascript
/* istanbul ignore next */
function handleRareError() {
    // Difficult to trigger in tests
}
```

---

## Integration Testing

### Workflow Test Pattern

```javascript
describe('Integration Tests - Sales Dashboard', () => {
    let ctx;
    
    beforeEach(() => {
        ctx = createStatefulMockContext();
    });
    
    test('complete dashboard workflow', async () => {
        const actions = [
            createAction('createTable', 'A1:E100', { name: 'Sales' }),
            createAction('styleTable', 'Sales', { style: 'TableStyleMedium2' }),
            createAction('createSlicer', 'Sales', { field: 'Region' }),
            createAction('createPivotTable', 'Sales', { name: 'SalesPivot' }),
            createAction('chart', 'SalesPivot', { type: 'Column' })
        ];
        
        const result = await executeWorkflow(ctx, actions);
        
        expect(result.successCount).toBe(actions.length);
        expect(result.totalTime).toBeLessThan(5000);
    });
});
```

### Testing Action Dependencies

```javascript
test('slicer requires table to exist first', async () => {
    const ctx = createStatefulMockContext();
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    
    // Create table first
    const table = sheet.tables.add('A1:D10', true);
    expect(sheet.tables.count).toBe(1);
    
    // Now create slicer
    const slicer = sheet.slicers.add(table, 'Column1', 'F1');
    expect(sheet.slicers.count).toBe(1);
});
```

### Error Recovery Testing

```javascript
test('workflow continues after non-critical error', async () => {
    const actions = [
        createAction('values', 'A1', [['Test']]),
        createAction('deleteTable', 'NonExistent', {}),  // Will fail
        createAction('values', 'B1', [['Continue']])     // Should still run
    ];
    
    const result = await executeWorkflow(ctx, actions);
    
    expect(result.results.length).toBe(3);
    expect(result.results[2].success).toBe(true);
});
```

---

## Performance Testing

### Measuring Execution Time

```javascript
async function measureExecutionTime(fn) {
    const start = performance.now();
    const result = await fn();
    const end = performance.now();
    return { result, executionTime: end - start };
}

test('10K rows insert < 2 seconds', async () => {
    const data = generateLargeDataset(10000, 5);
    
    const { executionTime } = await measureExecutionTime(async () => {
        range.values = data;
        await ctx.sync();
    });
    
    expect(executionTime).toBeLessThan(2000);
});
```

### Memory Testing

```javascript
function memorySnapshot() {
    if (process.memoryUsage) {
        return process.memoryUsage().heapUsed;
    }
    return 0;
}

test('no memory leak on repeated operations', async () => {
    const initial = memorySnapshot();
    
    for (let i = 0; i < 100; i++) {
        range.values = [[`Value${i}`]];
        await ctx.sync();
    }
    
    if (global.gc) global.gc();
    
    const final = memorySnapshot();
    const increase = (final - initial) / (1024 * 1024);
    
    expect(increase).toBeLessThan(50);  // <50MB
});
```

### Performance Benchmarks

| Operation | Dataset | Target | Warning |
|-----------|---------|--------|---------|
| Insert values | 10K rows | <2s | 50K rows |
| Create table | 10K rows | <2s | 50K rows |
| Apply formulas | 1K cells | <3s | 5K cells |
| Conditional format | 10K cells | <4s | 50K cells |
| Create chart | 1K points | <3s | 5K points |

---

## Debugging Tests

### Common Issues

**Test timeout**
```javascript
// Increase timeout for slow tests
test('slow operation', async () => {
    // ...
}, 10000);  // 10 second timeout
```

**Async issues**
```javascript
// Always await async operations
test('async test', async () => {
    await expect(asyncFn()).resolves.toBe(value);
});
```

**Mock not reset**
```javascript
beforeEach(() => {
    jest.clearAllMocks();
    ctx = createMockContext();
});
```

### Debugging Commands

```bash
# Run single test with debugging
node --inspect-brk node_modules/.bin/jest --runInBand -t "test name"

# Run with verbose output
npx jest --verbose

# Show console.log output
npx jest --silent=false
```

### Using Jest Debug Mode

```javascript
test('debug this test', () => {
    debugger;  // Breakpoint
    const result = someFunction();
    console.log('Result:', result);
    expect(result).toBeDefined();
});
```

---

## Contributing Tests

### Guidelines for New Tests

1. **Follow existing patterns** - Match the style of existing tests
2. **Test edge cases** - Empty inputs, large inputs, invalid inputs
3. **Use property-based tests** - For functions with many valid inputs
4. **Document complex tests** - Add comments explaining the test
5. **Keep tests focused** - One assertion per test when possible

### Checklist for New Features

- [ ] Unit tests for happy path
- [ ] Unit tests for edge cases
- [ ] Property-based tests (if applicable)
- [ ] Integration test for workflow
- [ ] Performance test (if large data)
- [ ] Update coverage targets

### Pull Request Requirements

1. All existing tests pass
2. New tests for new functionality
3. Coverage doesn't decrease
4. No console warnings/errors
5. Tests run in <5 minutes total

---

## Related Documentation

- [SETUP.md](../SETUP.md) - Installation and feature guide
- [Feature Matrix](feature-matrix.md) - Complete operation reference
- [Example Prompts](example-prompts.md) - Usage examples
