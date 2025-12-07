/**
 * Property-Based Tests for Preview Panel
 * Using fast-check for property-based testing
 */

const fc = require('fast-check');
const {
    ACTION_TYPES,
    getActionIcon,
    getActionSummary,
    getActionDetails,
    filterSelectedActions,
    hasSelectedActions,
    renderPreviewItem,
    renderPreviewList
} = require('./preview');

// Arbitrary generators for actions
const actionTypeArb = fc.constantFrom(...ACTION_TYPES);

// Generate Excel-like cell references
const cellRefArb = fc.tuple(
    fc.constantFrom('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'),
    fc.integer({ min: 1, max: 100 })
).map(([col, row]) => `${col}${row}`);

const rangeArb = fc.tuple(cellRefArb, cellRefArb).map(([start, end]) => `${start}:${end}`);

const actionArb = fc.record({
    type: actionTypeArb,
    target: fc.oneof(cellRefArb, rangeArb),
    source: fc.option(rangeArb, { nil: undefined }),
    chartType: fc.option(fc.constantFrom('column', 'bar', 'line', 'pie'), { nil: undefined }),
    title: fc.option(fc.string({ minLength: 1, maxLength: 30 }), { nil: undefined }),
    position: fc.option(cellRefArb, { nil: undefined }),
    data: fc.option(fc.string({ minLength: 0, maxLength: 100 }), { nil: undefined })
});

const actionsArrayArb = fc.array(actionArb, { minLength: 1, maxLength: 10 });

describe('Preview Panel - Property Based Tests', () => {
    
    /**
     * **Feature: preview-before-apply, Property 5: Action type maps to distinct icon**
     * **Validates: Requirements 3.1**
     */
    describe('Property 5: Action type maps to distinct icon', () => {
        test('all action types return different icons', () => {
            const icons = ACTION_TYPES.map(type => getActionIcon(type));
            const uniqueIcons = new Set(icons);
            expect(uniqueIcons.size).toBe(ACTION_TYPES.length);
        });

        test('each action type returns a valid SVG', () => {
            fc.assert(
                fc.property(actionTypeArb, (type) => {
                    const icon = getActionIcon(type);
                    expect(icon).toContain('<svg');
                    expect(icon).toContain('</svg>');
                }),
                { numRuns: 100 }
            );
        });

        test('unknown types fall back to formula icon', () => {
            const unknownIcon = getActionIcon('unknown');
            const formulaIcon = getActionIcon('formula');
            expect(unknownIcon).toBe(formulaIcon);
        });
    });

    /**
     * **Feature: preview-before-apply, Property 1: Preview renders all actions**
     * **Validates: Requirements 1.1**
     */
    describe('Property 1: Preview renders all actions', () => {
        test('preview list contains one item per action', () => {
            fc.assert(
                fc.property(actionsArrayArb, (actions) => {
                    const selections = actions.map(() => true);
                    const html = renderPreviewList(actions, selections, -1);
                    const itemCount = (html.match(/class="preview-item/g) || []).length;
                    expect(itemCount).toBe(actions.length);
                }),
                { numRuns: 100 }
            );
        });

        test('empty actions array returns empty string', () => {
            const html = renderPreviewList([], [], -1);
            expect(html).toBe('');
        });
    });

    /**
     * **Feature: preview-before-apply, Property 3: Each action has a checkbox**
     * **Validates: Requirements 2.1**
     */
    describe('Property 3: Each action has a checkbox', () => {
        test('each action has exactly one checkbox', () => {
            fc.assert(
                fc.property(actionsArrayArb, (actions) => {
                    const selections = actions.map(() => true);
                    const html = renderPreviewList(actions, selections, -1);
                    const checkboxCount = (html.match(/class="preview-checkbox"/g) || []).length;
                    expect(checkboxCount).toBe(actions.length);
                }),
                { numRuns: 100 }
            );
        });
    });

    /**
     * **Feature: preview-before-apply, Property 4: Filter returns only selected actions**
     * **Validates: Requirements 2.2, 2.3**
     */
    describe('Property 4: Filter returns only selected actions', () => {
        test('filter returns only actions where selection is true', () => {
            fc.assert(
                fc.property(
                    actionsArrayArb,
                    fc.array(fc.boolean(), { minLength: 1, maxLength: 10 }),
                    (actions, rawSelections) => {
                        // Ensure selections array matches actions length
                        const selections = actions.map((_, i) => rawSelections[i % rawSelections.length]);
                        const filtered = filterSelectedActions(actions, selections);
                        
                        // Count expected selected
                        const expectedCount = selections.filter(s => s === true).length;
                        expect(filtered.length).toBe(expectedCount);
                        
                        // Verify all returned actions were selected
                        filtered.forEach(action => {
                            const originalIndex = actions.indexOf(action);
                            expect(selections[originalIndex]).toBe(true);
                        });
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('filter preserves order of selected actions', () => {
            fc.assert(
                fc.property(actionsArrayArb, (actions) => {
                    // Select every other action
                    const selections = actions.map((_, i) => i % 2 === 0);
                    const filtered = filterSelectedActions(actions, selections);
                    
                    // Verify order is preserved
                    let lastIndex = -1;
                    filtered.forEach(action => {
                        const currentIndex = actions.indexOf(action);
                        expect(currentIndex).toBeGreaterThan(lastIndex);
                        lastIndex = currentIndex;
                    });
                }),
                { numRuns: 100 }
            );
        });

        test('empty inputs return empty array', () => {
            expect(filterSelectedActions([], [])).toEqual([]);
            expect(filterSelectedActions(null, null)).toEqual([]);
            expect(filterSelectedActions(undefined, undefined)).toEqual([]);
        });
    });

    /**
     * **Feature: preview-before-apply, Property 2: Action rendering includes required fields**
     * **Validates: Requirements 1.2, 1.3, 1.4, 1.5, 1.6**
     */
    describe('Property 2: Action rendering includes required fields', () => {
        test('formula actions show target and formula', () => {
            fc.assert(
                fc.property(
                    fc.record({
                        type: fc.constant('formula'),
                        target: fc.string({ minLength: 2, maxLength: 10 }),
                        data: fc.string({ minLength: 1, maxLength: 50 })
                    }),
                    (action) => {
                        const html = renderPreviewItem(action, 0, true, true, false);
                        expect(html).toContain(action.target);
                        // Details should contain the formula data
                        const details = getActionDetails(action);
                        expect(details).toBe(action.data);
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('chart actions show chart type, target, and position', () => {
            fc.assert(
                fc.property(
                    fc.record({
                        type: fc.constant('chart'),
                        target: fc.string({ minLength: 2, maxLength: 10 }),
                        chartType: fc.constantFrom('column', 'bar', 'line', 'pie'),
                        position: fc.string({ minLength: 2, maxLength: 5 }),
                        title: fc.option(fc.string({ minLength: 1, maxLength: 20 }), { nil: undefined })
                    }),
                    (action) => {
                        const details = getActionDetails(action);
                        expect(details).toContain(action.chartType);
                        expect(details).toContain(action.target);
                        expect(details).toContain(action.position);
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('validation actions show target and source', () => {
            fc.assert(
                fc.property(
                    fc.record({
                        type: fc.constant('validation'),
                        target: fc.string({ minLength: 2, maxLength: 10 }),
                        source: fc.string({ minLength: 2, maxLength: 20 })
                    }),
                    (action) => {
                        const html = renderPreviewItem(action, 0, true, true, false);
                        expect(html).toContain(action.target);
                        const details = getActionDetails(action);
                        expect(details).toContain(action.source);
                    }
                ),
                { numRuns: 100 }
            );
        });
    });

    /**
     * **Feature: preview-before-apply, Property 6: Collapsed view shows summary**
     * **Validates: Requirements 3.4**
     */
    describe('Property 6: Collapsed view shows summary', () => {
        test('collapsed view contains action type and target', () => {
            fc.assert(
                fc.property(actionArb, (action) => {
                    const html = renderPreviewItem(action, 0, false, true, false);
                    const summary = getActionSummary(action);
                    expect(html).toContain(summary);
                    expect(html).toContain(action.target);
                    // Should not have expanded class
                    expect(html).not.toContain('class="preview-item expanded');
                }),
                { numRuns: 100 }
            );
        });
    });

    /**
     * **Feature: preview-before-apply, Property 7: Expanded view shows full details**
     * **Validates: Requirements 3.3**
     */
    describe('Property 7: Expanded view shows full details', () => {
        test('expanded view has expanded class and shows details', () => {
            fc.assert(
                fc.property(actionArb, (action) => {
                    const html = renderPreviewItem(action, 0, true, true, false);
                    // Should have expanded class
                    expect(html).toContain('expanded');
                    // Should contain details section
                    expect(html).toContain('preview-details');
                }),
                { numRuns: 100 }
            );
        });
    });

    // Unit tests for hasSelectedActions
    describe('hasSelectedActions', () => {
        test('returns false for empty array', () => {
            expect(hasSelectedActions([])).toBe(false);
        });

        test('returns false for all false', () => {
            expect(hasSelectedActions([false, false, false])).toBe(false);
        });

        test('returns true for at least one true', () => {
            expect(hasSelectedActions([false, true, false])).toBe(true);
        });

        test('returns true for all true', () => {
            expect(hasSelectedActions([true, true, true])).toBe(true);
        });

        test('returns false for null/undefined', () => {
            expect(hasSelectedActions(null)).toBe(false);
            expect(hasSelectedActions(undefined)).toBe(false);
        });
    });

    // Unit tests for getActionSummary
    describe('getActionSummary', () => {
        test('returns correct labels for basic types', () => {
            expect(getActionSummary({ type: 'formula' })).toBe('Formula');
            expect(getActionSummary({ type: 'values' })).toBe('Values');
            expect(getActionSummary({ type: 'format' })).toBe('Format');
            expect(getActionSummary({ type: 'chart' })).toBe('Chart');
            expect(getActionSummary({ type: 'validation' })).toBe('Dropdown');
            expect(getActionSummary({ type: 'sort' })).toBe('Sort');
            expect(getActionSummary({ type: 'autofill' })).toBe('Autofill');
        });

        test('returns correct labels for table operations', () => {
            expect(getActionSummary({ type: 'createTable' })).toBe('Create Table');
            expect(getActionSummary({ type: 'styleTable' })).toBe('Style Table');
            expect(getActionSummary({ type: 'addTableRow' })).toBe('Add Table Row');
            expect(getActionSummary({ type: 'addTableColumn' })).toBe('Add Table Column');
            expect(getActionSummary({ type: 'resizeTable' })).toBe('Resize Table');
            expect(getActionSummary({ type: 'convertToRange' })).toBe('Convert to Range');
            expect(getActionSummary({ type: 'toggleTableTotals' })).toBe('Toggle Totals');
        });

        test('returns correct labels for pivot operations', () => {
            expect(getActionSummary({ type: 'createPivotTable' })).toBe('Create PivotTable');
            expect(getActionSummary({ type: 'addPivotField' })).toBe('Add Pivot Field');
            expect(getActionSummary({ type: 'configurePivotLayout' })).toBe('Configure Pivot');
            expect(getActionSummary({ type: 'refreshPivotTable' })).toBe('Refresh PivotTable');
            expect(getActionSummary({ type: 'deletePivotTable' })).toBe('Delete PivotTable');
        });

        test('returns correct labels for slicer operations', () => {
            expect(getActionSummary({ type: 'createSlicer' })).toBe('Create Slicer');
            expect(getActionSummary({ type: 'configureSlicer' })).toBe('Configure Slicer');
            expect(getActionSummary({ type: 'connectSlicerToTable' })).toBe('Connect Slicer to Table');
            expect(getActionSummary({ type: 'connectSlicerToPivot' })).toBe('Connect Slicer to Pivot');
            expect(getActionSummary({ type: 'deleteSlicer' })).toBe('Delete Slicer');
        });

        test('returns correct labels for comment operations', () => {
            expect(getActionSummary({ type: 'addComment' })).toBe('Add Comment');
            expect(getActionSummary({ type: 'addNote' })).toBe('Add Note');
            expect(getActionSummary({ type: 'editComment' })).toBe('Edit Comment');
            expect(getActionSummary({ type: 'deleteComment' })).toBe('Delete Comment');
            expect(getActionSummary({ type: 'replyToComment' })).toBe('Reply to Comment');
            expect(getActionSummary({ type: 'resolveComment' })).toBe('Resolve Comment');
        });

        test('returns correct labels for protection operations', () => {
            expect(getActionSummary({ type: 'protectWorksheet' })).toBe('Protect Sheet');
            expect(getActionSummary({ type: 'unprotectWorksheet' })).toBe('Unprotect Sheet');
            expect(getActionSummary({ type: 'protectRange' })).toBe('Protect Range');
            expect(getActionSummary({ type: 'protectWorkbook' })).toBe('Protect Workbook');
        });

        test('returns correct labels for page setup operations', () => {
            expect(getActionSummary({ type: 'setPageSetup' })).toBe('Page Setup');
            expect(getActionSummary({ type: 'setPageMargins' })).toBe('Set Margins');
            expect(getActionSummary({ type: 'setPageOrientation' })).toBe('Set Orientation');
            expect(getActionSummary({ type: 'setPrintArea' })).toBe('Set Print Area');
            expect(getActionSummary({ type: 'setHeaderFooter' })).toBe('Set Header/Footer');
            expect(getActionSummary({ type: 'setPageBreaks' })).toBe('Set Page Breaks');
        });

        test('returns type for unknown types', () => {
            expect(getActionSummary({ type: 'unknown' })).toBe('unknown');
        });

        test('returns correct labels for advanced formatting types', () => {
            expect(getActionSummary({ type: 'conditionalFormat' })).toBe('Conditional Format');
            expect(getActionSummary({ type: 'clearFormat' })).toBe('Clear Format');
        });

        test('returns correct labels for copy/filter/duplicates types', () => {
            expect(getActionSummary({ type: 'copy' })).toBe('Copy');
            expect(getActionSummary({ type: 'copyValues' })).toBe('Copy Values');
            expect(getActionSummary({ type: 'filter' })).toBe('Filter');
            expect(getActionSummary({ type: 'clearFilter' })).toBe('Clear Filter');
            expect(getActionSummary({ type: 'removeDuplicates' })).toBe('Remove Duplicates');
        });

        test('returns correct labels for named range types', () => {
            expect(getActionSummary({ type: 'createNamedRange' })).toBe('Create Named Range');
            expect(getActionSummary({ type: 'deleteNamedRange' })).toBe('Delete Named Range');
            expect(getActionSummary({ type: 'updateNamedRange' })).toBe('Update Named Range');
            expect(getActionSummary({ type: 'listNamedRanges' })).toBe('List Named Ranges');
        });

        test('returns correct labels for shape types', () => {
            expect(getActionSummary({ type: 'insertShape' })).toBe('Insert Shape');
            expect(getActionSummary({ type: 'insertImage' })).toBe('Insert Image');
            expect(getActionSummary({ type: 'insertTextBox' })).toBe('Insert Text Box');
            expect(getActionSummary({ type: 'formatShape' })).toBe('Format Shape');
            expect(getActionSummary({ type: 'deleteShape' })).toBe('Delete Shape');
            expect(getActionSummary({ type: 'groupShapes' })).toBe('Group Shapes');
            expect(getActionSummary({ type: 'arrangeShapes' })).toBe('Arrange Shapes');
            expect(getActionSummary({ type: 'ungroupShapes' })).toBe('Ungroup Shapes');
        });

        test('returns correct labels for sparkline types', () => {
            expect(getActionSummary({ type: 'createSparkline' })).toBe('Create Sparkline');
            expect(getActionSummary({ type: 'configureSparkline' })).toBe('Configure Sparkline');
            expect(getActionSummary({ type: 'deleteSparkline' })).toBe('Delete Sparkline');
        });

        test('returns correct labels for worksheet management types', () => {
            expect(getActionSummary({ type: 'renameSheet' })).toBe('Rename Sheet');
            expect(getActionSummary({ type: 'moveSheet' })).toBe('Move Sheet');
            expect(getActionSummary({ type: 'hideSheet' })).toBe('Hide Sheet');
            expect(getActionSummary({ type: 'unhideSheet' })).toBe('Unhide Sheet');
            expect(getActionSummary({ type: 'freezePanes' })).toBe('Freeze Panes');
            expect(getActionSummary({ type: 'unfreezePane' })).toBe('Unfreeze Panes');
            expect(getActionSummary({ type: 'setZoom' })).toBe('Set Zoom');
            expect(getActionSummary({ type: 'splitPane' })).toBe('Split Panes');
            expect(getActionSummary({ type: 'createView' })).toBe('Create View');
        });

        test('returns correct labels for data manipulation types', () => {
            expect(getActionSummary({ type: 'insertRows' })).toBe('Insert Rows');
            expect(getActionSummary({ type: 'insertColumns' })).toBe('Insert Columns');
            expect(getActionSummary({ type: 'deleteRows' })).toBe('Delete Rows');
            expect(getActionSummary({ type: 'deleteColumns' })).toBe('Delete Columns');
            expect(getActionSummary({ type: 'mergeCells' })).toBe('Merge Cells');
            expect(getActionSummary({ type: 'unmergeCells' })).toBe('Unmerge Cells');
            expect(getActionSummary({ type: 'findReplace' })).toBe('Find & Replace');
            expect(getActionSummary({ type: 'textToColumns' })).toBe('Text to Columns');
        });

        test('returns correct labels for data type operations', () => {
            expect(getActionSummary({ type: 'insertDataType' })).toBe('Insert Entity');
            expect(getActionSummary({ type: 'refreshDataType' })).toBe('Refresh Entity');
        });
    });

    /**
     * **Property 8: All action types have distinct icons**
     * **Validates: Complete icon coverage for 87 action types**
     */
    describe('Property 8: All action types have distinct icons', () => {
        test('all 87 action types have unique icons', () => {
            const icons = ACTION_TYPES.map(type => getActionIcon(type));
            const uniqueIcons = new Set(icons);
            // All icons should be unique
            expect(uniqueIcons.size).toBe(ACTION_TYPES.length);
        });

        test('ACTION_TYPES contains expected count', () => {
            expect(ACTION_TYPES.length).toBe(90);
        });
    });

    /**
     * **Property 9: All action types have summary labels**
     * **Validates: Complete label coverage**
     */
    describe('Property 9: All action types have summary labels', () => {
        test('all action types return non-empty labels', () => {
            fc.assert(
                fc.property(actionTypeArb, (type) => {
                    const label = getActionSummary({ type });
                    expect(label).toBeTruthy();
                    expect(label.length).toBeGreaterThan(0);
                }),
                { numRuns: 100 }
            );
        });
    });

    /**
     * **Property 10: Complex actions parse JSON correctly**
     * **Validates: JSON parsing in getActionDetails**
     */
    describe('Property 10: Complex actions parse JSON correctly', () => {
        test('table actions parse JSON data', () => {
            const action = {
                type: 'createTable',
                target: 'A1:D10',
                data: JSON.stringify({ tableName: 'SalesData', style: 'TableStyleMedium2', hasHeaders: true })
            };
            const details = getActionDetails(action);
            expect(details).toContain('SalesData');
            expect(details).toContain('TableStyleMedium2');
        });

        test('pivot actions parse JSON data', () => {
            const action = {
                type: 'addPivotField',
                target: 'PivotTable1',
                data: JSON.stringify({ field: 'Revenue', area: 'data', aggregation: 'Sum' })
            };
            const details = getActionDetails(action);
            expect(details).toContain('Revenue');
            expect(details).toContain('data');
        });

        test('comment actions parse JSON data', () => {
            const action = {
                type: 'addComment',
                target: 'A1',
                data: JSON.stringify({ author: 'John', content: 'This is a test comment' })
            };
            const details = getActionDetails(action);
            expect(details).toContain('John');
            expect(details).toContain('test comment');
        });

        test('page setup actions parse JSON data', () => {
            const action = {
                type: 'setPageMargins',
                target: 'Sheet1',
                data: JSON.stringify({ top: 1, bottom: 1, left: 0.75, right: 0.75 })
            };
            const details = getActionDetails(action);
            expect(details).toContain('Top');
            expect(details).toContain('Bottom');
        });
    });

    /**
     * **Property 11: Action details handle missing data gracefully**
     * **Validates: Fallback behavior**
     */
    describe('Property 11: Action details handle missing data gracefully', () => {
        test('actions with null data return fallback', () => {
            fc.assert(
                fc.property(actionTypeArb, (type) => {
                    const action = { type, target: 'A1', data: null };
                    const details = getActionDetails(action);
                    expect(details).toBeTruthy();
                }),
                { numRuns: 100 }
            );
        });

        test('actions with invalid JSON return fallback', () => {
            const action = {
                type: 'createTable',
                target: 'A1:D10',
                data: 'not valid json'
            };
            const details = getActionDetails(action);
            // Falls back to descriptive message when JSON parsing fails
            expect(details).toContain('A1:D10');
        });
    });

    /**
     * **Property 12: Preview rendering works for all action types**
     * **Validates: renderPreviewItem produces valid HTML**
     */
    describe('Property 12: Preview rendering works for all action types', () => {
        test('all action types render valid HTML', () => {
            fc.assert(
                fc.property(actionTypeArb, (type) => {
                    const action = { type, target: 'A1', data: '{}' };
                    const html = renderPreviewItem(action, 0, false, true, false);
                    expect(html).toContain('preview-item');
                    expect(html).toContain('preview-checkbox');
                    expect(html).toContain('preview-icon');
                    expect(html).toContain('preview-content');
                }),
                { numRuns: 100 }
            );
        });
    });
});
