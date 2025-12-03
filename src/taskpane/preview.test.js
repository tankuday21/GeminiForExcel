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
        test('returns correct labels for all types', () => {
            expect(getActionSummary({ type: 'formula' })).toBe('Formula');
            expect(getActionSummary({ type: 'values' })).toBe('Values');
            expect(getActionSummary({ type: 'format' })).toBe('Format');
            expect(getActionSummary({ type: 'chart' })).toBe('Chart');
            expect(getActionSummary({ type: 'validation' })).toBe('Dropdown');
            expect(getActionSummary({ type: 'sort' })).toBe('Sort');
            expect(getActionSummary({ type: 'autofill' })).toBe('Autofill');
        });

        test('returns type for unknown types', () => {
            expect(getActionSummary({ type: 'unknown' })).toBe('unknown');
        });
    });
});
