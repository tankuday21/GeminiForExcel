/**
 * Property-Based Tests for History Module
 * Using fast-check for property-based testing
 */

const fc = require('fast-check');
const {
    MAX_ENTRIES,
    createHistoryEntry,
    addToHistory,
    removeFromHistory,
    getLatestEntry,
    hasHistory,
    formatRelativeTime,
    renderHistoryEntry,
    renderHistoryList
} = require('./history');

// Arbitrary generators
const actionTypeArb = fc.constantFrom('formula', 'values', 'format', 'chart', 'validation', 'sort', 'autofill');

const cellRefArb = fc.tuple(
    fc.constantFrom('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'),
    fc.integer({ min: 1, max: 100 })
).map(([col, row]) => `${col}${row}`);

const actionArb = fc.record({
    type: actionTypeArb,
    target: cellRefArb,
    data: fc.option(fc.string({ minLength: 0, maxLength: 50 }), { nil: undefined })
});

const undoDataArb = fc.record({
    values: fc.array(fc.array(fc.oneof(fc.string(), fc.integer(), fc.constant(null)), { minLength: 1, maxLength: 5 }), { minLength: 1, maxLength: 10 }),
    formulas: fc.array(fc.array(fc.string(), { minLength: 1, maxLength: 5 }), { minLength: 1, maxLength: 10 }),
    address: cellRefArb
});

const historyEntryArb = fc.record({
    id: fc.string({ minLength: 5, maxLength: 15 }),
    type: actionTypeArb,
    target: cellRefArb,
    timestamp: fc.integer({ min: Date.now() - 86400000, max: Date.now() }),
    undoData: undoDataArb
});

const historyEntriesArb = fc.array(historyEntryArb, { minLength: 0, maxLength: 25 });

describe('History Module - Property Based Tests', () => {

    /**
     * **Feature: undo-history, Property 3: New entries are prepended to history**
     * **Validates: Requirements 2.4**
     */
    describe('Property 3: New entries are prepended to history', () => {
        test('new entry is always at index 0', () => {
            fc.assert(
                fc.property(historyEntriesArb, actionArb, undoDataArb, (entries, action, undoData) => {
                    const newEntry = createHistoryEntry(action, undoData);
                    const updated = addToHistory(entries, newEntry);
                    
                    expect(updated[0].id).toBe(newEntry.id);
                    expect(updated[0].type).toBe(action.type);
                    expect(updated[0].target).toBe(action.target);
                }),
                { numRuns: 100 }
            );
        });

        test('existing entries shift down by one index', () => {
            fc.assert(
                fc.property(
                    fc.array(historyEntryArb, { minLength: 1, maxLength: 10 }),
                    actionArb,
                    undoDataArb,
                    (entries, action, undoData) => {
                        const newEntry = createHistoryEntry(action, undoData);
                        const updated = addToHistory(entries, newEntry, 100); // High limit to avoid truncation
                        
                        // Each original entry should be at index + 1
                        entries.forEach((entry, i) => {
                            expect(updated[i + 1].id).toBe(entry.id);
                        });
                    }
                ),
                { numRuns: 100 }
            );
        });
    });

    /**
     * **Feature: undo-history, Property 5: History respects maximum limit**
     * **Validates: Requirements 3.3**
     */
    describe('Property 5: History respects maximum limit', () => {
        test('adding to full history removes oldest entry', () => {
            fc.assert(
                fc.property(
                    fc.integer({ min: 1, max: 10 }),
                    actionArb,
                    undoDataArb,
                    (maxLimit, action, undoData) => {
                        // Create entries at max limit
                        let entries = [];
                        for (let i = 0; i < maxLimit; i++) {
                            entries.push(createHistoryEntry({ type: 'formula', target: `A${i}` }, {}));
                        }
                        
                        const newEntry = createHistoryEntry(action, undoData);
                        const updated = addToHistory(entries, newEntry, maxLimit);
                        
                        expect(updated.length).toBe(maxLimit);
                        expect(updated[0].id).toBe(newEntry.id);
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('history never exceeds max limit', () => {
            fc.assert(
                fc.property(
                    fc.array(historyEntryArb, { minLength: 0, maxLength: 30 }),
                    actionArb,
                    undoDataArb,
                    fc.integer({ min: 1, max: 25 }),
                    (entries, action, undoData, maxLimit) => {
                        const newEntry = createHistoryEntry(action, undoData);
                        const updated = addToHistory(entries, newEntry, maxLimit);
                        
                        expect(updated.length).toBeLessThanOrEqual(maxLimit);
                    }
                ),
                { numRuns: 100 }
            );
        });
    });

    /**
     * **Feature: undo-history, Property 4: Undo removes entry from history**
     * **Validates: Requirements 1.3**
     */
    describe('Property 4: Undo removes entry from history', () => {
        test('removeFromHistory decreases length by 1', () => {
            fc.assert(
                fc.property(
                    fc.array(historyEntryArb, { minLength: 1, maxLength: 20 }),
                    (entries) => {
                        const originalLength = entries.length;
                        const { entries: updated, removed } = removeFromHistory(entries);
                        
                        expect(updated.length).toBe(originalLength - 1);
                        expect(removed).not.toBeNull();
                        expect(removed.id).toBe(entries[0].id);
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('removeFromHistory removes the first (most recent) entry', () => {
            fc.assert(
                fc.property(
                    fc.array(historyEntryArb, { minLength: 2, maxLength: 20 }),
                    (entries) => {
                        const { entries: updated, removed } = removeFromHistory(entries);
                        
                        // The removed entry should be the first one
                        expect(removed.id).toBe(entries[0].id);
                        
                        // The new first entry should be the old second entry
                        expect(updated[0].id).toBe(entries[1].id);
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('removeFromHistory on empty returns null', () => {
            const { entries, removed } = removeFromHistory([]);
            expect(entries).toEqual([]);
            expect(removed).toBeNull();
        });
    });

    /**
     * **Feature: undo-history, Property 2: History entry contains required fields**
     * **Validates: Requirements 2.2**
     */
    describe('Property 2: History entry contains required fields', () => {
        test('rendered entry contains type, target, and timestamp', () => {
            fc.assert(
                fc.property(historyEntryArb, (entry) => {
                    const html = renderHistoryEntry(entry, () => '<svg></svg>');
                    
                    // Should contain the target
                    expect(html).toContain(entry.target);
                    
                    // Should contain a time string
                    expect(html).toMatch(/(\d+ (min|hr|days?) ago|just now|yesterday)/);
                    
                    // Should contain the entry id
                    expect(html).toContain(entry.id);
                }),
                { numRuns: 100 }
            );
        });
    });

    /**
     * **Feature: undo-history, Property 6: History panel renders all entries**
     * **Validates: Requirements 2.1**
     */
    describe('Property 6: History panel renders all entries', () => {
        test('renderHistoryList contains one entry per history item', () => {
            fc.assert(
                fc.property(
                    fc.array(historyEntryArb, { minLength: 1, maxLength: 20 }),
                    (entries) => {
                        const html = renderHistoryList(entries, () => '<svg></svg>');
                        const entryCount = (html.match(/class="history-entry"/g) || []).length;
                        
                        expect(entryCount).toBe(entries.length);
                    }
                ),
                { numRuns: 100 }
            );
        });

        test('empty history shows empty message', () => {
            const html = renderHistoryList([], () => '<svg></svg>');
            expect(html).toContain('No actions yet');
        });
    });

    // Unit tests for formatRelativeTime
    describe('formatRelativeTime', () => {
        test('returns "just now" for recent timestamps', () => {
            const now = Date.now();
            expect(formatRelativeTime(now)).toBe('just now');
            expect(formatRelativeTime(now - 30000)).toBe('just now'); // 30 seconds ago
        });

        test('returns minutes for timestamps within an hour', () => {
            const now = Date.now();
            expect(formatRelativeTime(now - 60000)).toBe('1 min ago');
            expect(formatRelativeTime(now - 300000)).toBe('5 min ago');
            expect(formatRelativeTime(now - 3540000)).toBe('59 min ago');
        });

        test('returns hours for timestamps within a day', () => {
            const now = Date.now();
            expect(formatRelativeTime(now - 3600000)).toBe('1 hr ago');
            expect(formatRelativeTime(now - 7200000)).toBe('2 hr ago');
        });

        test('returns days for older timestamps', () => {
            const now = Date.now();
            expect(formatRelativeTime(now - 86400000)).toBe('yesterday');
            expect(formatRelativeTime(now - 172800000)).toBe('2 days ago');
        });
    });

    // Unit tests for hasHistory
    describe('hasHistory', () => {
        test('returns false for empty array', () => {
            expect(hasHistory([])).toBe(false);
        });

        test('returns false for null/undefined', () => {
            expect(hasHistory(null)).toBe(false);
            expect(hasHistory(undefined)).toBe(false);
        });

        test('returns true for non-empty array', () => {
            expect(hasHistory([{ id: '1' }])).toBe(true);
        });
    });

    // Unit tests for getLatestEntry
    describe('getLatestEntry', () => {
        test('returns null for empty array', () => {
            expect(getLatestEntry([])).toBeNull();
        });

        test('returns first entry for non-empty array', () => {
            const entries = [{ id: 'first' }, { id: 'second' }];
            expect(getLatestEntry(entries).id).toBe('first');
        });
    });
});
