# Design Document

## Overview

The Undo/History feature adds the ability to track applied AI actions and revert the most recent one. Before each action is applied, the original cell values are captured. Users can view a history panel showing all applied actions and click Undo to restore the previous state.

## Architecture

The feature integrates into the existing taskpane architecture:

```
┌─────────────────────────────────────┐
│           Excel Copilot             │
├─────────────────────────────────────┤
│  Header: [Clear] [History] [Settings]│
├─────────────────────────────────────┤
│  Chat Messages / Preview Panel      │
├─────────────────────────────────────┤
│  History Panel (toggleable)         │
│  ┌───────────────────────────────┐  │
│  │ ↩ Formula → A1 (2 min ago)   │  │
│  │   Format → A1:D1 (5 min ago) │  │
│  │   Chart → H2 (10 min ago)    │  │
│  └───────────────────────────────┘  │
├─────────────────────────────────────┤
│  [Undo Last] [Apply Changes]        │
└─────────────────────────────────────┘
```

## Components and Interfaces

### History Management

```javascript
/**
 * @typedef {Object} HistoryEntry
 * @property {string} id - Unique identifier for the entry
 * @property {string} type - Action type (formula, values, format, chart, validation)
 * @property {string} target - Target cell or range address
 * @property {number} timestamp - Unix timestamp when action was applied
 * @property {Object} undoData - Data needed to restore previous state
 * @property {string[][]} undoData.values - Original cell values
 * @property {string[][]} undoData.formulas - Original cell formulas
 * @property {string} undoData.address - Full range address
 */

/**
 * Captures the current state of a range for undo
 * @param {string} rangeAddress - The range to capture
 * @returns {Promise<Object>} The captured undo data
 */
async function captureUndoData(rangeAddress)

/**
 * Adds an entry to the action history
 * @param {Object} action - The action that was applied
 * @param {Object} undoData - The captured undo data
 * @returns {HistoryEntry} The created history entry
 */
function addToHistory(action, undoData)

/**
 * Removes the most recent entry from history
 * @returns {HistoryEntry|null} The removed entry, or null if history empty
 */
function removeFromHistory()

/**
 * Gets all history entries
 * @returns {HistoryEntry[]} All history entries, newest first
 */
function getHistory()

/**
 * Clears all history entries
 * @returns {void}
 */
function clearHistory()
```

### Undo Operations

```javascript
/**
 * Performs undo of the most recent action
 * @returns {Promise<boolean>} True if undo succeeded
 */
async function performUndo()

/**
 * Restores cells to their previous state using undo data
 * @param {Object} undoData - The undo data containing original values
 * @returns {Promise<void>}
 */
async function restoreFromUndoData(undoData)
```

### History Panel UI

```javascript
/**
 * Renders the history panel HTML
 * @param {HistoryEntry[]} entries - History entries to display
 * @returns {string} HTML string for the history panel
 */
function renderHistoryPanel(entries)

/**
 * Renders a single history entry
 * @param {HistoryEntry} entry - The entry to render
 * @returns {string} HTML string for the entry
 */
function renderHistoryEntry(entry)

/**
 * Formats a timestamp as relative time (e.g., "2 min ago")
 * @param {number} timestamp - Unix timestamp
 * @returns {string} Formatted relative time string
 */
function formatRelativeTime(timestamp)

/**
 * Shows or hides the history panel
 * @param {boolean} visible - Whether to show the panel
 */
function toggleHistoryPanel(visible)
```

## Data Models

### HistoryState (added to application state)

```javascript
/**
 * @typedef {Object} HistoryState
 * @property {HistoryEntry[]} entries - All history entries, newest first
 * @property {boolean} panelVisible - Whether history panel is shown
 * @property {number} maxEntries - Maximum entries to retain (default: 20)
 */
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Undo restores original state (Round-trip)
*For any* cell range and action, capturing undo data before applying the action and then restoring from that undo data SHALL return the cells to their original values.
**Validates: Requirements 1.1, 1.2**

### Property 2: History entry contains required fields
*For any* history entry, the rendered output SHALL contain the action type, target range, and a formatted timestamp.
**Validates: Requirements 2.2**

### Property 3: New entries are prepended to history
*For any* sequence of actions added to history, the most recently added action SHALL always be at index 0 of the history array.
**Validates: Requirements 2.4**

### Property 4: Undo removes entry from history
*For any* non-empty history, performing undo SHALL decrease the history length by exactly 1 and remove the first (most recent) entry.
**Validates: Requirements 1.3**

### Property 5: History respects maximum limit
*For any* history with entries at the maximum limit, adding a new entry SHALL maintain the limit by removing the oldest entry.
**Validates: Requirements 3.3**

### Property 6: History panel renders all entries
*For any* non-empty history, the rendered history panel SHALL contain exactly one entry element for each history entry.
**Validates: Requirements 2.1**

## Error Handling

| Error Scenario | Handling Strategy |
|----------------|-------------------|
| Excel API error during undo | Show error toast, retain entry in history, log error |
| Invalid range in undo data | Show error message, remove corrupted entry from history |
| Empty history on undo attempt | Disable undo button, show "Nothing to undo" message |
| Capture fails before action | Abort action, show error, don't add to history |

## Testing Strategy

### Unit Tests
- Test `formatRelativeTime` with various timestamps
- Test `addToHistory` adds entries correctly
- Test `removeFromHistory` removes first entry
- Test `getHistory` returns entries in correct order
- Test history max limit enforcement

### Property-Based Tests
Property-based tests will use fast-check library to verify correctness properties hold across many random inputs.

- **Property 1**: Generate random cell values, capture, modify, restore - verify match
- **Property 2**: Generate random history entries, verify rendered output contains required fields
- **Property 3**: Generate random action sequences, verify newest is always first
- **Property 4**: Generate random history, perform undo, verify length decreased by 1
- **Property 5**: Generate history at max limit, add entry, verify limit maintained
- **Property 6**: Generate random history, verify rendered entry count matches

Each property-based test will run a minimum of 100 iterations. Tests will be tagged with format: `**Feature: undo-history, Property {number}: {property_text}**`
