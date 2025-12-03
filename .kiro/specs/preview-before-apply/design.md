# Design Document

## Overview

The Preview Before Apply feature adds a visual preview panel to the Excel AI Copilot that displays all proposed changes before they are executed. Users can review each action, selectively include/exclude actions, and see target cells highlighted in Excel when hovering over preview items.

## Architecture

The feature integrates into the existing taskpane architecture:

```
┌─────────────────────────────────────┐
│           Excel Copilot             │
├─────────────────────────────────────┤
│  Chat Messages                      │
│  ┌───────────────────────────────┐  │
│  │ AI Response with actions      │  │
│  └───────────────────────────────┘  │
├─────────────────────────────────────┤
│  Preview Panel (NEW)                │
│  ┌───────────────────────────────┐  │
│  │ ☑ Formula → A1: =SUM(B1:B10) │  │
│  │ ☑ Format → A1:D1: Bold, Blue │  │
│  │ ☐ Chart → Column at H2       │  │
│  └───────────────────────────────┘  │
├─────────────────────────────────────┤
│  [Apply Selected] [Clear All]       │
└─────────────────────────────────────┘
```

## Components and Interfaces

### PreviewPanel Component

```javascript
/**
 * Renders the preview panel for pending actions
 * @param {Action[]} actions - Array of pending actions
 * @param {Function} onSelectionChange - Callback when action selection changes
 * @returns {HTMLElement} The preview panel element
 */
function renderPreviewPanel(actions, onSelectionChange)

/**
 * Renders a single action preview item
 * @param {Action} action - The action to render
 * @param {number} index - Index in the actions array
 * @param {boolean} isExpanded - Whether to show full details
 * @returns {HTMLElement} The preview item element
 */
function renderPreviewItem(action, index, isExpanded)

/**
 * Gets the icon for an action type
 * @param {string} type - Action type (formula, values, format, chart, validation)
 * @returns {string} SVG icon markup
 */
function getActionIcon(type)

/**
 * Gets a summary string for an action
 * @param {Action} action - The action to summarize
 * @returns {string} Human-readable summary
 */
function getActionSummary(action)

/**
 * Gets detailed description for an action
 * @param {Action} action - The action to describe
 * @returns {string} Detailed description with all properties
 */
function getActionDetails(action)
```

### Selection Management

```javascript
/**
 * Filters actions based on selection state
 * @param {Action[]} actions - All pending actions
 * @param {boolean[]} selections - Selection state for each action
 * @returns {Action[]} Only the selected actions
 */
function filterSelectedActions(actions, selections)

/**
 * Checks if any actions are selected
 * @param {boolean[]} selections - Selection state array
 * @returns {boolean} True if at least one action is selected
 */
function hasSelectedActions(selections)
```

### Excel Highlighting

```javascript
/**
 * Highlights a range in Excel
 * @param {string} rangeAddress - The range to highlight (e.g., "A1:B10")
 * @returns {Promise<boolean>} True if successful, false if range invalid
 */
async function highlightRange(rangeAddress)

/**
 * Clears any active highlighting
 * @returns {Promise<void>}
 */
async function clearHighlight()
```

## Data Models

### Action (existing, extended)

```javascript
/**
 * @typedef {Object} Action
 * @property {string} type - Action type: 'formula' | 'values' | 'format' | 'chart' | 'validation' | 'sort' | 'autofill'
 * @property {string} target - Target cell or range address
 * @property {string} [source] - Source range for validation/autofill
 * @property {string} [chartType] - Chart type for chart actions
 * @property {string} [title] - Chart title
 * @property {string} [position] - Chart position
 * @property {string} [data] - Action data (formula text, values JSON, format JSON)
 */

/**
 * @typedef {Object} PreviewState
 * @property {Action[]} actions - All pending actions
 * @property {boolean[]} selections - Selection state for each action (true = selected)
 * @property {number} expandedIndex - Index of currently expanded action (-1 if none)
 * @property {number} highlightedIndex - Index of currently highlighted action (-1 if none)
 */
```

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Preview renders all actions
*For any* non-empty array of actions, the preview panel rendering SHALL produce an element containing exactly one preview item for each action in the array.
**Validates: Requirements 1.1**

### Property 2: Action rendering includes required fields
*For any* action of a given type, the rendered preview item SHALL contain the target range and all type-specific required information (formula text for formulas, values for values, format properties for format, chart type/position for charts, source for validation).
**Validates: Requirements 1.2, 1.3, 1.4, 1.5, 1.6**

### Property 3: Each action has a checkbox
*For any* array of actions, the preview panel SHALL render exactly one checkbox input element for each action.
**Validates: Requirements 2.1**

### Property 4: Filter returns only selected actions
*For any* array of actions and corresponding selection states, the filterSelectedActions function SHALL return only the actions where the corresponding selection is true, preserving order.
**Validates: Requirements 2.2, 2.3**

### Property 5: Action type maps to distinct icon
*For any* two different action types, the getActionIcon function SHALL return different icon markup.
**Validates: Requirements 3.1**

### Property 6: Collapsed view shows summary
*For any* action rendered in collapsed state, the output SHALL contain the action type and target range.
**Validates: Requirements 3.4**

### Property 7: Expanded view shows full details
*For any* action rendered in expanded state, the output SHALL contain all non-null properties of the action.
**Validates: Requirements 3.3**

## Error Handling

| Error Scenario | Handling Strategy |
|----------------|-------------------|
| Invalid range address | Show warning icon on preview item, disable highlight on hover |
| Empty actions array | Hide preview panel entirely |
| Excel API unavailable | Disable highlighting, show tooltip explaining feature unavailable |
| Highlight timeout | Clear highlight after 100ms, continue without error |

## Testing Strategy

### Unit Tests
- Test `getActionIcon` returns valid SVG for each action type
- Test `getActionSummary` produces readable summaries
- Test `getActionDetails` includes all action properties
- Test `filterSelectedActions` with various selection combinations
- Test `hasSelectedActions` edge cases (empty, all false, all true, mixed)

### Property-Based Tests
Property-based tests will use fast-check library to verify correctness properties hold across many random inputs.

- **Property 1**: Generate random action arrays, verify preview item count matches
- **Property 2**: Generate random actions of each type, verify required fields present in output
- **Property 3**: Generate random action arrays, verify checkbox count matches action count
- **Property 4**: Generate random actions and selections, verify filter output correctness
- **Property 5**: Test all action type pairs return different icons
- **Property 6**: Generate random actions, verify collapsed output contains type and target
- **Property 7**: Generate random actions, verify expanded output contains all properties

Each property-based test will run a minimum of 100 iterations. Tests will be tagged with format: `**Feature: preview-before-apply, Property {number}: {property_text}**`
