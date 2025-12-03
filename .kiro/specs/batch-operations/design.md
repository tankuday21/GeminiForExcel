# Design Document

## Overview

The Batch Operations feature enables users to apply AI-generated actions across multiple sheets in an Excel workbook. Users can select target sheets via a sheet selector UI, preview changes for each sheet, execute operations with progress feedback, and undo entire batches as a single operation. The feature handles errors gracefully, allowing processing to continue even when individual sheets fail.

## Architecture

The feature extends the existing taskpane with batch mode capabilities:

```
┌─────────────────────────────────────────┐
│           Taskpane UI                    │
├─────────────────────────────────────────┤
│  ┌─────────────────────────────────┐    │
│  │  [Batch Mode Toggle]            │    │
│  └─────────────────────────────────┘    │
│  ┌─────────────────────────────────┐    │
│  │     Sheet Selector              │    │
│  │  ☑ Sheet1                       │    │
│  │  ☑ Sheet2                       │    │
│  │  ☐ Sheet3                       │    │
│  │  [Select All] [Select None]     │    │
│  │  Selected: 2 of 3               │    │
│  └─────────────────────────────────┘    │
│  ┌─────────────────────────────────┐    │
│  │     Batch Preview               │    │
│  │  ▼ Sheet1 (3 actions)           │    │
│  │    - Formula → A10              │    │
│  │    - Format → B1:B20            │    │
│  │  ▶ Sheet2 (3 actions)           │    │
│  └─────────────────────────────────┘    │
│  ┌─────────────────────────────────┐    │
│  │  [Apply to All Sheets]          │    │
│  └─────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

## Components and Interfaces

### Data Structures

```typescript
interface SheetInfo {
    name: string;
    index: number;
    selected: boolean;
}

interface BatchAction {
    sheetName: string;
    actions: Action[];  // Reuses existing Action type
}

type SheetResultStatus = 'pending' | 'processing' | 'success' | 'error' | 'skipped';

interface SheetResult {
    sheetName: string;
    status: SheetResultStatus;
    error?: string;
    actionsApplied: number;
    undoData?: any;  // Data needed to revert changes
}

interface BatchResult {
    results: SheetResult[];
    totalSheets: number;
    successCount: number;
    errorCount: number;
    skippedCount: number;
    startTime: Date;
    endTime?: Date;
}

interface BatchProgress {
    currentSheet: string;
    currentIndex: number;
    totalSheets: number;
    percentage: number;
}

interface BatchSettings {
    skipEmptySheets: boolean;
    stopOnError: boolean;
}

interface BatchState {
    enabled: boolean;
    sheets: SheetInfo[];
    preview: BatchAction[];
    isProcessing: boolean;
    progress: BatchProgress | null;
    result: BatchResult | null;
    settings: BatchSettings;
    cancelled: boolean;
}
```

### Core Functions

```typescript
// Sheet management
async function getWorkbookSheets(): Promise<SheetInfo[]>
function selectAllSheets(sheets: SheetInfo[]): SheetInfo[]
function selectNoSheets(sheets: SheetInfo[]): SheetInfo[]
function getSelectedSheets(sheets: SheetInfo[]): SheetInfo[]
function countSelectedSheets(sheets: SheetInfo[]): number

// Batch preview
function groupActionsBySheet(actions: Action[], sheets: SheetInfo[]): BatchAction[]
function renderBatchPreview(batchActions: BatchAction[]): string
function renderSheetSection(batchAction: BatchAction, expanded: boolean): string

// Batch execution
async function executeBatch(
    batchActions: BatchAction[], 
    settings: BatchSettings,
    onProgress: (progress: BatchProgress) => void
): Promise<BatchResult>
async function applyToSheet(sheetName: string, actions: Action[]): Promise<SheetResult>
function calculateProgress(currentIndex: number, total: number): BatchProgress

// Error handling
function collectErrors(results: SheetResult[]): SheetResult[]
function getFailedSheets(result: BatchResult): string[]
async function retryFailed(result: BatchResult, batchActions: BatchAction[]): Promise<BatchResult>

// Undo support
function createBatchHistoryEntry(result: BatchResult): HistoryEntry
async function undoBatch(entry: HistoryEntry): Promise<void>

// Settings
function saveBatchSettings(settings: BatchSettings): void
function loadBatchSettings(): BatchSettings

// UI rendering
function renderSheetSelector(sheets: SheetInfo[]): string
function renderBatchProgress(progress: BatchProgress): string
function renderBatchSummary(result: BatchResult): string
```

## Data Models

### Batch Processing Flow
1. User enables batch mode
2. System loads all sheets from workbook
3. User selects target sheets
4. User sends prompt to AI
5. AI generates actions (same as single-sheet mode)
6. System groups actions by sheet for preview
7. User reviews and clicks "Apply to All Sheets"
8. System processes each sheet sequentially
9. Progress updates shown in real-time
10. Summary displayed on completion

### Settings Defaults
- `skipEmptySheets`: true (skip sheets where actions don't apply)
- `stopOnError`: false (continue processing on failure)

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Sheet selector renders all sheets
*For any* list of workbook sheets, the rendered sheet selector SHALL contain an entry for each sheet name.
**Validates: Requirements 1.2**

### Property 2: Selected count matches selection state
*For any* list of sheets with selection states, the countSelectedSheets function SHALL return the exact count of sheets where selected is true.
**Validates: Requirements 1.3**

### Property 3: Select all sets all selections to true
*For any* list of sheets, after calling selectAllSheets, every sheet SHALL have selected equal to true.
**Validates: Requirements 1.4**

### Property 4: Select none sets all selections to false
*For any* list of sheets, after calling selectNoSheets, every sheet SHALL have selected equal to false.
**Validates: Requirements 1.5**

### Property 5: Batch preview groups actions by sheet
*For any* list of actions and sheets, the groupActionsBySheet function SHALL return batch actions where each group contains only actions for that specific sheet.
**Validates: Requirements 2.1**

### Property 6: Progress percentage is accurate
*For any* current index and total count, the calculateProgress function SHALL return percentage equal to (currentIndex / totalSheets) * 100.
**Validates: Requirements 3.2**

### Property 7: Result status reflects outcome
*For any* sheet processing outcome, the SheetResult status SHALL be 'success' when no error occurs and 'error' when an error occurs, with error message populated.
**Validates: Requirements 3.3, 3.4**

### Property 8: Summary counts match results
*For any* batch result, the successCount plus errorCount plus skippedCount SHALL equal totalSheets.
**Validates: Requirements 3.5**

### Property 9: Continue on error processes all sheets
*For any* batch with stopOnError set to false, when a sheet fails, the batch processor SHALL continue and process all remaining selected sheets.
**Validates: Requirements 4.1**

### Property 10: Retry processes only failed sheets
*For any* batch result with failed sheets, the retryFailed function SHALL only attempt to process sheets with status 'error'.
**Validates: Requirements 4.3**

### Property 11: Batch creates single history entry
*For any* successful batch operation affecting multiple sheets, exactly one history entry SHALL be created containing all affected sheet names.
**Validates: Requirements 5.1**

### Property 12: Settings persist correctly
*For any* batch settings saved to storage, loading settings SHALL return the same values that were saved.
**Validates: Requirements 6.3**

## Error Handling

| Scenario | Handling |
|----------|----------|
| No sheets in workbook | Display message "No sheets available" |
| No sheets selected | Disable "Apply to All Sheets" button |
| Sheet deleted during batch | Mark as error, continue processing |
| Excel API unavailable | Show error toast, abort batch |
| User cancels mid-batch | Stop processing, show partial results |
| All sheets fail | Show error summary, offer retry option |

## Testing Strategy

### Unit Tests
- Test sheet selection functions
- Test progress calculation
- Test action grouping by sheet
- Test result summary calculation

### Property-Based Tests
Using fast-check library:
- Property 1: Generate random sheet lists, verify all rendered
- Property 2: Generate random selection states, verify count
- Property 3: Generate random sheets, verify selectAll result
- Property 4: Generate random sheets, verify selectNone result
- Property 5: Generate random actions/sheets, verify grouping
- Property 6: Generate random index/total pairs, verify percentage
- Property 7: Generate random outcomes, verify status mapping
- Property 8: Generate random results, verify count sum
- Property 9: Generate batch with failures, verify continuation
- Property 10: Generate results with failures, verify retry targets
- Property 11: Generate batch operations, verify single history entry
- Property 12: Generate random settings, verify round-trip persistence

Each property-based test will run a minimum of 100 iterations. Tests will be tagged with format: `**Feature: batch-operations, Property {number}: {property_text}**`
