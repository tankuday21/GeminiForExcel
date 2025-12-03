# Design Document

## Overview

The Data Validation Suggestions feature adds automatic data quality analysis to the Excel AI Copilot. It scans the current data range, detects common issues (duplicates, missing values, format inconsistencies, outliers), and presents actionable suggestions in a dedicated validation panel. Users can review issues by severity, apply one-click fixes, and track changes through the existing undo system.

## Architecture

The feature integrates into the existing taskpane architecture:

```
┌─────────────────────────────────────────┐
│           Taskpane UI                    │
├─────────────────────────────────────────┤
│  ┌─────────────────────────────────┐    │
│  │     Validation Panel            │    │
│  │  ┌───────────────────────────┐  │    │
│  │  │ Issue Summary (counts)    │  │    │
│  │  ├───────────────────────────┤  │    │
│  │  │ Severity Filters          │  │    │
│  │  ├───────────────────────────┤  │    │
│  │  │ Issue Cards List          │  │    │
│  │  │  - Issue Card 1           │  │    │
│  │  │  - Issue Card 2           │  │    │
│  │  │  - ...                    │  │    │
│  │  └───────────────────────────┘  │    │
│  └─────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

## Components and Interfaces

### Data Structures

```typescript
type IssueSeverity = 'error' | 'warning' | 'info';

type IssueType = 'duplicate' | 'missing' | 'format' | 'outlier' | 'type-mismatch';

interface ValidationIssue {
    id: string;
    type: IssueType;
    severity: IssueSeverity;
    column: string;           // Column letter (e.g., "A", "B")
    columnHeader: string;     // Column header name
    affectedCells: string[];  // Cell addresses (e.g., ["A2", "A5"])
    value?: any;              // The problematic value (for duplicates/outliers)
    count: number;            // Number of affected cells
    message: string;          // Human-readable description
    details?: {               // Type-specific details
        duplicateCount?: number;
        mean?: number;
        stdDev?: number;
        deviation?: number;
        formatExamples?: string[];
    };
    fixAvailable: boolean;
    fixAction?: FixAction;
}

interface FixAction {
    type: 'remove-duplicates' | 'fill-empty' | 'standardize-format' | 'highlight';
    label: string;
    execute: () => Promise<void>;
}

interface ValidationResult {
    issues: ValidationIssue[];
    scannedRows: number;
    scannedColumns: number;
    timestamp: Date;
}

interface ValidationState {
    isScanning: boolean;
    result: ValidationResult | null;
    activeFilter: IssueSeverity | 'all';
    expandedIssueId: string | null;
}
```

### Core Functions

```typescript
// Main validation entry point
async function validateData(): Promise<ValidationResult>

// Individual detectors
function detectDuplicates(values: any[][], columnMap: ColumnInfo[]): ValidationIssue[]
function detectMissingValues(values: any[][], columnMap: ColumnInfo[]): ValidationIssue[]
function detectFormatIssues(values: any[][], columnMap: ColumnInfo[]): ValidationIssue[]
function detectOutliers(values: any[][], columnMap: ColumnInfo[]): ValidationIssue[]

// Issue management
function sortIssuesBySeverity(issues: ValidationIssue[]): ValidationIssue[]
function filterIssuesBySeverity(issues: ValidationIssue[], severity: IssueSeverity): ValidationIssue[]
function countIssuesBySeverity(issues: ValidationIssue[]): Record<IssueSeverity, number>

// UI rendering
function renderValidationPanel(result: ValidationResult, filter: IssueSeverity | 'all'): void
function renderIssueCard(issue: ValidationIssue, expanded: boolean): string
function renderIssueSummary(counts: Record<IssueSeverity, number>): string

// Fix actions
async function applyFix(issue: ValidationIssue): Promise<void>
function createFixAction(issue: ValidationIssue): FixAction | null
```

## Data Models

### Severity Priority
- `error`: Critical issues that likely indicate data corruption (priority 1)
- `warning`: Issues that may cause problems (duplicates, missing values) (priority 2)
- `info`: Suggestions for improvement (outliers, casing) (priority 3)

### Detection Thresholds
- Outliers: Values > 3 standard deviations from mean
- Duplicates: Any value appearing more than once in a column
- Missing: Empty cells or cells with only whitespace
- Format inconsistency: Mixed data types in same column (>10% different)

## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Duplicate detection finds all duplicates
*For any* column of values containing known duplicate entries, the detectDuplicates function SHALL return issues that collectively reference all cells containing those duplicate values.
**Validates: Requirements 2.1**

### Property 2: Missing value detection finds all empty cells
*For any* column of values containing empty cells, the detectMissingValues function SHALL return issues that collectively reference all empty cell positions.
**Validates: Requirements 3.1**

### Property 3: Outlier detection uses correct statistical threshold
*For any* column of numeric values, the detectOutliers function SHALL only flag values where |value - mean| > 3 * standardDeviation.
**Validates: Requirements 5.1**

### Property 4: Issue severity count is accurate
*For any* list of validation issues, the countIssuesBySeverity function SHALL return counts where the sum of all severity counts equals the total issue count.
**Validates: Requirements 1.2**

### Property 5: Issues are sorted by severity priority
*For any* list of validation issues, after sorting, all 'error' issues SHALL appear before 'warning' issues, and all 'warning' issues SHALL appear before 'info' issues.
**Validates: Requirements 7.1**

### Property 6: Severity filter returns only matching issues
*For any* list of validation issues and a severity filter, the filterIssuesBySeverity function SHALL return only issues with the specified severity.
**Validates: Requirements 7.2**

### Property 7: Issue card rendering includes required fields
*For any* validation issue, the rendered issue card SHALL contain the column name, affected cell count, and issue message.
**Validates: Requirements 2.2, 3.2, 4.3, 5.2**

### Property 8: Fix button presence matches fixAvailable flag
*For any* validation issue, the rendered issue card SHALL display a "Fix" button if and only if fixAvailable is true.
**Validates: Requirements 6.1**

### Property 9: Format inconsistency detection identifies mixed types
*For any* column containing both text and numeric values (excluding headers), the detectFormatIssues function SHALL create an issue when the minority type exceeds 10% of values.
**Validates: Requirements 4.1**

## Error Handling

| Scenario | Handling |
|----------|----------|
| No data in worksheet | Display message "No data to validate" |
| Excel API unavailable | Show error toast, disable validation button |
| Very large dataset (>10000 rows) | Show warning, offer to scan first 10000 rows |
| Fix action fails | Show error message, do not add to undo history |
| Invalid cell reference | Skip issue, log warning to console |

## Testing Strategy

### Unit Tests
- Test each detector function with known inputs
- Test sorting and filtering functions
- Test issue card rendering output

### Property-Based Tests
Using fast-check library:
- Property 1: Generate random arrays with injected duplicates, verify all found
- Property 2: Generate random arrays with injected empty cells, verify all found
- Property 3: Generate random numeric arrays, verify outlier threshold math
- Property 4: Generate random issue lists, verify count sum equals total
- Property 5: Generate random issue lists, verify sort order
- Property 6: Generate random issue lists, verify filter correctness
- Property 7: Generate random issues, verify rendered output contains fields
- Property 8: Generate random issues with/without fixes, verify button presence
- Property 9: Generate mixed-type columns, verify detection threshold

Each property-based test will run a minimum of 100 iterations. Tests will be tagged with format: `**Feature: data-validation-suggestions, Property {number}: {property_text}**`
