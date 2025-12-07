# Feature Matrix - Excel AI Copilot

Complete reference of all 87 supported operations with Excel version requirements and API compatibility.

## Summary Statistics

| Metric | Value |
|--------|-------|
| Total Actions | 87 |
| Fully Supported | 85 (97.7%) |
| Partially Supported | 2 (2.3%) |
| Works on All Versions | 47 (54%) |
| Requires Excel 2019+ | 25 actions |
| Requires Excel 365 | 15 actions |

---

## Legend

### Status
- ✅ **Full**: Fully implemented and tested
- ⚠️ **Partial**: Implemented with known limitations
- ❌ **Not Supported**: API limitation, cannot implement

### Excel Version
- **All**: Excel 2016+, Excel 365, Excel Online
- **2019+**: Excel 2019, Excel 365, Excel Online
- **365**: Microsoft 365 subscription required
- **2021+**: Excel 2021, Excel 365, Excel Online

### Office.js API Version
- **1.1+**: All supported platforms
- **1.7+**: Excel 2019+
- **1.10+**: Excel 2019+ (sparklines)
- **1.11+**: Excel 365 (comments)
- **1.14+**: Excel 365 (custom views)
- **1.16+**: Excel 2021+ (data types)

---

## Basic Operations (6 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `formula` | Apply formula to cell/range | All | 1.1+ | ✅ Full | Circular refs not detected |
| `values` | Insert values (2D array) | All | 1.1+ | ✅ Full | Max 1M rows |
| `format` | Apply cell formatting | All | 1.1+ | ✅ Full | Some styles require 1.2+ |
| `validation` | Add data validation rules | All | 1.1+ | ✅ Full | Custom formulas limited |
| `sort` | Sort range by columns | All | 1.1+ | ✅ Full | Max 64 sort levels |
| `autofill` | Auto-fill patterns/series | All | 1.4+ | ✅ Full | Pattern detection varies |

---

## Advanced Formatting (2 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `conditionalFormat` | Apply conditional formatting | All | 1.6+ | ✅ Full | 8 format types supported |
| `clearFormat` | Remove formatting from range | All | 1.1+ | ✅ Full | None |

### Conditional Format Types
| Type | Description | Example |
|------|-------------|---------|
| cellValue | Compare cell values | Greater than, less than, between |
| colorScale | 2-color or 3-color gradient | Red to green heatmap |
| dataBar | Progress bar visualization | Sales progress |
| iconSet | Icon indicators | Traffic lights, arrows |
| topBottom | Top/bottom N items | Top 10 sales |
| preset | Predefined rules | Duplicates, blanks, errors |
| textComparison | Text matching | Contains, begins with |
| custom | Formula-based | `=MOD(ROW(),2)=0` |

---

## Charts (2 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `chart` | Create chart from data | All | 1.1+ | ✅ Full | 20+ chart types |
| `pivotChart` | Create chart from PivotTable | All | 1.1+ | ✅ Full | Requires existing pivot |

### Supported Chart Types
Column, Bar, Line, Pie, Doughnut, Area, Scatter, Radar, Combo, Waterfall, Funnel, Stock, Surface, Treemap, Sunburst, Histogram, Box & Whisker, Pareto

---

## Copy/Filter/Duplicates (5 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `copy` | Copy range with formatting | All | 1.1+ | ✅ Full | None |
| `copyValues` | Copy values only | All | 1.1+ | ✅ Full | None |
| `filter` | Apply AutoFilter | All | 1.1+ | ✅ Full | Single filter per sheet |
| `clearFilter` | Remove AutoFilter | All | 1.1+ | ✅ Full | None |
| `removeDuplicates` | Remove duplicate rows | All | 1.2+ | ✅ Full | Compares selected columns |

---

## Sheet Management (1 Action)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `sheet` | Add/delete/copy worksheet | All | 1.1+ | ✅ Full | 31-char name limit |

---

## Table Operations (7 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `createTable` | Convert range to table | All | 1.1+ | ✅ Full | Cannot overlap tables |
| `styleTable` | Apply table style | All | 1.1+ | ✅ Full | 60 built-in styles |
| `addTableRow` | Add rows to table | All | 1.1+ | ✅ Full | None |
| `addTableColumn` | Add columns to table | All | 1.1+ | ✅ Full | None |
| `resizeTable` | Change table range | All | 1.1+ | ✅ Full | Cannot overlap |
| `convertToRange` | Convert table to range | All | 1.1+ | ✅ Full | Preserves data |
| `toggleTableTotals` | Show/hide totals row | All | 1.1+ | ✅ Full | 11 aggregation functions |

---

## Data Manipulation (8 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `insertRows` | Insert rows at position | All | 1.1+ | ✅ Full | Shifts formulas |
| `insertColumns` | Insert columns at position | All | 1.1+ | ✅ Full | Shifts formulas |
| `deleteRows` | Delete specified rows | All | 1.1+ | ✅ Full | May cause #REF! |
| `deleteColumns` | Delete specified columns | All | 1.1+ | ✅ Full | May cause #REF! |
| `mergeCells` | Merge range into one cell | All | 1.1+ | ✅ Full | Keeps top-left value |
| `unmergeCells` | Split merged cells | All | 1.1+ | ✅ Full | None |
| `findReplace` | Find and replace text | All | 1.1+ | ✅ Full | Regex supported |
| `textToColumns` | Split text by delimiter | All | 1.1+ | ✅ Full | Overwrites adjacent |

---

## PivotTable Operations (5 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `createPivotTable` | Create PivotTable | All | 1.3+ | ✅ Full | Contiguous source only |
| `addPivotField` | Add field to pivot area | All | 1.3+ | ✅ Full | 4 areas supported |
| `configurePivotLayout` | Set layout options | All | 1.3+ | ✅ Full | 3 layout types |
| `refreshPivotTable` | Refresh pivot data | All | 1.3+ | ✅ Full | None |
| `deletePivotTable` | Remove PivotTable | All | 1.3+ | ✅ Full | None |

---

## Slicer Operations (5 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `createSlicer` | Create slicer for table/pivot | 2019+ | 1.10+ | ✅ Full | One per field |
| `configureSlicer` | Change slicer settings | 2019+ | 1.10+ | ✅ Full | 18 styles |
| `connectSlicerToTable` | Connect to table | 2019+ | 1.10+ | ✅ Full | Same workbook |
| `connectSlicerToPivot` | Connect to PivotTable | 2019+ | 1.10+ | ✅ Full | Same workbook |
| `deleteSlicer` | Remove slicer | 2019+ | 1.10+ | ✅ Full | None |

---

## Named Range Operations (4 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `createNamedRange` | Create named range/constant | All | 1.1+ | ✅ Full | 255-char name limit |
| `deleteNamedRange` | Remove named range | All | 1.1+ | ✅ Full | May cause #NAME! |
| `updateNamedRange` | Update range reference | All | 1.1+ | ✅ Full | None |
| `listNamedRanges` | List all named ranges | All | 1.1+ | ✅ Full | None |

---

## Protection Operations (6 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `protectWorksheet` | Protect sheet | All | 1.2+ | ✅ Full | Basic protection only |
| `unprotectWorksheet` | Remove sheet protection | All | 1.2+ | ✅ Full | Requires password |
| `protectRange` | Lock specific cells | All | 1.2+ | ✅ Full | Requires sheet protection |
| `unprotectRange` | Unlock specific cells | All | 1.2+ | ✅ Full | None |
| `protectWorkbook` | Protect workbook structure | All | 1.7+ | ✅ Full | Basic protection only |
| `unprotectWorkbook` | Remove workbook protection | All | 1.7+ | ✅ Full | Requires password |

---

## Shape Operations (8 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `insertShape` | Add geometric shape | All | 1.9+ | ✅ Full | 20+ shape types |
| `insertImage` | Add Base64 image | All | 1.9+ | ✅ Full | ~5MB limit |
| `insertTextBox` | Add text box | All | 1.9+ | ✅ Full | None |
| `formatShape` | Change shape properties | All | 1.9+ | ✅ Full | Fill, line, text |
| `deleteShape` | Remove shape | All | 1.9+ | ✅ Full | None |
| `groupShapes` | Group multiple shapes | All | 1.9+ | ✅ Full | 2+ shapes required |
| `arrangeShapes` | Change z-order | All | 1.9+ | ✅ Full | 4 positions |
| `ungroupShapes` | Ungroup shapes | All | 1.9+ | ✅ Full | None |

---

## Comment Operations (8 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `addComment` | Add threaded comment | 365 | 1.11+ | ✅ Full | Modern comments |
| `addNote` | Add legacy note | All | 1.1+ | ✅ Full | Yellow sticky |
| `editComment` | Modify comment | 365 | 1.11+ | ✅ Full | None |
| `editNote` | Modify note | All | 1.1+ | ✅ Full | None |
| `deleteComment` | Remove comment thread | 365 | 1.11+ | ✅ Full | None |
| `deleteNote` | Remove note | All | 1.1+ | ✅ Full | None |
| `replyToComment` | Add reply to thread | 365 | 1.11+ | ✅ Full | None |
| `resolveComment` | Mark as resolved | 365 | 1.11+ | ✅ Full | None |

---

## Sparkline Operations (3 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `createSparkline` | Create in-cell chart | 2019+ | 1.10+ | ✅ Full | Line, Column, Win/Loss |
| `configureSparkline` | Change sparkline settings | 2019+ | 1.10+ | ✅ Full | Markers, colors, axis |
| `deleteSparkline` | Remove sparkline | 2019+ | 1.10+ | ✅ Full | None |

---

## Worksheet Management (9 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `renameSheet` | Change sheet name | All | 1.1+ | ✅ Full | 31-char limit |
| `moveSheet` | Reposition sheet | All | 1.1+ | ✅ Full | 0-based index |
| `hideSheet` | Hide sheet | All | 1.1+ | ✅ Full | Cannot hide last |
| `unhideSheet` | Show hidden sheet | All | 1.1+ | ✅ Full | None |
| `freezePanes` | Freeze rows/columns | All | 1.7+ | ✅ Full | None |
| `unfreezePane` | Remove freeze | All | 1.7+ | ✅ Full | None |
| `setZoom` | Set zoom level | All | 1.1+ | ✅ Full | 10-400% |
| `splitPane` | Split view | All | 1.7+ | ✅ Full | None |
| `createView` | Create custom view | 365 | 1.14+ | ⚠️ Partial | Limited API support |

---

## Page Setup Operations (6 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `setPageSetup` | Set page options | All | 1.9+ | ✅ Full | Orientation, size, scale |
| `setPageMargins` | Set margin sizes | All | 1.9+ | ✅ Full | 0-10 inches |
| `setPageOrientation` | Portrait/Landscape | All | 1.9+ | ✅ Full | None |
| `setPrintArea` | Define print range | All | 1.9+ | ✅ Full | Multiple ranges |
| `setHeaderFooter` | Add headers/footers | All | 1.9+ | ✅ Full | Dynamic fields |
| `setPageBreaks` | Insert page breaks | All | 1.9+ | ✅ Full | Manual only |

---

## Hyperlink Operations (3 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `addHyperlink` | Add link to cell | All | 1.7+ | ✅ Full | URL, email, internal |
| `removeHyperlink` | Remove link | All | 1.7+ | ✅ Full | Preserves value |
| `editHyperlink` | Modify link | All | 1.7+ | ✅ Full | None |

---

## Data Type Operations (2 Actions)

| Action | Description | Excel Version | API | Status | Limitations |
|--------|-------------|---------------|-----|--------|-------------|
| `insertDataType` | Create custom entity | 2021+ | 1.16+ | ⚠️ Partial | Custom entities only |
| `refreshDataType` | Update entity properties | 2021+ | 1.16+ | ⚠️ Partial | Custom entities only |

**Note**: Stocks and Geography data types cannot be created via API. Users must manually convert cells using Excel's Data tab.

---

## Quick Reference

### Most Commonly Used Actions
1. `formula` - Apply formulas
2. `values` - Insert data
3. `format` - Cell formatting
4. `createTable` - Create tables
5. `chart` - Create charts
6. `conditionalFormat` - Conditional formatting
7. `sort` - Sort data
8. `filter` - Filter data
9. `createPivotTable` - Create PivotTables
10. `findReplace` - Find and replace

### Actions Requiring User Confirmation
- `deleteRows` / `deleteColumns` - Destructive
- `findReplace` with `replaceAll` - Bulk changes
- `removeDuplicates` - Data removal
- `convertToRange` - Removes table features
- `clearFormat` - Removes all formatting

### Actions with Performance Considerations
- `values` with >50K rows
- `conditionalFormat` on >10K cells
- `createPivotTable` with >100K source rows
- `chart` with >5K data points
- `createSparkline` with >100 sparklines

---

## Version Compatibility Matrix

| Feature Category | Excel 2016 | Excel 2019 | Excel 2021 | Excel 365 | Excel Online |
|-----------------|------------|------------|------------|-----------|--------------|
| Basic Operations | ✅ | ✅ | ✅ | ✅ | ✅ |
| Tables | ✅ | ✅ | ✅ | ✅ | ✅ |
| Charts | ✅ | ✅ | ✅ | ✅ | ✅ |
| PivotTables | ✅ | ✅ | ✅ | ✅ | ✅ |
| Slicers | ❌ | ✅ | ✅ | ✅ | ✅ |
| Sparklines | ❌ | ✅ | ✅ | ✅ | ✅ |
| Comments (Modern) | ❌ | ❌ | ❌ | ✅ | ✅ |
| Data Types | ❌ | ❌ | ✅ | ✅ | ✅ |
| Custom Views | ❌ | ❌ | ❌ | ✅ | ✅ |
| Dynamic Arrays | ❌ | ❌ | ✅ | ✅ | ✅ |

---

## Related Documentation

- [SETUP.md](../SETUP.md) - Installation and feature guide
- [Example Prompts](example-prompts.md) - 500+ prompt examples
- [Testing Guide](testing-guide.md) - Developer testing documentation
- [Data Types User Guide](data-types-user-guide.md) - Data types documentation
- [Data Types Limitations](data-types-limitations.md) - Known limitations
