# Excel Data Types - Limitations and Workarounds

## Critical Limitation: Built-in Stocks/Geography

**Office.js does NOT provide an API to insert built-in Stocks or Geography data types programmatically.**

These data types are UI-driven only and can only be created through:
- Excel UI: Insert > Data Types
- Data tab > Data Types > Stocks/Geography

### Workaround for Users

1. **AI inserts text values** (e.g., "MSFT", "Seattle, WA")
2. **User manually selects cells**
3. **User clicks Data tab > Data Types > Stocks (or Geography)**
4. **Excel converts text to linked entity**

### Example AI Response
```
I've inserted the stock symbols in cells B2:B5. To convert them to Stocks data type:
1. Select cells B2:B5
2. Click the Data tab
3. Click the Stocks button in the Data Types group
4. Excel will convert the text to linked stock data
```

## Custom EntityCellValue - Fully Supported

Custom entities created by add-ins are fully supported:
- Insert via `range.valuesAsJson`
- Update properties via `range.valuesAsJson`
- Read via `range.valuesAsJson` and `range.valueTypes`

### Supported Operations
| Operation | Support Level | Notes |
|-----------|--------------|-------|
| Insert custom entity | ✅ Full | Use `insertDataType` action |
| Read entity properties | ✅ Full | Via `valuesAsJson` |
| Update entity properties | ✅ Full | Use `refreshDataType` action |
| Insert Stocks | ❌ None | UI-only, provide workaround |
| Insert Geography | ❌ None | UI-only, provide workaround |
| Refresh LinkedEntity | ⚠️ Auto | Service handles refresh |

## Refresh Limitations

- **LinkedEntityCellValue**: Auto-refreshes from external service (Microsoft.Stocks, Microsoft.Geography)
- **Custom EntityCellValue**: Requires manual property updates via `refreshDataType`

## Card Layout Options

- **1-column layout**: Default, currently supported
- **2-column layout**: May 2025 feature, not yet available
- **Custom layouts**: Not supported via API

## Performance Guidelines

- **Limit entities**: <100 per sheet for optimal performance
- **Initial load**: ~50-100ms per entity cell
- **Detection sampling**: Use first 50x10 cells for large sheets
- **Batch operations**: Group entity insertions when possible

## convertToDataType Action - NOT SUPPORTED

**Do NOT implement** a `convertToDataType` action handler - it would fail at runtime.

The Office.js API does not support:
- Converting text to Stocks
- Converting text to Geography
- Programmatic data type conversion

### Alternative Approach
Use custom entities for structured data that doesn't require external service integration:
- Product catalogs
- Employee records
- Project tracking
- Custom business data
