# Excel Data Types API Research

## Overview
Office.js provides partial support for Excel data types through the `Range.valuesAsJson` API (ExcelApi 1.16+, Excel 365/2021+).

## API Version Requirements
- **ExcelApi 1.16+**: EntityCellValue read/write support
- **ExcelApi 1.17+**: Custom provider support (May 2025)
- **Supported Platforms**: Excel 365, Excel 2021, Excel Online

## Range.valuesAsJson API

### Reading Data Types
```javascript
range.load(["valueTypes", "valuesAsJson"]);
await ctx.sync();

// valueTypes returns: "Entity", "LinkedEntity", "FormattedNumber", "WebImage", etc.
const cellType = range.valueTypes[0][0];

// valuesAsJson returns full EntityCellValue objects
const cellValue = range.valuesAsJson[0][0];
```

### Writing Custom Entities
```javascript
const entityValue = {
    type: "Entity",
    text: "Product A",
    basicType: "String",
    basicValue: "Product A",
    properties: {
        SKU: { type: "String", basicValue: "P001" },
        Price: { type: "Double", basicValue: 29.99 },
        InStock: { type: "Boolean", basicValue: true }
    }
};

range.valuesAsJson = [[entityValue]];
await ctx.sync();
```

## EntityCellValue Structure
```json
{
    "type": "Entity",
    "text": "Display Text",
    "basicType": "String",
    "basicValue": "Fallback Value",
    "properties": {
        "PropertyName": {
            "type": "String|Double|Boolean",
            "basicValue": "value"
        }
    }
}
```

## LinkedEntityCellValue Structure
```json
{
    "type": "LinkedEntity",
    "text": "MSFT",
    "id": "unique-id",
    "serviceId": "Microsoft.Stocks",
    "properties": {
        "Price": { "type": "Double", "basicValue": 350.00 },
        "Change": { "type": "Double", "basicValue": 2.50 }
    }
}
```

## Built-in Providers
- **Stocks**: serviceId = "Microsoft.Stocks"
- **Geography**: serviceId = "Microsoft.Geography"

## Property Types
- `String`: Text values
- `Double`: Numeric values
- `Boolean`: True/false values

## Detection Code Example
```javascript
async function detectDataTypes(ctx, range) {
    range.load(["valueTypes", "valuesAsJson"]);
    await ctx.sync();
    
    const dataTypeCells = [];
    for (let r = 0; r < range.valueTypes.length; r++) {
        for (let c = 0; c < range.valueTypes[r].length; c++) {
            const cellType = range.valueTypes[r][c];
            if (cellType === "Entity" || cellType === "LinkedEntity") {
                const cellValue = range.valuesAsJson[r][c];
                dataTypeCells.push({
                    row: r,
                    col: c,
                    type: cellType,
                    text: cellValue.text,
                    properties: Object.keys(cellValue.properties || {})
                });
            }
        }
    }
    return dataTypeCells;
}
```

## Performance Considerations
- Entity cards add ~50-100ms per cell on initial load
- Limit to <100 entities per sheet for optimal performance
- Use sampling (e.g., first 50x10 cells) for detection in large sheets
