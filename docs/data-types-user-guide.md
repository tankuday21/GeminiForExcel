# Excel Data Types User Guide

## What Are Data Types?

Entity cards that display structured data with properties. When you hover over a cell containing a data type, you see a card with additional information.

## Supported Operations

### Custom Entities (Fully Supported)

Create your own entity cards with custom properties:

**Example Prompts:**
- "Create a product entity in A2 with SKU P001, Price 29.99"
- "Insert an employee entity with name John, department Sales, email john@company.com"
- "Update the price in A2 to 34.99"

**What You Get:**
- Cell displays the entity name
- Hover to see all properties in a card
- Properties can be numbers, text, or true/false values

### Built-in Types (Manual Conversion Required)

**Stocks and Geography data types cannot be inserted automatically.**

**For Stocks:**
1. Ask: "Insert stock symbols MSFT, AAPL, GOOGL in B2:B4"
2. AI inserts the text values
3. You manually convert: Select cells → Data tab → Stocks

**For Geography:**
1. Ask: "Insert city names Seattle, New York, Chicago in C2:C4"
2. AI inserts the text values
3. You manually convert: Select cells → Data tab → Geography

**Why Manual?** This is an Office.js API limitation, not an add-in limitation.

## Examples

### Product Catalog
```
Prompt: "Create product entities in A2:A4 with SKU, Name, Price, and InStock properties"

Result: Entity cards with:
- A2: Product A (SKU: P001, Price: 29.99, InStock: true)
- A3: Product B (SKU: P002, Price: 49.99, InStock: true)
- A4: Product C (SKU: P003, Price: 19.99, InStock: false)
```

### Employee Records
```
Prompt: "Create employee entity for John Smith in B2 with ID 1001, Department Sales, Email john@company.com"

Result: Entity card showing:
- Cell displays: "John Smith"
- Hover card shows: ID, Department, Email
```

### Stock Tracking (Manual Steps Required)
```
Prompt: "Insert stock symbols MSFT, AAPL, GOOGL in B2:B4"

AI Response: "I've inserted the symbols. To convert to Stocks:
1. Select B2:B4
2. Click Data tab
3. Click Stocks button"
```

## Limitations

| Limitation | Details |
|------------|---------|
| Single cell per entity | Cannot insert entity across multiple cells |
| Max 5-10 properties | Card display becomes crowded with more |
| Performance | Keep <100 entities per sheet |
| Built-in types | Stocks/Geography require manual UI conversion |
| Excel version | Requires Excel 365, Excel 2021, or Excel Online |

## Troubleshooting

### "Data types not supported"
Your Excel version doesn't support data types. Requires Excel 365, Excel 2021, or Excel Online.

### "Entity not displaying"
Hover over the cell to see the entity card. The cell shows the text value; properties appear in the hover card.

### "Properties not updating"
Use the refresh command: "Update the price in A2 to 34.99"

### "Can't insert Stocks/Geography"
This is an API limitation. Insert text values and manually convert via Data tab → Data Types.

## Best Practices

1. **Use descriptive text values** - The text shown in the cell should be meaningful
2. **Limit properties** - 5-10 properties per entity for best card display
3. **Set basicValue** - Provides fallback for older Excel versions
4. **Group related data** - Use entities for items with multiple attributes
5. **Consider tables** - For simple tabular data, use Excel Tables instead
