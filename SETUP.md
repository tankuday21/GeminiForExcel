# Excel AI Copilot - Setup Guide

## 🚀 Two Ways to Use This Add-in

### Option A: Development Mode (localhost)
For testing and development on your machine.

### Option B: Production Deployment (Recommended)
Deploy once, use anywhere - like the built-in Copilot!

---

## Option A: Development Mode

### 1. Get a Gemini API Key
1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. Create a new API key
3. Copy the key

### 2. Install & Run
```bash
npm install
npm run start
```

### 3. Configure API Key in Excel
Click "AI Copilot" → Enter your API key when prompted

**Security Note:** API keys are stored with basic obfuscation in browser storage. For maximum security:
- Use the "Remove API Key" button in Settings when not actively using the add-in
- Consider re-entering your API key each session rather than storing it
- Never share your API key or workbooks that might contain stored keys

---

## Option B: Production Deployment (GitHub Pages - FREE)

### Step 1: Create GitHub Repository
1. Go to [github.com/new](https://github.com/new)
2. Create a new repo named `excel-ai-copilot`
3. Make it **Public** (required for free GitHub Pages)

### Step 2: Build & Push
```bash
# Build production files
npm run build

# Initialize git (if not already)
git init
git add .
git commit -m "Initial commit"

# Push to GitHub
git remote add origin https://github.com/YOUR_USERNAME/excel-ai-copilot.git
git branch -M main
git push -u origin main
```

### Step 3: Enable GitHub Pages
1. Go to your repo → Settings → Pages
2. Source: **Deploy from a branch**
3. Branch: **main** → **/dist** folder
4. Click Save
5. Wait 2-3 minutes for deployment

### Step 4: Update Manifest
1. Open `manifest.prod.xml`
2. Replace all `YOUR_GITHUB_USERNAME` with your GitHub username
3. Replace all `YOUR_REPO_NAME` with `excel-ai-copilot`
4. Save the file

### Step 5: Install the Add-in Permanently
**For Personal Use:**
1. Open Excel → Insert → Get Add-ins → My Add-ins
2. Click "Upload My Add-in"
3. Upload your `manifest.prod.xml`

**For Organization-wide:**
1. Go to Microsoft 365 Admin Center
2. Settings → Integrated Apps → Upload custom apps
3. Upload `manifest.prod.xml`

Now the add-in works in ANY workbook, on ANY device with your Microsoft account!

---

## Features

### Quick Actions
- **Analyze Data** - Get insights, patterns, and statistics
- **Create Formula** - Generate Excel formulas from natural language
- **Create Chart** - Get chart recommendations and create visualizations
- **Format Data** - Apply formatting and conditional formatting
- **Clean Data** - Find duplicates, fix inconsistencies
- **Summarize** - Create summaries with totals and averages

### Chat Interface
- Ask any question about your data
- Request specific formulas or calculations
- Get explanations of Excel functions
- Ask for data transformation help

### Apply Changes
When the AI suggests modifications, click "Apply Changes to Sheet" to execute them.

## Usage Tips

1. **Select data first** - The AI works best when you select the relevant data range
2. **Be specific** - "Sum column B" works better than "add up the numbers"
3. **Use quick actions** - They're optimized prompts for common tasks
4. **Review before applying** - Always check the AI's suggestion before applying

## Security

### API Key Storage
- API keys are stored with basic obfuscation (base64 encoding) in browser localStorage
- This is NOT encryption - it only prevents casual viewing
- For maximum security, use the "Remove API Key" button in Settings when done
- You may need to re-enter your API key after clearing browser data

### Removing Your API Key
1. Open Settings (gear icon)
2. Click "Remove API Key" button
3. Your key will be cleared from storage immediately

## Sparklines

Sparklines are compact, in-cell visualizations for trend analysis. The add-in supports creating, configuring, and deleting sparklines.

### Sparkline Types
- **Line**: Best for trends over time, continuous data (5+ data points)
- **Column**: Best for comparing magnitudes, discrete values
- **Win/Loss**: Best for binary outcomes (positive/negative, win/loss)

### Requirements and Limitations
- **API Version**: Requires ExcelApi 1.10+ (Excel 365, Excel 2019+, or Excel Online)
- **Source Data**: Must be a contiguous range (e.g., `B2:F2` for row, `C3:C20` for column)
- **Performance**: Limit to 50-100 sparklines per sheet for optimal performance
- **Protected Sheets**: Cannot add/modify sparklines on protected worksheets

### Best Practices
1. Place sparklines adjacent to data (typically in the last column of a table)
2. Use consistent sparkline types within the same table
3. Enable markers for high/low points in Line sparklines
4. Use colorblind-friendly color schemes
5. Consider data bars as an alternative for single-value magnitude comparisons

### Example Prompts
- "Create a line sparkline in G2 showing the trend from B2 to F2"
- "Add column sparklines for each row showing monthly sales"
- "Configure the sparkline at G2 to show high and low markers"
- "Delete the sparkline at H5"

---

## Tables

Excel Tables provide structured data management with automatic formatting, filtering, and formula references.

### Supported Operations
- **createTable**: Convert a range to a table with optional headers
- **styleTable**: Apply predefined table styles (Light, Medium, Dark)
- **addTableRow**: Add rows at end or specific position
- **addTableColumn**: Add columns with optional data
- **resizeTable**: Expand or shrink table range
- **convertToRange**: Convert table back to regular range
- **toggleTableTotals**: Enable/disable totals row with aggregation functions

### Table Styles
- **Light**: TableStyleLight1-21 (subtle formatting)
- **Medium**: TableStyleMedium1-28 (balanced formatting)
- **Dark**: TableStyleDark1-11 (bold formatting)

### Best Practices
1. Use descriptive table names (e.g., `SalesData`, `EmployeeList`)
2. Use structured references in formulas: `=SUM(SalesTable[Revenue])`
3. Enable auto-expand for dynamic data ranges
4. Avoid manual range references when working with tables

### Limitations
- Tables cannot overlap with other tables
- Table names limited to 31 characters
- Names cannot contain spaces or special characters
- Cannot create table on protected sheet

### Example Prompts
- "Create a table from A1:E100 with headers"
- "Style the table as TableStyleMedium2"
- "Add a totals row with Sum for the Sales column"
- "Add a new column called 'Status' to the table"
- "Convert Table1 back to a regular range"

---

## PivotTables

PivotTables enable dynamic data summarization, cross-tabulation, and aggregation.

### Supported Operations
- **createPivotTable**: Create from range or table
- **addPivotField**: Add fields to row/column/data/filter areas
- **configurePivotLayout**: Set layout type and display options
- **refreshPivotTable**: Update with latest source data
- **deletePivotTable**: Remove PivotTable

### Aggregation Functions
Sum, Count, Average, Max, Min, Product, CountNumbers, StandardDeviation, Variance

### Layout Types
- **Compact**: Default, indented layout (best for most cases)
- **Outline**: Expanded with subtotals below
- **Tabular**: Traditional spreadsheet format

### Best Practices
1. Use tables as source data (auto-updates when data changes)
2. Refresh PivotTable after source data changes
3. Limit to 5-10 fields for readability
4. Use slicers for interactive filtering

### Limitations
- Cannot create from non-contiguous ranges
- Field names must match source column headers exactly
- Large source data (100K+ rows) may impact performance

### Example Prompts
- "Create a pivot table showing sales by region and product"
- "Add Sum of Revenue to the data area"
- "Change the pivot layout to tabular form"
- "Add Year to the filter area"
- "Refresh all PivotTables in the workbook"

---

## Slicers

Slicers provide visual filtering for Tables and PivotTables.

### Supported Operations
- **createSlicer**: Create slicer for table column or pivot field
- **configureSlicer**: Change position, size, style, multi-select
- **connectSlicerToTable**: Connect to different table
- **connectSlicerToPivot**: Connect to different PivotTable
- **deleteSlicer**: Remove slicer

### Slicer Styles
- SlicerStyleLight1-6
- SlicerStyleMedium1-6
- SlicerStyleDark1-6

### Best Practices
1. Place slicers near the data they filter
2. Use consistent styles across related slicers
3. Enable multi-select for flexible filtering
4. Group related slicers together

### Limitations
- One slicer per field per table/PivotTable
- Cannot filter multiple unrelated tables simultaneously
- Requires ExcelApi 1.10+

### Example Prompts
- "Add a slicer for the Region column"
- "Create a slicer for the PivotTable Product field"
- "Change the slicer style to SlicerStyleLight3"
- "Enable multi-select on the Region slicer"

---

## Data Manipulation

Structural changes to worksheets including rows, columns, and cell operations.

### Supported Operations
- **insertRows**: Insert rows at specified position
- **insertColumns**: Insert columns at specified position
- **deleteRows**: Delete specified rows
- **deleteColumns**: Delete specified columns
- **mergeCells**: Merge range into single cell
- **unmergeCells**: Split merged cells
- **findReplace**: Find and replace text with options
- **textToColumns**: Split text by delimiter

### Find/Replace Options
- Case-sensitive matching
- Whole word matching
- Regular expression patterns
- Replace all or first occurrence

### Text to Columns Delimiters
- Comma, Tab, Semicolon, Space
- Custom delimiters
- Multiple delimiters simultaneously

### Best Practices
1. Backup data before destructive operations
2. Test find/replace on small range first
3. Avoid excessive cell merging (breaks formulas)
4. Check adjacent columns before text-to-columns

### Limitations
- Cannot undo after workbook save
- Formulas may show #REF! after row/column deletion
- Text to columns overwrites adjacent cells

### Example Prompts
- "Insert 5 rows at row 10"
- "Delete columns C through E"
- "Merge cells A1 to C1"
- "Find 'old' and replace with 'new' in the entire sheet"
- "Split column A by comma delimiter"

---

## Named Ranges

Named ranges provide descriptive names for cells and ranges, improving formula readability.

### Supported Operations
- **createNamedRange**: Create named range or constant
- **updateNamedRange**: Update range reference or value
- **deleteNamedRange**: Remove named range
- **listNamedRanges**: List all named ranges

### Scope Options
- **Workbook**: Available in all sheets
- **Worksheet**: Available only in specific sheet

### Best Practices
1. Use descriptive names (SalesData, TaxRate, InputRange)
2. Avoid names that look like cell references (A1, XFD1)
3. Use workbook scope for global constants
4. Document named ranges with comments

### Limitations
- 255-character name limit
- Cannot start with a number
- No spaces or special characters (except underscore)
- Cannot duplicate cell reference names

### Example Prompts
- "Create a named range 'SalesData' for A1:E100"
- "Create a named constant 'TaxRate' with value 0.08"
- "Update the TaxRate to 0.09"
- "List all named ranges in the workbook"
- "Delete the named range 'OldData'"

---

## Protection

Secure worksheets, ranges, and workbooks from unauthorized changes.

### Supported Operations
- **protectWorksheet**: Protect sheet with optional password and permissions
- **unprotectWorksheet**: Remove sheet protection
- **protectRange**: Lock specific cells (requires sheet protection)
- **unprotectRange**: Unlock specific cells
- **protectWorkbook**: Protect workbook structure
- **unprotectWorkbook**: Remove workbook protection

### Worksheet Protection Options
- Allow formatting cells/rows/columns
- Allow inserting/deleting rows/columns
- Allow sorting and filtering
- Allow using PivotTables
- Allow editing objects

### Best Practices
1. Use passwords for sensitive data
2. Allow specific operations (filtering, sorting) for usability
3. Lock formula cells, unlock input cells
4. Document passwords securely (cannot be recovered)

### Limitations
- Passwords provide basic protection only (not encryption)
- Cannot protect individual cells without sheet protection
- Forgotten passwords cannot be recovered

### Example Prompts
- "Protect the worksheet with password 'secret123'"
- "Allow sorting and filtering on the protected sheet"
- "Lock cells A1:A10 to prevent editing"
- "Protect the workbook structure"

---

## Shapes & Images

Visual elements for annotations, diagrams, and branding.

### Supported Operations
- **insertShape**: Add geometric shapes (rectangle, oval, arrow, etc.)
- **insertImage**: Add Base64-encoded images
- **insertTextBox**: Add text boxes with formatting
- **formatShape**: Change fill, line, text properties
- **deleteShape**: Remove shape
- **groupShapes**: Group multiple shapes
- **arrangeShapes**: Change z-order (front, back)
- **ungroupShapes**: Ungroup grouped shapes

### Shape Types
Rectangle, RoundRectangle, Oval, Diamond, Triangle, Pentagon, Hexagon, Star4, Star5, Arrow, Chevron, and 20+ more

### Best Practices
1. Use cell-based positioning for auto-move with data
2. Group related shapes for easier management
3. Use SVG format for scalable images
4. Keep image sizes under 5MB

### Limitations
- Images must be Base64 encoded (no URL support)
- ~5MB size limit for images
- Z-order is relative (no absolute positioning)

### Example Prompts
- "Insert a rectangle at cell D5"
- "Add a text box with 'Instructions: Fill in data below'"
- "Change the shape fill color to blue"
- "Group shapes Shape1 and Shape2"
- "Bring the chart to the front"

---

## Comments & Notes

Collaboration tools for team communication and personal annotations.

### Supported Operations
- **addComment**: Add threaded comment (modern)
- **addNote**: Add legacy note (yellow sticky)
- **editComment**: Modify comment content
- **editNote**: Modify note content
- **deleteComment**: Remove comment thread
- **deleteNote**: Remove note
- **replyToComment**: Add reply to comment thread
- **resolveComment**: Mark comment as resolved

### Comments vs Notes
- **Comments**: Modern, threaded, support @mentions, require ExcelApi 1.11+
- **Notes**: Legacy, single text block, available in all versions

### Best Practices
1. Use comments for team collaboration
2. Use notes for personal reminders
3. @mention team members for notifications
4. Resolve comments when addressed

### Limitations
- Comments require ExcelApi 1.11+ (Excel 365, Excel Online)
- Notes are legacy and don't support threading
- Very long content (>1000 chars) may be truncated

### Example Prompts
- "Add a comment 'Review this formula' to cell A1"
- "Reply to the comment in B5 with 'Fixed'"
- "@mention John in the comment"
- "Resolve the comment in C10"
- "Add a note to D5 explaining the calculation"

---

## Worksheet Management

Organize and navigate workbooks efficiently.

### Supported Operations
- **renameSheet**: Change worksheet name
- **moveSheet**: Reposition worksheet
- **hideSheet**: Hide worksheet from view
- **unhideSheet**: Show hidden worksheet
- **freezePanes**: Freeze rows/columns for scrolling
- **unfreezePane**: Remove freeze panes
- **setZoom**: Set zoom level (10-400%)
- **splitPane**: Split view at cell reference
- **createView**: Create custom view (limited support)

### Best Practices
1. Use descriptive sheet names (31-char limit)
2. Freeze header rows for large datasets
3. Use split panes to compare distant sections
4. Hide scratch/calculation sheets

### Limitations
- Cannot hide the last visible sheet
- Sheet names limited to 31 characters
- Custom views have limited API support (ExcelApi 1.14+)

### Example Prompts
- "Rename Sheet1 to 'Sales Data'"
- "Freeze the top row"
- "Freeze the first 2 columns"
- "Set zoom to 85%"
- "Split panes at cell B3"
- "Hide the 'Calculations' sheet"

---

## Page Setup & Printing

Configure printing and PDF export settings.

### Supported Operations
- **setPageSetup**: Set orientation, paper size, scaling
- **setPageMargins**: Set margin sizes
- **setPageOrientation**: Portrait or Landscape
- **setPrintArea**: Define print range
- **setHeaderFooter**: Add headers and footers
- **setPageBreaks**: Insert manual page breaks

### Header/Footer Codes
- `&[Page]` - Current page number
- `&[Pages]` - Total pages
- `&[Date]` - Current date
- `&[Time]` - Current time
- `&[File]` - Filename
- `&[Tab]` - Sheet name
- `&[Path]` - File path

### Best Practices
1. Use "Fit to 1 page wide" for reports
2. Set print area to exclude scratch data
3. Add page numbers to multi-page documents
4. Use landscape for wide tables

### Limitations
- Header/footer formatting is limited
- Page breaks are manual only
- Some settings may not apply to PDF export

### Example Prompts
- "Set the page orientation to landscape"
- "Set all margins to 1 inch"
- "Add page numbers to the footer"
- "Set the print area to A1:F50"
- "Insert a page break at row 20"

---

## Hyperlinks

Link to websites, emails, or internal document locations.

### Supported Operations
- **addHyperlink**: Add link to cell
- **editHyperlink**: Modify existing link
- **removeHyperlink**: Remove link (preserve cell value)

### Link Types
- Web URL: `https://example.com`
- Email: `mailto:user@example.com`
- Internal: `'Sheet2'!A1`

### Best Practices
1. Use descriptive display text
2. Add tooltips for context
3. Test links before sharing
4. Use internal links for navigation

### Limitations
- 255-character URL limit
- Cannot link to external files (security)
- Display text separate from URL

### Example Prompts
- "Add a hyperlink to https://example.com in cell A1"
- "Link B2 to Sheet2 cell A1"
- "Create an email link to support@company.com"
- "Change the link text to 'Click here'"
- "Remove the hyperlink from C3"

---

## Data Types

Structured entity cards with properties for rich data representation.

### Supported Operations
- **insertDataType**: Create custom entity with properties
- **refreshDataType**: Update entity properties

### Custom Entity Properties
- String, Double, Boolean values
- 1-10 properties per entity
- Descriptive text display
- BasicValue fallback for calculations

### Requirements
- ExcelApi 1.16+ (Excel 365, Excel 2021, Excel Online)
- Stocks/Geography require manual UI conversion (API limitation)

### Best Practices
1. Use for structured data (products, employees, locations)
2. Limit to 5-10 properties per entity
3. Set descriptive text for display
4. Use basicValue for calculations

### Limitations
- Requires ExcelApi 1.16+
- Stocks/Geography data types cannot be created via API
- Single cell per entity
- Performance impact with many entities

### Documentation
- See [Data Types User Guide](docs/data-types-user-guide.md)
- See [Data Types Limitations](docs/data-types-limitations.md)

### Example Prompts
- "Create a product entity with SKU, Price, and InStock properties"
- "Insert an employee record with ID, Name, and Department"
- "Update the price in the entity at A2 to 34.99"

---

## Advanced Formatting

Conditional formatting and cell styles for data visualization.

### Supported Operations
- **conditionalFormat**: Apply conditional formatting rules
- **clearFormat**: Remove formatting from range

### Conditional Format Types
1. **cellValue**: Highlight based on value comparison
2. **colorScale**: 2-color or 3-color gradient
3. **dataBar**: Progress bar visualization
4. **iconSet**: Traffic lights, arrows, ratings
5. **topBottom**: Top/bottom N items or percent
6. **preset**: Duplicates, blanks, errors, dates
7. **textComparison**: Contains, begins with, ends with
8. **custom**: Formula-based rules

### Best Practices
1. Use color scales for heatmaps
2. Use data bars for progress/magnitude
3. Use icon sets for ratings/status
4. Limit to 2-3 rules per range to avoid confusion

### Limitations
- Custom formulas can be slow on large ranges
- Overlapping rules may conflict (check priority)
- Some formats require specific Excel versions

### Example Prompts
- "Highlight cells greater than 100 in red"
- "Apply a color scale from red to green"
- "Add data bars to show progress"
- "Use traffic light icons for status"
- "Highlight duplicate values"
- "Format cells containing 'Error' in yellow"

---

## Feature Reference

For a complete list of all 87 supported operations with Excel version requirements, see:
- [Feature Matrix](docs/feature-matrix.md) - Complete operation reference
- [Example Prompts](docs/example-prompts.md) - 500+ prompt examples by category
- [Testing Guide](docs/testing-guide.md) - Developer testing documentation

---

## Troubleshooting

### Add-in doesn't load
- Ensure you're running `npm run start`
- Check that port 3000 is available
- Try clearing the Office cache

### API errors
- Verify your Gemini API key is correct
- Check your API quota at Google AI Studio
- Ensure you have internet connectivity

### Changes not applying
- Make sure you have a cell/range selected
- Check the status bar for error messages
- Open the Diagnostics panel (document icon) to see detailed logs

### Debugging Issues
1. Click the Diagnostics button (document icon) in the header
2. Enable Debug Mode in Settings for more verbose logging
3. Check the logs for specific error messages and skipped actions

## Development

### Build for production
```bash
npm run build
```

### Validate manifest
```bash
npm run validate
```

## Deployment

To deploy for organization-wide use:
1. Build the production version
2. Host the files on a web server with HTTPS
3. Update manifest.xml URLs to point to your server
4. Deploy via Microsoft 365 Admin Center or SharePoint App Catalog
