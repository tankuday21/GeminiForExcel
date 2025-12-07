# Excel AI Copilot - Complete Documentation

**Version:** 3.6.2  
**Repository:** https://github.com/tankuday21/GeminiForExcel  
**License:** MIT

---

## Table of Contents

1. [Overview](#overview)
2. [What It Does](#what-it-does)
3. [Key Features](#key-features)
4. [How It Works](#how-it-works)
5. [Installation & Setup](#installation--setup)
6. [User Interface](#user-interface)
7. [Supported Operations](#supported-operations)
8. [AI Capabilities](#ai-capabilities)
9. [Usage Examples](#usage-examples)
10. [Technical Architecture](#technical-architecture)
11. [Security & Privacy](#security--privacy)
12. [Troubleshooting](#troubleshooting)
13. [Version Compatibility](#version-compatibility)
14. [Development](#development)

---

## Overview

**Excel AI Copilot** is an intelligent Office Add-in that brings AI-powered assistance directly into Microsoft Excel. It uses Google's Gemini AI to understand your data and natural language requests, then automatically performs Excel operations like creating formulas, generating charts, formatting data, and much more.

### What Makes It Special

- **Natural Language Interface**: Just describe what you want in plain English
- **Context-Aware**: Understands your data structure and suggests relevant actions
- **90 Supported Operations**: From basic formulas to advanced PivotTables and sparklines
- **Smart Suggestions**: Provides intelligent recommendations based on your data
- **Multi-Sheet Support**: Can work across multiple worksheets
- **Learning System**: Remembers your preferences and corrections
- **Preview Before Apply**: See exactly what will change before committing

---

## What It Does

### Core Capabilities

1. **Formula Creation**: Generate Excel formulas from natural language
   - "Sum column A" → Creates SUM formula
   - "Calculate average of sales" → Creates AVERAGE formula
   - "VLOOKUP to find price" → Creates VLOOKUP with proper syntax

2. **Data Analysis**: Extract insights and statistics
   - Identifies trends and patterns
   - Calculates key metrics
   - Detects outliers and anomalies
   - Provides actionable recommendations

3. **Chart Generation**: Create visualizations automatically
   - Chooses appropriate chart type for your data
   - Adds trendlines and data labels
   - Formats charts professionally
   - Supports 20+ chart types

4. **Data Formatting**: Apply professional styling
   - Conditional formatting with color scales, data bars, icon sets
   - Number formatting (currency, dates, percentages)
   - Cell styles and borders
   - Table creation and styling

5. **Data Manipulation**: Transform and clean data
   - Insert/delete rows and columns
   - Find and replace
   - Remove duplicates
   - Text to columns
   - Merge/unmerge cells

6. **Advanced Features**:
   - PivotTables and slicers
   - Sparklines for trend visualization
   - Data validation and dropdowns
   - Named ranges
   - Worksheet protection
   - Comments and annotations
   - Page setup for printing

---

## Key Features

### 🤖 AI-Powered Intelligence

- **Task Detection**: Automatically identifies what you're trying to do
- **Context Understanding**: Analyzes your data structure and content
- **Smart Prompts**: Task-specific system prompts for better accuracy
- **Function Calling**: Structured output for reliable Excel operations
- **RAG (Retrieval Augmented Generation)**: Uses documentation for accurate API usage
- **Learning System**: Remembers corrections and preferences

### 💡 Smart Suggestions

- Analyzes your data and suggests relevant actions
- Detects numeric columns → suggests totals and charts
- Finds date columns → suggests sorting by date
- Identifies email columns → suggests validation
- Adapts suggestions based on data characteristics

### 🎯 Preview System

- See all proposed changes before applying
- Select/deselect individual actions
- Expand actions to see details
- Highlight actions to preview location
- Undo support for applied changes

### 📊 Comprehensive Operations

**90 supported operations** across 16 categories:
- Basic Operations (6)
- Advanced Formatting (2)
- Charts (2)
- Copy/Filter/Duplicates (5)
- Sheet Management (1)
- Tables (7)
- PivotTables (5)
- Slicers (5)
- Data Manipulation (8)
- Named Ranges (4)
- Protection (6)
- Shapes & Images (9)
- Comments (8)
- Sparklines (3)
- Worksheet Management (9)
- Page Setup (6)
- Hyperlinks (3)
- Data Types (2)

### 🔄 Multi-Sheet Support

- Work with single sheet or entire workbook
- Toggle between modes in settings
- Automatically reads all sheets when enabled
- Maintains context across sheets

### 🎨 Modern UI

- Clean, intuitive interface
- Dark/light theme support
- Responsive design
- Smooth animations
- Keyboard shortcuts
- Diagnostics panel for debugging

---

## How It Works

### Step-by-Step Process

1. **User Input**: You type a natural language request
   - Example: "Create a chart showing sales by region"

2. **Data Context**: The add-in reads your Excel data
   - Captures selected range or entire sheet
   - Identifies headers and data types
   - Builds structured context

3. **AI Processing**: Request sent to Gemini AI
   - Task type detected (chart, formula, analysis, etc.)
   - Task-specific system prompt applied
   - AI generates structured actions in XML format

4. **Action Parsing**: Response converted to executable actions
   - XML parsed into action objects
   - Validation and error checking
   - Preview generated

5. **User Review**: You see proposed changes
   - Preview panel shows all actions
   - Select which actions to apply
   - Expand to see details

6. **Execution**: Selected actions applied to Excel
   - Uses Office.js API
   - Batch operations for performance
   - Error handling and logging
   - History saved for undo

### AI Engine Architecture

```
User Prompt
    ↓
Task Type Detection (formula, chart, analysis, etc.)
    ↓
Task-Specific System Prompt
    ↓
Data Context Building (headers, values, structure)
    ↓
RAG Context (relevant documentation)
    ↓
Gemini AI Processing
    ↓
Structured XML Response
    ↓
Action Parsing & Validation
    ↓
Preview Generation
    ↓
User Approval
    ↓
Excel API Execution
```

### Supported AI Models

- **Gemini 2.5 Flash** (Default): Fast, efficient, best for most tasks
- **Gemini 2.0 Flash Exp**: Experimental features
- **Gemini 2.0 Pro Exp**: Advanced reasoning for complex tasks

---

## Installation & Setup

### Option A: Development Mode (localhost)

**For testing and development on your machine**

1. **Get Gemini API Key**
   - Visit [Google AI Studio](https://aistudio.google.com/apikey)
   - Create new API key
   - Copy the key

2. **Install Dependencies**
   ```bash
   cd GeminiForExcel
   npm install
   ```

3. **Start Development Server**
   ```bash
   npm run start
   ```

4. **Load in Excel**
   - Excel will open automatically
   - Add-in appears in Home tab
   - Click "AI Copilot" to open

5. **Configure API Key**
   - Click Settings (gear icon)
   - Paste your API key
   - Click Save

### Option B: Production Deployment (GitHub Pages)

**Deploy once, use anywhere - like built-in Copilot**

1. **Create GitHub Repository**
   - Go to github.com/new
   - Name: `excel-ai-copilot`
   - Make it Public (required for free GitHub Pages)

2. **Build Production Files**
   ```bash
   npm run build:prod
   ```

3. **Push to GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/YOUR_USERNAME/excel-ai-copilot.git
   git branch -M main
   git push -u origin main
   ```

4. **Enable GitHub Pages**
   - Repo → Settings → Pages
   - Source: Deploy from branch
   - Branch: main → /docs folder
   - Save and wait 2-3 minutes

5. **Update Manifest**
   - Open `manifest.prod.xml`
   - Replace `YOUR_GITHUB_USERNAME` with your username
   - Replace `YOUR_REPO_NAME` with `excel-ai-copilot`

6. **Install Add-in**
   - Excel → Insert → Get Add-ins → My Add-ins
   - Upload My Add-in
   - Select `manifest.prod.xml`

Now works in ANY workbook, on ANY device!

---

## User Interface

### Main Components

#### 1. Header Bar
- **Version Badge**: Shows current version, click to check for updates
- **Refresh Button**: Reload Excel data
- **Settings Button**: Configure API key, model, preferences
- **Theme Toggle**: Switch between light/dark mode
- **Diagnostics Button**: View logs and debug information

#### 2. Welcome Screen (First Use)
- Quick action buttons for common tasks
- Smart suggestions based on your data
- Example prompts to get started

#### 3. Chat Interface
- **Message History**: Conversation with AI
- **User Messages**: Your requests
- **AI Responses**: Explanations and suggestions
- **Action Previews**: Proposed changes

#### 4. Input Area
- **Text Input**: Type your request
- **Send Button**: Submit request
- **Quick Actions**: Pre-defined prompts

#### 5. Preview Panel
- **Action List**: All proposed changes
- **Checkboxes**: Select actions to apply
- **Expand/Collapse**: View action details
- **Highlight**: Preview location in Excel

#### 6. Action Buttons
- **Apply Changes**: Execute selected actions
- **Clear Chat**: Start fresh conversation
- **History**: View and undo previous actions

#### 7. Settings Modal
- **API Key**: Enter/remove Gemini API key
- **Model Selection**: Choose AI model
- **Worksheet Scope**: Single sheet or all sheets
- **Debug Mode**: Enable verbose logging
- **Update Check**: Check for new versions

### Keyboard Shortcuts

- **Enter**: Send message (Shift+Enter for new line)
- **Ctrl+K**: Focus input
- **Ctrl+L**: Clear chat
- **Ctrl+Z**: Undo last action
- **Ctrl+H**: Toggle history panel
- **Ctrl+D**: Toggle diagnostics

---

## Supported Operations

### Complete Operation List (90 Actions)

#### Basic Operations (6)
1. `formula` - Apply formulas to cells
2. `values` - Insert data values
3. `format` - Cell formatting (font, fill, borders)
4. `validation` - Data validation rules
5. `sort` - Sort data by columns
6. `autofill` - Auto-fill patterns

#### Advanced Formatting (2)
7. `conditionalFormat` - Conditional formatting rules
8. `clearFormat` - Remove formatting

#### Charts (2)
9. `chart` - Create charts (20+ types)
10. `pivotChart` - Create PivotChart

#### Copy/Filter/Duplicates (5)
11. `copy` - Copy range with formatting
12. `copyValues` - Copy values only
13. `filter` - Apply AutoFilter
14. `clearFilter` - Remove filters
15. `removeDuplicates` - Remove duplicate rows

#### Tables (7)
16. `createTable` - Convert range to table
17. `styleTable` - Apply table style
18. `addTableRow` - Add rows to table
19. `addTableColumn` - Add columns to table
20. `resizeTable` - Change table range
21. `convertToRange` - Convert table to range
22. `toggleTableTotals` - Show/hide totals row

#### PivotTables (5)
23. `createPivotTable` - Create PivotTable
24. `addPivotField` - Add field to pivot
25. `configurePivotLayout` - Set layout options
26. `refreshPivotTable` - Refresh pivot data
27. `deletePivotTable` - Remove PivotTable

#### Slicers (5)
28. `createSlicer` - Create slicer
29. `configureSlicer` - Change slicer settings
30. `connectSlicerToTable` - Connect to table
31. `connectSlicerToPivot` - Connect to pivot
32. `deleteSlicer` - Remove slicer

#### Data Manipulation (8)
33. `insertRows` - Insert rows
34. `insertColumns` - Insert columns
35. `deleteRows` - Delete rows
36. `deleteColumns` - Delete columns
37. `mergeCells` - Merge cells
38. `unmergeCells` - Unmerge cells
39. `findReplace` - Find and replace
40. `textToColumns` - Split text by delimiter

#### Named Ranges (4)
41. `createNamedRange` - Create named range
42. `updateNamedRange` - Update range reference
43. `deleteNamedRange` - Remove named range
44. `listNamedRanges` - List all named ranges

#### Protection (6)
45. `protectWorksheet` - Protect sheet
46. `unprotectWorksheet` - Unprotect sheet
47. `protectRange` - Lock cells
48. `unprotectRange` - Unlock cells
49. `protectWorkbook` - Protect workbook
50. `unprotectWorkbook` - Unprotect workbook

#### Shapes & Images (8)
51. `insertShape` - Add geometric shape
52. `insertImage` - Add image
53. `insertTextBox` - Add text box
54. `formatShape` - Change shape properties
55. `deleteShape` - Remove shape
56. `groupShapes` - Group shapes
57. `arrangeShapes` - Change z-order
58. `ungroupShapes` - Ungroup shapes

#### Comments (8)
59. `addComment` - Add threaded comment
60. `addNote` - Add legacy note
61. `editComment` - Modify comment
62. `editNote` - Modify note
63. `deleteComment` - Remove comment
64. `deleteNote` - Remove note
65. `replyToComment` - Reply to comment
66. `resolveComment` - Mark as resolved

#### Sparklines (3)
67. `createSparkline` - Create in-cell chart
68. `configureSparkline` - Change sparkline settings
69. `deleteSparkline` - Remove sparkline

#### Worksheet Management (9)
70. `renameSheet` - Change sheet name
71. `moveSheet` - Reposition sheet
72. `hideSheet` - Hide sheet
73. `unhideSheet` - Show sheet
74. `freezePanes` - Freeze rows/columns
75. `unfreezePane` - Remove freeze
76. `setZoom` - Set zoom level
77. `splitPane` - Split view
78. `createView` - Create custom view

#### Page Setup (6)
79. `setPageSetup` - Set page options
80. `setPageMargins` - Set margins
81. `setPageOrientation` - Portrait/landscape
82. `setPrintArea` - Define print range
83. `setHeaderFooter` - Add headers/footers
84. `setPageBreaks` - Insert page breaks

#### Hyperlinks (3)
85. `addHyperlink` - Add link
86. `editHyperlink` - Modify link
87. `removeHyperlink` - Remove link

#### Sheet Management (1)
88. `sheet` - Create new worksheet

#### Data Types (2)
89. `insertDataType` - Insert custom entity with properties
90. `refreshDataType` - Update entity data

### Excel Version Requirements

| Feature | Excel 2016 | Excel 2019 | Excel 2021 | Excel 365 |
|---------|-----------|-----------|-----------|-----------|
| Basic Operations | ✅ | ✅ | ✅ | ✅ |
| Tables & Charts | ✅ | ✅ | ✅ | ✅ |
| PivotTables | ✅ | ✅ | ✅ | ✅ |
| Slicers | ❌ | ✅ | ✅ | ✅ |
| Sparklines | ❌ | ✅ | ✅ | ✅ |
| Modern Comments | ❌ | ❌ | ❌ | ✅ |
| Data Types | ❌ | ❌ | ✅ | ✅ |
| Dynamic Arrays | ❌ | ❌ | ✅ | ✅ |

---

## AI Capabilities

### Task Type Detection

The AI automatically detects what you're trying to do:

- **Formula**: Creating calculations
- **Chart**: Data visualization
- **Analysis**: Extracting insights
- **Format**: Styling and appearance
- **Data Entry**: Inserting values
- **Validation**: Dropdowns and rules
- **Table**: Table operations
- **Pivot**: PivotTable operations
- **Data Manipulation**: Structural changes
- **Shapes**: Visual elements
- **Comments**: Annotations
- **Protection**: Security
- **Page Setup**: Printing
- **Sparkline**: Inline charts
- **Worksheet Management**: Sheet organization
- **Data Types**: Entity cards
- **General**: Other requests

### Task-Specific Prompts

Each task type has specialized system prompts:

**Formula Expert**: Knows all Excel functions, best practices, structured references
**Chart Expert**: Chooses appropriate chart types, adds trendlines, formats professionally
**Analysis Expert**: Identifies patterns, calculates statistics, provides insights
**Format Expert**: Applies professional styling, conditional formatting, accessibility

### Learning System

The AI learns from your corrections:

1. **Correction Detection**: Recognizes when you correct its output
2. **Pattern Storage**: Saves correction patterns
3. **Future Application**: Applies learned preferences automatically
4. **Clear Preferences**: Option to reset learned corrections

### RAG (Retrieval Augmented Generation)

Uses documentation to improve accuracy:

- Embeds feature documentation
- Retrieves relevant context for each task
- Ensures correct API usage
- Reduces hallucinations

---

## Usage Examples

### Example 1: Create Formula

**User**: "Sum column B"

**AI Response**:
```
I'll create a SUM formula for column B.

<ACTION type="formula" target="B51">
=SUM(B2:B50)
</ACTION>

This formula adds all values in column B from row 2 to 50.
```

**Result**: Formula applied to B51 showing total

---

### Example 2: Generate Chart

**User**: "Create a chart showing sales by region"

**AI Response**:
```
I'll create a column chart to visualize sales by region.

<ACTION type="chart" target="A1:B6" chartType="columnClustered" title="Sales by Region" position="D2">
</ACTION>

This chart compares sales across different regions using a column chart,
which is ideal for category comparisons.
```

**Result**: Professional column chart created at D2

---

### Example 3: Data Analysis

**User**: "Analyze this sales data"

**AI Response**:
```
## Sales Data Analysis

**Overview**: 50 sales records across 5 regions

**Key Findings**:
- Total Sales: $125,450
- Average Sale: $2,509
- Top Region: West ($35,200)
- Growth Trend: +15% month-over-month

**Recommendations**:
1. Focus marketing on West region (highest performer)
2. Investigate East region (below average)
3. Consider seasonal promotions in Q4

Would you like me to create a chart or PivotTable to visualize these insights?
```

---

### Example 4: Conditional Formatting

**User**: "Highlight values above average in green"

**AI Response**:
```
I'll apply conditional formatting to highlight above-average values.

<ACTION type="conditionalFormat" target="B2:B50">
{"type":"cellValue","operator":"GreaterThan","formula":"=AVERAGE($B$2:$B$50)","format":{"fill":"#90EE90"}}
</ACTION>

Cells with values above the average will be highlighted in light green.
```

**Result**: Above-average cells highlighted

---

### Example 5: Create PivotTable

**User**: "Create a pivot table showing sales by region and product"

**AI Response**:
```
I'll create a PivotTable to summarize sales by region and product.

<ACTION type="createPivotTable" target="A1:D50" destination="F2">
{"name":"SalesPivot"}
</ACTION>

<ACTION type="addPivotField" target="SalesPivot">
{"field":"Region","area":"row"}
</ACTION>

<ACTION type="addPivotField" target="SalesPivot">
{"field":"Product","area":"column"}
</ACTION>

<ACTION type="addPivotField" target="SalesPivot">
{"field":"Sales","area":"data","function":"Sum"}
</ACTION>

This creates a cross-tabulation showing total sales for each region-product combination.
```

**Result**: Interactive PivotTable created

---

## Technical Architecture

### Technology Stack

- **Frontend**: HTML, CSS, JavaScript (ES6+)
- **Office Integration**: Office.js API
- **AI Engine**: Google Gemini AI
- **Build Tool**: Webpack 5
- **Testing**: Jest
- **Linting**: ESLint

### Project Structure

```
GeminiForExcel/
├── src/
│   └── taskpane/
│       ├── taskpane.js          # Main application logic
│       ├── taskpane.html        # UI structure
│       ├── taskpane.css         # Styling
│       ├── ai-engine.js         # AI processing
│       ├── action-executor.js   # Excel operations
│       ├── excel-data.js        # Data reading
│       └── diagnostics.js       # Logging system
├── assets/                      # Icons and images
├── docs/                        # Documentation
├── dist/                        # Production build
├── manifest.xml                 # Development manifest
├── manifest.prod.xml            # Production manifest
├── webpack.config.js            # Build configuration
└── package.json                 # Dependencies

```

### Key Modules

#### 1. taskpane.js (Main Application)
- UI initialization and event handling
- Chat interface management
- Settings and configuration
- History and undo functionality
- Theme management

#### 2. ai-engine.js (AI Processing)
- Task type detection
- Prompt engineering
- Gemini API communication
- Response parsing
- Learning system
- RAG context retrieval

#### 3. action-executor.js (Excel Operations)
- Action execution engine
- Office.js API calls
- Error handling
- Formula reference adjustment
- Batch operations

#### 4. excel-data.js (Data Reading)
- Data context building
- Header detection
- Selection monitoring
- Multi-sheet support
- Column mapping

#### 5. diagnostics.js (Logging)
- Debug logging
- Error tracking
- Performance monitoring
- Log export

### API Integration

#### Gemini AI API

**Endpoint**: `https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent`

**Request Format**:
```json
{
  "contents": [{
    "role": "user",
    "parts": [{"text": "prompt"}]
  }],
  "generationConfig": {
    "temperature": 0.2,
    "topP": 0.8,
    "topK": 40
  }
}
```

**Response Format**:
```json
{
  "candidates": [{
    "content": {
      "parts": [{"text": "response"}]
    }
  }]
}
```

#### Office.js API

**Initialization**:
```javascript
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initApp();
  }
});
```

**Excel Operations**:
```javascript
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:B10");
  range.values = [[1, 2], [3, 4]];
  await context.sync();
});
```

---

## Security & Privacy

### API Key Storage

**Current Implementation**:
- API keys stored in browser localStorage
- Basic obfuscation using Base64 encoding
- **NOT true encryption** - prevents casual viewing only

**Security Recommendations**:
1. Use "Remove API Key" button when not actively using
2. Consider re-entering key each session
3. Never share workbooks that might contain stored keys
4. For production, implement server-side key management

### Data Privacy

**What Gets Sent to AI**:
- Selected Excel data (values and headers)
- Your natural language prompt
- Conversation history (last 10 messages)
- Task context and metadata

**What Does NOT Get Sent**:
- Your API key (used for authentication only)
- Entire workbook (only selected/active sheet)
- Personal information (unless in your data)
- File names or paths

**Data Retention**:
- Google Gemini: Per Google's AI Studio terms
- Add-in: No server-side storage (runs client-side)
- Browser: localStorage for API key and preferences only

### Best Practices

1. **Review Before Applying**: Always preview AI suggestions
2. **Sensitive Data**: Avoid using with confidential information
3. **API Key Security**: Treat like a password
4. **Regular Updates**: Keep add-in updated for security patches
5. **Backup Data**: Save workbooks before major operations

---

## Troubleshooting

### Common Issues

#### Add-in Doesn't Load

**Symptoms**: Add-in button missing or not responding

**Solutions**:
1. Check if development server is running (`npm run start`)
2. Verify port 3000 is available
3. Clear Office cache:
   - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
   - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches`
4. Restart Excel
5. Re-sideload manifest

#### API Errors

**Symptoms**: "API key invalid" or "Request failed"

**Solutions**:
1. Verify API key is correct (check Google AI Studio)
2. Check API quota (free tier has limits)
3. Ensure internet connectivity
4. Try different AI model in settings
5. Check browser console for detailed errors

#### Changes Not Applying

**Symptoms**: Actions selected but nothing happens

**Solutions**:
1. Check if range is protected
2. Verify cell references are valid
3. Enable Debug Mode in settings
4. Check Diagnostics panel for errors
5. Try smaller operations first

#### Data Not Reading

**Symptoms**: "No data found" or incorrect context

**Solutions**:
1. Ensure sheet has data
2. Click Refresh button
3. Select specific range before asking
4. Check if sheet is hidden
5. Verify worksheet scope setting

#### Performance Issues

**Symptoms**: Slow responses or timeouts

**Solutions**:
1. Reduce data size (work with smaller ranges)
2. Use single sheet mode instead of all sheets
3. Clear conversation history
4. Close other Excel workbooks
5. Try faster AI model (Gemini 2.5 Flash)

### Debug Mode

Enable in Settings for verbose logging:

1. Click Settings (gear icon)
2. Check "Debug Mode"
3. Click Diagnostics (document icon)
4. View detailed logs
5. Export logs for support

### Getting Help

1. **Documentation**: Check SETUP.md and feature-matrix.md
2. **Examples**: See example-prompts.md for 500+ examples
3. **GitHub Issues**: Report bugs at repository
4. **Diagnostics**: Export logs and include in reports

---

## Version Compatibility

### Excel Versions

| Version | Support Level | Notes |
|---------|--------------|-------|
| Excel 2016 | Partial | Basic operations only |
| Excel 2019 | Full | All features except modern comments |
| Excel 2021 | Full | Includes data types |
| Excel 365 | Full | All features including dynamic arrays |
| Excel Online | Full | All features |

### Browser Support (Excel Online)

- Chrome 90+
- Edge 90+
- Firefox 88+
- Safari 14+

### Office.js API Versions

- **1.1+**: Basic operations (all versions)
- **1.7+**: Workbook protection, freeze panes
- **1.9+**: Shapes, page setup
- **1.10+**: Slicers, sparklines
- **1.11+**: Modern comments
- **1.14+**: Custom views
- **1.16+**: Data types

---

## Development

### Prerequisites

- Node.js 14+
- npm 6+
- Excel 2016+ or Excel Online
- Google Gemini API key

### Setup Development Environment

```bash
# Clone repository
git clone https://github.com/tankuday21/GeminiForExcel.git
cd GeminiForExcel

# Install dependencies
npm install

# Start development server
npm run start
```

### Build Commands

```bash
# Development build
npm run build:dev

# Production build
npm run build

# Production build with custom URL
npm run build:prod

# Watch mode
npm run watch
```

### Testing

```bash
# Run all tests
npm test

# Run unit tests only
npm run test:unit

# Run integration tests
npm run test:integration

# Run with coverage
npm run test:coverage

# Watch mode
npm run test:watch
```

### Linting

```bash
# Check for issues
npm run lint

# Auto-fix issues
npm run lint:fix

# Format code
npm run prettier
```

### Manifest Validation

```bash
# Validate manifest
npm run validate
```

### Project Scripts

| Script | Description |
|--------|-------------|
| `npm start` | Start dev server and sideload |
| `npm run build` | Production build |
| `npm run build:prod` | Build with GitHub Pages URL |
| `npm test` | Run tests |
| `npm run lint` | Check code quality |
| `npm run validate` | Validate manifest |

### Contributing

1. Fork the repository
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open Pull Request

### Code Style

- ES6+ JavaScript
- 2-space indentation
- Semicolons required
- Single quotes for strings
- JSDoc comments for functions
- Descriptive variable names

---

## Appendix

### Useful Links

- **Repository**: https://github.com/tankuday21/GeminiForExcel
- **Google AI Studio**: https://aistudio.google.com/apikey
- **Office.js Docs**: https://docs.microsoft.com/office/dev/add-ins/
- **Gemini API Docs**: https://ai.google.dev/docs

### Related Documentation

- [SETUP.md](SETUP.md) - Installation guide
- [feature-matrix.md](docs/feature-matrix.md) - Complete operation reference
- [example-prompts.md](docs/example-prompts.md) - 500+ prompt examples
- [data-types-user-guide.md](docs/data-types-user-guide.md) - Data types guide
- [testing-guide.md](docs/testing-guide.md) - Developer testing

### Version History

- **3.6.2** (Current): Enhanced diagnostics, improved error handling
- **3.6.0**: Added data types support
- **3.5.0**: Multi-sheet support, learning system
- **3.4.0**: Preview system, undo functionality
- **3.3.0**: Smart suggestions, task detection
- **3.2.0**: Sparklines, worksheet management
- **3.1.0**: PivotTables, slicers
- **3.0.0**: Complete rewrite with AI engine

### License

MIT License - See LICENSE file for details

### Credits

- **AI**: Google Gemini
- **Framework**: Office.js
- **Build**: Webpack
- **Testing**: Jest

---

**Last Updated**: December 2024  
**Maintained By**: tankuday21  
**Status**: Active Development

