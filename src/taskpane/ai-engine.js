/*
 * AI Engine - Advanced AI capabilities for Excel Copilot
 * Features: Task-specific prompts, function calling, RAG, multi-step reasoning, learning
 */

/* global localStorage */

// ============================================================================
// Configuration
// ============================================================================
const AI_CONFIG = {
    CORRECTIONS_KEY: "excel_copilot_corrections",
    PATTERNS_KEY: "excel_copilot_patterns",
    MAX_CORRECTIONS: 50,
    MAX_PATTERNS: 100
};

// ============================================================================
// Task Type Detection
// ============================================================================
const TASK_TYPES = {
    FORMULA: "formula",
    CHART: "chart",
    ANALYSIS: "analysis",
    FORMAT: "format",
    DATA_ENTRY: "data_entry",
    VALIDATION: "validation",
    GENERAL: "general"
};

const TASK_KEYWORDS = {
    [TASK_TYPES.FORMULA]: [
        "formula", "sum", "average", "count", "vlookup", "xlookup", "if", "calculate",
        "total", "add up", "multiply", "divide", "percentage", "sumif", "countif",
        "index", "match", "concatenate", "lookup", "function",
        "clean", "trim", "upper", "lower", "remove spaces", "text manipulation",
        "convert", "proper case", "title case", "capitalize"
    ],
    [TASK_TYPES.CHART]: [
        "chart", "graph", "visualize", "plot", "pie", "bar", "line", "column",
        "histogram", "scatter", "trend", "visualization", "diagram"
    ],
    [TASK_TYPES.ANALYSIS]: [
        "analyze", "analysis", "insight", "summary", "summarize", "statistics",
        "trend", "pattern", "outlier", "anomaly", "compare", "correlation",
        "distribution", "breakdown", "report", "findings"
    ],
    [TASK_TYPES.FORMAT]: [
        "format", "style", "color", "bold", "italic", "font", "border",
        "highlight", "conditional", "table", "header", "align", "merge"
    ],
    [TASK_TYPES.DATA_ENTRY]: [
        "fill", "enter", "input", "write", "set", "update", "change value",
        "put", "insert", "add data", "populate"
    ],
    [TASK_TYPES.VALIDATION]: [
        "dropdown", "validation", "list", "restrict", "allow", "select from",
        "choices", "options", "pick list"
    ]
};

/**
 * Detects the task type from user prompt
 * @param {string} prompt - User's input
 * @returns {string} Task type
 */
function detectTaskType(prompt) {
    const lower = prompt.toLowerCase();
    
    for (const [taskType, keywords] of Object.entries(TASK_KEYWORDS)) {
        for (const keyword of keywords) {
            if (lower.includes(keyword)) {
                return taskType;
            }
        }
    }
    
    return TASK_TYPES.GENERAL;
}

// ============================================================================
// Task-Specific System Prompts
// ============================================================================
const TASK_PROMPTS = {
    [TASK_TYPES.FORMULA]: `You are an Excel Formula Expert. Your specialty is creating precise, efficient Excel formulas.

## YOUR EXPERTISE
- All Excel functions: SUM, AVERAGE, VLOOKUP, XLOOKUP, INDEX/MATCH, IF, SUMIF, COUNTIF, etc.
- Array formulas and dynamic arrays
- Nested functions and complex logic
- Error handling with IFERROR, IFNA
- Date/time calculations
- Text manipulation functions

## FORMULA BEST PRACTICES
1. Use XLOOKUP over VLOOKUP when possible (more flexible)
2. Prefer INDEX/MATCH for complex lookups
3. Always wrap lookups in IFERROR for robustness
4. Use structured references when working with tables
5. Consider performance for large datasets

## CRITICAL: DATA CLEANING IN-PLACE
**NEVER apply formulas to the same column they reference (causes circular reference)!**

When cleaning/converting data (TRIM, UPPER, LOWER, PROPER, etc.):
- **WRONG**: Apply =TRIM(C2) to C2 or =PROPER(D2) to D2 (circular reference!)
- **RIGHT**: Apply formula to a NEW column (like H2), then use copyValues to replace

Example for "Convert City to Proper Case":
Step 1 - Create formulas in helper column:
<ACTION type="formula" target="H2:H51">
=PROPER(D2)
</ACTION>

Step 2 - Copy values back to original column:
<ACTION type="copyValues" target="D2" source="H2:H51">
</ACTION>

**ALWAYS use both steps for data cleaning/conversion!**

## OUTPUT FORMAT
Always provide formulas in ACTION tags:
<ACTION type="formula" target="CELL">
=YOUR_FORMULA
</ACTION>

Explain what the formula does and why you chose this approach.`,

    [TASK_TYPES.CHART]: `You are an Excel Data Visualization Expert. Your specialty is creating effective charts.

## YOUR EXPERTISE
- Choosing the right chart type for the data
- Chart design and formatting
- Data storytelling through visuals
- Dashboard creation

## CHART SELECTION GUIDE
- **Column/Bar**: Comparing categories
- **Line**: Trends over time (use for trend analysis)
- **Pie/Doughnut**: Parts of a whole (use sparingly, max 5-7 slices)
- **Scatter**: Correlation between variables
- **Area**: Cumulative totals over time
- **Combo**: Multiple data types on one chart

## CRITICAL CHART RULES
1. **ALWAYS use CONTIGUOUS ranges** - e.g., A1:B10, NOT A1:A10,C1:C10
2. For trend analysis with non-adjacent columns, include ALL columns between them
3. If data columns are far apart, use the full data range (e.g., A1:G100)
4. Include headers in the first row for proper labels
5. For line/trend charts, ensure date/time is in the first column of the range

## OUTPUT FORMAT
<ACTION type="chart" target="DATARANGE" chartType="TYPE" title="TITLE" position="CELL">
</ACTION>

Example for trend: target="A1:G100" (full range), NOT "B1:B100,G1:G100"

Always explain why you chose this chart type and what story it tells.`,

    [TASK_TYPES.ANALYSIS]: `You are an Excel Data Analyst. Your specialty is extracting insights from data.

## YOUR EXPERTISE
- Statistical analysis (mean, median, mode, std dev)
- Trend identification
- Outlier detection
- Data quality assessment
- Pattern recognition
- Comparative analysis

## ANALYSIS APPROACH
1. First, understand the data structure
2. Identify key metrics and dimensions
3. Look for patterns, trends, anomalies
4. Provide actionable insights
5. Suggest visualizations if helpful

## OUTPUT FORMAT
Provide your analysis in clear sections:
- **Overview**: What the data represents
- **Key Findings**: Most important insights
- **Statistics**: Relevant numbers
- **Recommendations**: Suggested actions

If formulas or charts would help, include them in ACTION tags.`,

    [TASK_TYPES.FORMAT]: `You are an Excel Formatting Expert. Your specialty is making data visually clear and professional.

## YOUR EXPERTISE
- Professional table styling
- Conditional formatting rules
- Color schemes and accessibility
- Data presentation best practices

## FORMATTING BEST PRACTICES
1. Use consistent color schemes
2. Headers should stand out (bold, background color)
3. Align numbers right, text left
4. Use borders sparingly
5. Consider colorblind-friendly palettes

## OUTPUT FORMAT
<ACTION type="format" target="RANGE">
{"bold":true,"fill":"#4472C4","fontColor":"#FFFFFF"}
</ACTION>

Available format options: bold, italic, fill, fontColor, fontSize, numberFormat, border`,

    [TASK_TYPES.VALIDATION]: `You are an Excel Data Validation Expert. Your specialty is creating dropdowns and input controls.

## YOUR EXPERTISE
- Data validation rules
- Dropdown lists from ranges
- Input restrictions
- Error messages and input prompts

## VALIDATION BEST PRACTICES
1. Use source ranges for dynamic dropdowns
2. Provide clear error messages
3. Consider dependent dropdowns for hierarchical data

## OUTPUT FORMAT
<ACTION type="validation" target="CELL" source="DATARANGE">
</ACTION>

The source should be the range containing the list values.`,

    [TASK_TYPES.DATA_ENTRY]: `You are an Excel Data Entry Assistant. Your specialty is efficiently populating cells.

## YOUR EXPERTISE
- Bulk data entry
- Pattern-based filling
- Data transformation
- Autofill sequences

## REPLACING FORMULA RESULTS WITH VALUES
If user asks to "replace original with updated values" or "copy values back":
Use copyValues action (NOT values action):
<ACTION type="copyValues" target="A2" source="F2:F51">
</ACTION>

This copies only the calculated values (not formulas) from source to target.

## OUTPUT FORMAT FOR NEW DATA
<ACTION type="values" target="RANGE">
[["value1","value2"],["value3","value4"]]
</ACTION>

Values should be a 2D array matching the target range dimensions.`,

    [TASK_TYPES.GENERAL]: `You are Excel Copilot, a versatile Excel assistant.

## YOUR CAPABILITIES
- Create formulas and functions
- Build charts and visualizations
- Analyze data and provide insights
- Format cells and tables
- Set up data validation
- Automate repetitive tasks

Determine what the user needs and provide the most helpful response.`
};

/**
 * Gets the task-specific system prompt
 * @param {string} taskType - Detected task type
 * @param {Object} corrections - User corrections to incorporate
 * @returns {string} Complete system prompt
 */
function getTaskSpecificPrompt(taskType, corrections = {}) {
    const basePrompt = TASK_PROMPTS[taskType] || TASK_PROMPTS[TASK_TYPES.GENERAL];
    
    // Add corrections context if any
    let correctionsContext = "";
    if (Object.keys(corrections).length > 0) {
        correctionsContext = "\n\n## USER PREFERENCES (from past corrections)\n";
        for (const [key, value] of Object.entries(corrections)) {
            correctionsContext += `- ${key}: ${value}\n`;
        }
        correctionsContext += "\nAlways apply these preferences unless explicitly told otherwise.";
    }
    
    return basePrompt + correctionsContext + getCommonRules();
}

/**
 * Common rules appended to all prompts
 */
function getCommonRules() {
    return `

## CRITICAL RULES
1. **CHECK THE COLUMN STRUCTURE TABLE** - Find the exact column letter for each header name
2. **Data starts at row 2** (row 1 is headers)
3. Always verify column letters before creating formulas
4. Use the exact cell references from the data context

## ACTION TYPES REFERENCE
- formula: <ACTION type="formula" target="CELL">=FORMULA</ACTION>
- values: <ACTION type="values" target="RANGE">[["val"]]</ACTION>
- format: <ACTION type="format" target="RANGE">{"bold":true}</ACTION>
- conditionalFormat: <ACTION type="conditionalFormat" target="RANGE">{"type":"cellValue","operator":"GreaterThan","value":"40","fill":"#FFFF00"}</ACTION>
- chart: <ACTION type="chart" target="RANGE" chartType="TYPE" title="TITLE" position="CELL"></ACTION>
- validation: <ACTION type="validation" target="CELL" source="RANGE"></ACTION>
- sort: <ACTION type="sort" target="DATARANGE">{"column":1,"ascending":true}</ACTION>
- filter: <ACTION type="filter" target="DATARANGE">{"column":2,"values":["Mumbai"]}</ACTION>
- clearFilter: <ACTION type="clearFilter" target="DATARANGE"></ACTION>
- copy: <ACTION type="copy" target="DESTINATION" source="SOURCE"></ACTION>
- copyValues: <ACTION type="copyValues" target="DESTINATION" source="SOURCE"></ACTION>

## CONDITIONAL FORMATTING
**CRITICAL: For multiple conditions on the same range, use a SINGLE ACTION with an ARRAY of rules!**

Single condition:
<ACTION type="conditionalFormat" target="C2:C51">
{"type":"cellValue","operator":"GreaterThan","value":"40","fill":"#FFFF00"}
</ACTION>

**Multiple conditions (CORRECT WAY - use array):**
<ACTION type="conditionalFormat" target="E2:E51">
[
  {"type":"cellValue","operator":"GreaterThan","value":"70","fill":"#00FF00"},
  {"type":"cellValue","operator":"Between","value":"40","value2":"70","fill":"#FFFF00"},
  {"type":"cellValue","operator":"LessThan","value":"40","fill":"#FF0000"}
]
</ACTION>

To REMOVE/CLEAR conditional formatting:
<ACTION type="clearFormat" target="C2:C51">
</ACTION>

Operators: "GreaterThan", "LessThan", "EqualTo", "NotEqualTo", "GreaterThanOrEqual", "LessThanOrEqual", "Between"
Colors: Use hex codes like "#FFFF00" (yellow), "#FF0000" (red), "#00FF00" (green)
**Note:** For "Between" operator, use both "value" and "value2" properties

## SORTING DATA
To sort data, use the sort action type (NOT formulas like SORT()):
<ACTION type="sort" target="A1:L51">
{"column":1,"ascending":true}
</ACTION>

- target: The full data range including headers (e.g., A1:L51)
- column: 0-based index of the column to sort by (0=first column, 1=second, etc.)
- ascending: true for A-Z/smallest first, false for Z-A/largest first

## FILTERING DATA
To apply AutoFilter and filter by specific values:
<ACTION type="filter" target="A1:L51">
{"column":2,"values":["Mumbai","Delhi"]}
</ACTION>

- target: The full data range including headers (e.g., A1:L51)
- column: 0-based index of the column to filter by (0=first column, 1=second, etc.)
- values: Array of values to show (all other values will be hidden)

To REMOVE/CLEAR all filters and show all data:
<ACTION type="clearFilter" target="A1:L51">
</ACTION>

**Note:** Use clearFilter when user says "remove filter", "clear filter", "show all data", or "remove filtering".

## COPYING DATA
To copy formulas and formatting from one range to another:
<ACTION type="copy" target="A52" source="A1:L51">
</ACTION>

To copy ONLY VALUES (no formulas) - useful for replacing original data with cleaned values:
<ACTION type="copyValues" target="A2" source="F2:F51">
</ACTION>

- source: The range to copy FROM
- target: The starting cell to paste TO (top-left corner of destination)
- Use "copyValues" when replacing original data with formula results (e.g., after TRIM, UPPER, etc.)

## CREATING SHEETS
To create a new sheet:
<ACTION type="sheet" target="SheetName">
</ACTION>

- target: The name of the new sheet to create
- data: (optional) JSON array of values to populate the sheet

Example: Create a sheet named "Summary":
<ACTION type="sheet" target="Summary">
</ACTION>`;
}

// ============================================================================
// Function Calling - Direct Excel Operations
// ============================================================================
const EXCEL_FUNCTIONS = {
    // Aggregation functions
    SUM: {
        description: "Add numbers in a range",
        signature: "SUM(range)",
        example: "=SUM(A1:A10)"
    },
    AVERAGE: {
        description: "Calculate average of numbers",
        signature: "AVERAGE(range)",
        example: "=AVERAGE(B2:B100)"
    },
    COUNT: {
        description: "Count cells with numbers",
        signature: "COUNT(range)",
        example: "=COUNT(C1:C50)"
    },
    COUNTA: {
        description: "Count non-empty cells",
        signature: "COUNTA(range)",
        example: "=COUNTA(A:A)"
    },
    MAX: {
        description: "Find maximum value",
        signature: "MAX(range)",
        example: "=MAX(D2:D100)"
    },
    MIN: {
        description: "Find minimum value",
        signature: "MIN(range)",
        example: "=MIN(D2:D100)"
    },
    
    // Lookup functions
    VLOOKUP: {
        description: "Vertical lookup",
        signature: "VLOOKUP(lookup_value, table_array, col_index, [range_lookup])",
        example: "=VLOOKUP(A2, Sheet2!A:D, 3, FALSE)"
    },
    XLOOKUP: {
        description: "Modern flexible lookup",
        signature: "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])",
        example: "=XLOOKUP(A2, B:B, C:C, \"Not found\")"
    },
    INDEX: {
        description: "Return value at position",
        signature: "INDEX(array, row_num, [col_num])",
        example: "=INDEX(A1:C10, 5, 2)"
    },
    MATCH: {
        description: "Find position of value",
        signature: "MATCH(lookup_value, lookup_array, [match_type])",
        example: "=MATCH(\"Apple\", A:A, 0)"
    },
    
    // Conditional functions
    IF: {
        description: "Conditional logic",
        signature: "IF(condition, value_if_true, value_if_false)",
        example: "=IF(A1>100, \"High\", \"Low\")"
    },
    SUMIF: {
        description: "Sum with condition",
        signature: "SUMIF(range, criteria, [sum_range])",
        example: "=SUMIF(A:A, \"Sales\", B:B)"
    },
    COUNTIF: {
        description: "Count with condition",
        signature: "COUNTIF(range, criteria)",
        example: "=COUNTIF(A:A, \">100\")"
    },
    SUMIFS: {
        description: "Sum with multiple conditions",
        signature: "SUMIFS(sum_range, criteria_range1, criteria1, ...)",
        example: "=SUMIFS(C:C, A:A, \"Sales\", B:B, \">2023\")"
    },
    
    // Text functions
    CONCATENATE: {
        description: "Join text strings",
        signature: "CONCATENATE(text1, text2, ...)",
        example: "=CONCATENATE(A1, \" \", B1)"
    },
    LEFT: {
        description: "Extract left characters",
        signature: "LEFT(text, num_chars)",
        example: "=LEFT(A1, 3)"
    },
    RIGHT: {
        description: "Extract right characters",
        signature: "RIGHT(text, num_chars)",
        example: "=RIGHT(A1, 4)"
    },
    MID: {
        description: "Extract middle characters",
        signature: "MID(text, start_num, num_chars)",
        example: "=MID(A1, 2, 5)"
    },
    TRIM: {
        description: "Remove extra spaces",
        signature: "TRIM(text)",
        example: "=TRIM(A1)"
    },
    UPPER: {
        description: "Convert to uppercase",
        signature: "UPPER(text)",
        example: "=UPPER(A1)"
    },
    LOWER: {
        description: "Convert to lowercase",
        signature: "LOWER(text)",
        example: "=LOWER(A1)"
    },
    
    // Date functions
    TODAY: {
        description: "Current date",
        signature: "TODAY()",
        example: "=TODAY()"
    },
    NOW: {
        description: "Current date and time",
        signature: "NOW()",
        example: "=NOW()"
    },
    YEAR: {
        description: "Extract year from date",
        signature: "YEAR(date)",
        example: "=YEAR(A1)"
    },
    MONTH: {
        description: "Extract month from date",
        signature: "MONTH(date)",
        example: "=MONTH(A1)"
    },
    DAY: {
        description: "Extract day from date",
        signature: "DAY(date)",
        example: "=DAY(A1)"
    },
    DATEDIF: {
        description: "Difference between dates",
        signature: "DATEDIF(start_date, end_date, unit)",
        example: "=DATEDIF(A1, B1, \"D\")"
    },
    
    // Error handling
    IFERROR: {
        description: "Handle errors gracefully",
        signature: "IFERROR(value, value_if_error)",
        example: "=IFERROR(A1/B1, 0)"
    },
    IFNA: {
        description: "Handle #N/A errors",
        signature: "IFNA(value, value_if_na)",
        example: "=IFNA(VLOOKUP(...), \"Not found\")"
    }
};

/**
 * Generates function calling context for the AI
 * @returns {string} Function definitions for AI context
 */
function getFunctionCallingContext() {
    let context = "\n\n## AVAILABLE EXCEL FUNCTIONS\n";
    context += "You can use these functions directly. Here are the most common ones:\n\n";
    
    const categories = {
        "Aggregation": ["SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN"],
        "Lookup": ["VLOOKUP", "XLOOKUP", "INDEX", "MATCH"],
        "Conditional": ["IF", "SUMIF", "COUNTIF", "SUMIFS"],
        "Text": ["CONCATENATE", "LEFT", "RIGHT", "MID", "TRIM"],
        "Date": ["TODAY", "NOW", "YEAR", "MONTH", "DAY"],
        "Error Handling": ["IFERROR", "IFNA"]
    };
    
    for (const [category, funcs] of Object.entries(categories)) {
        context += `### ${category}\n`;
        for (const func of funcs) {
            const def = EXCEL_FUNCTIONS[func];
            if (def) {
                context += `- **${func}**: ${def.description}\n`;
                context += `  Syntax: \`${def.signature}\`\n`;
            }
        }
        context += "\n";
    }
    
    return context;
}

/**
 * Parses function calls from AI response and converts to actions
 * @param {string} response - AI response text
 * @returns {Object[]} Parsed function calls as actions
 */
function parseFunctionCalls(response) {
    const functionCalls = [];
    
    // Pattern: CALL_FUNCTION(name, target, args)
    const callPattern = /CALL_FUNCTION\s*\(\s*"?(\w+)"?\s*,\s*"?([A-Z]+\d+(?::[A-Z]+\d+)?)"?\s*(?:,\s*(.+?))?\s*\)/gi;
    
    let match;
    while ((match = callPattern.exec(response)) !== null) {
        const funcName = match[1].toUpperCase();
        const target = match[2];
        const args = match[3] ? match[3].trim() : "";
        
        if (EXCEL_FUNCTIONS[funcName]) {
            functionCalls.push({
                type: "formula",
                target: target,
                data: `=${funcName}(${args})`
            });
        }
    }
    
    return functionCalls;
}

// ============================================================================
// RAG - Pattern Knowledge Base
// ============================================================================
const FORMULA_PATTERNS = [
    // Aggregation patterns
    {
        id: "sum_column",
        keywords: ["sum", "total", "add up", "add all"],
        pattern: "=SUM({range})",
        description: "Sum all values in a column",
        example: "=SUM(B2:B100)"
    },
    {
        id: "average_column",
        keywords: ["average", "mean", "avg"],
        pattern: "=AVERAGE({range})",
        description: "Calculate average of values",
        example: "=AVERAGE(C2:C100)"
    },
    {
        id: "count_values",
        keywords: ["count", "how many", "number of"],
        pattern: "=COUNTA({range})",
        description: "Count non-empty cells",
        example: "=COUNTA(A2:A100)"
    },
    
    // Conditional patterns
    {
        id: "sumif_category",
        keywords: ["sum by", "total for", "sum where", "sum if"],
        pattern: "=SUMIF({criteria_range}, \"{criteria}\", {sum_range})",
        description: "Sum values matching a condition",
        example: "=SUMIF(A:A, \"Sales\", B:B)"
    },
    {
        id: "countif_condition",
        keywords: ["count where", "count if", "how many have"],
        pattern: "=COUNTIF({range}, \"{criteria}\")",
        description: "Count cells matching condition",
        example: "=COUNTIF(A:A, \"Complete\")"
    },
    {
        id: "if_condition",
        keywords: ["if then", "when", "check if", "condition"],
        pattern: "=IF({condition}, \"{true_value}\", \"{false_value}\")",
        description: "Return different values based on condition",
        example: "=IF(A1>100, \"High\", \"Low\")"
    },
    
    // Lookup patterns
    {
        id: "vlookup_basic",
        keywords: ["lookup", "find", "get value", "vlookup"],
        pattern: "=VLOOKUP({lookup_value}, {table}, {col_index}, FALSE)",
        description: "Look up a value in a table",
        example: "=VLOOKUP(A2, Products!A:C, 2, FALSE)"
    },
    {
        id: "xlookup_modern",
        keywords: ["xlookup", "modern lookup", "flexible lookup"],
        pattern: "=XLOOKUP({lookup_value}, {lookup_array}, {return_array}, \"Not found\")",
        description: "Modern flexible lookup",
        example: "=XLOOKUP(A2, B:B, C:C, \"Not found\")"
    },
    {
        id: "index_match",
        keywords: ["index match", "two-way lookup", "flexible lookup"],
        pattern: "=INDEX({return_range}, MATCH({lookup_value}, {lookup_range}, 0))",
        description: "Flexible lookup using INDEX/MATCH",
        example: "=INDEX(C:C, MATCH(A2, B:B, 0))"
    },
    
    // Percentage patterns
    {
        id: "percentage",
        keywords: ["percentage", "percent", "%", "ratio"],
        pattern: "={value}/{total}",
        description: "Calculate percentage",
        example: "=B2/SUM(B:B)"
    },
    {
        id: "percentage_change",
        keywords: ["change", "growth", "increase", "decrease"],
        pattern: "=({new_value}-{old_value})/{old_value}",
        description: "Calculate percentage change",
        example: "=(B2-A2)/A2"
    },
    
    // Text patterns
    {
        id: "concat_text",
        keywords: ["combine", "join", "concatenate", "merge text"],
        pattern: "=CONCATENATE({text1}, \" \", {text2})",
        description: "Join text values",
        example: "=CONCATENATE(A1, \" \", B1)"
    },
    {
        id: "extract_text",
        keywords: ["extract", "get part", "substring"],
        pattern: "=MID({text}, {start}, {length})",
        description: "Extract part of text",
        example: "=MID(A1, 1, 5)"
    },
    
    // Date patterns
    {
        id: "date_diff",
        keywords: ["days between", "date difference", "how long"],
        pattern: "=DATEDIF({start_date}, {end_date}, \"D\")",
        description: "Calculate days between dates",
        example: "=DATEDIF(A1, B1, \"D\")"
    },
    {
        id: "current_date",
        keywords: ["today", "current date", "now"],
        pattern: "=TODAY()",
        description: "Get current date",
        example: "=TODAY()"
    },
    
    // Error handling patterns
    {
        id: "safe_divide",
        keywords: ["divide", "safe division", "avoid error"],
        pattern: "=IFERROR({numerator}/{denominator}, 0)",
        description: "Safe division avoiding #DIV/0!",
        example: "=IFERROR(A1/B1, 0)"
    },
    {
        id: "safe_lookup",
        keywords: ["safe lookup", "handle not found"],
        pattern: "=IFERROR(VLOOKUP({value}, {range}, {col}, FALSE), \"Not found\")",
        description: "Lookup with error handling",
        example: "=IFERROR(VLOOKUP(A1, B:C, 2, FALSE), \"Not found\")"
    }
];

/**
 * Searches for relevant patterns based on user query
 * @param {string} query - User's request
 * @param {number} limit - Max patterns to return
 * @returns {Object[]} Matching patterns
 */
function searchPatterns(query, limit = 5) {
    const lower = query.toLowerCase();
    const scored = [];
    
    for (const pattern of FORMULA_PATTERNS) {
        let score = 0;
        
        // Check keyword matches
        for (const keyword of pattern.keywords) {
            if (lower.includes(keyword)) {
                score += 10;
            }
        }
        
        // Check description match
        if (pattern.description.toLowerCase().split(" ").some(w => lower.includes(w))) {
            score += 5;
        }
        
        if (score > 0) {
            scored.push({ ...pattern, score });
        }
    }
    
    // Sort by score and return top matches
    return scored
        .sort((a, b) => b.score - a.score)
        .slice(0, limit);
}

/**
 * Gets RAG context for the AI based on user query
 * @param {string} query - User's request
 * @returns {string} RAG context with relevant patterns
 */
function getRAGContext(query) {
    const patterns = searchPatterns(query);
    
    if (patterns.length === 0) {
        return "";
    }
    
    let context = "\n\n## RELEVANT FORMULA PATTERNS\n";
    context += "Based on your request, here are some useful patterns:\n\n";
    
    for (const pattern of patterns) {
        context += `### ${pattern.description}\n`;
        context += `Pattern: \`${pattern.pattern}\`\n`;
        context += `Example: \`${pattern.example}\`\n\n`;
    }
    
    return context;
}

/**
 * Adds a custom pattern to the knowledge base
 * @param {Object} pattern - Pattern to add
 */
function addCustomPattern(pattern) {
    const stored = JSON.parse(localStorage.getItem(AI_CONFIG.PATTERNS_KEY) || "[]");
    stored.push({
        ...pattern,
        id: `custom_${Date.now()}`,
        custom: true
    });
    
    // Keep only recent patterns
    if (stored.length > AI_CONFIG.MAX_PATTERNS) {
        stored.splice(0, stored.length - AI_CONFIG.MAX_PATTERNS);
    }
    
    localStorage.setItem(AI_CONFIG.PATTERNS_KEY, JSON.stringify(stored));
}

/**
 * Gets all patterns including custom ones
 * @returns {Object[]} All patterns
 */
function getAllPatterns() {
    const custom = JSON.parse(localStorage.getItem(AI_CONFIG.PATTERNS_KEY) || "[]");
    return [...FORMULA_PATTERNS, ...custom];
}

// ============================================================================
// Multi-Step Reasoning
// ============================================================================
const REASONING_STEPS = {
    ANALYZE: "analyze",
    PLAN: "plan",
    EXECUTE: "execute",
    VERIFY: "verify"
};

/**
 * Determines if a task requires multi-step reasoning
 * @param {string} prompt - User's request
 * @returns {boolean} True if complex task
 */
function requiresMultiStep(prompt) {
    const complexIndicators = [
        "and then", "after that", "multiple", "several", "all columns",
        "each row", "entire", "whole", "complex", "advanced",
        "step by step", "breakdown", "analyze and", "create and"
    ];
    
    const lower = prompt.toLowerCase();
    return complexIndicators.some(ind => lower.includes(ind)) || prompt.length > 200;
}

/**
 * Breaks down a complex task into steps
 * @param {string} prompt - User's request
 * @param {Object} dataContext - Current Excel data context
 * @returns {Object[]} Array of steps
 */
function decomposeTask(prompt, dataContext) {
    const steps = [];
    const lower = prompt.toLowerCase();
    
    // Step 1: Always analyze first
    steps.push({
        step: REASONING_STEPS.ANALYZE,
        description: "Understand the data structure and user intent",
        prompt: `First, analyze the data:
- What columns are available?
- What is the data type of each column?
- What is the user trying to achieve?`
    });
    
    // Step 2: Plan based on task type
    const taskType = detectTaskType(prompt);
    
    if (taskType === TASK_TYPES.FORMULA) {
        steps.push({
            step: REASONING_STEPS.PLAN,
            description: "Determine the formula approach",
            prompt: `Plan the formula:
- Which Excel function(s) are needed?
- What are the exact cell references?
- Are there any edge cases to handle?`
        });
    } else if (taskType === TASK_TYPES.CHART) {
        steps.push({
            step: REASONING_STEPS.PLAN,
            description: "Design the visualization",
            prompt: `Plan the chart:
- What chart type best represents this data?
- What should be on each axis?
- What title and labels are needed?`
        });
    } else if (taskType === TASK_TYPES.ANALYSIS) {
        steps.push({
            step: REASONING_STEPS.PLAN,
            description: "Plan the analysis approach",
            prompt: `Plan the analysis:
- What metrics should be calculated?
- What patterns should be looked for?
- What insights would be valuable?`
        });
    }
    
    // Step 3: Execute
    steps.push({
        step: REASONING_STEPS.EXECUTE,
        description: "Generate the solution",
        prompt: `Now execute the plan:
- Create the necessary ACTION tags
- Use exact cell references from the data
- Provide clear explanations`
    });
    
    // Step 4: Verify (for complex tasks)
    if (lower.includes("verify") || lower.includes("check") || prompt.length > 300) {
        steps.push({
            step: REASONING_STEPS.VERIFY,
            description: "Verify the solution",
            prompt: `Verify the solution:
- Are all cell references correct?
- Does the formula handle edge cases?
- Is the output format appropriate?`
        });
    }
    
    return steps;
}

/**
 * Generates a multi-step reasoning prompt
 * @param {string} userPrompt - Original user request
 * @param {Object} dataContext - Excel data context
 * @returns {string} Enhanced prompt with reasoning structure
 */
function generateReasoningPrompt(userPrompt, dataContext) {
    if (!requiresMultiStep(userPrompt)) {
        return userPrompt;
    }
    
    const steps = decomposeTask(userPrompt, dataContext);
    
    let enhancedPrompt = `## TASK DECOMPOSITION
This is a complex task. Please follow these steps:

`;
    
    for (let i = 0; i < steps.length; i++) {
        enhancedPrompt += `### Step ${i + 1}: ${steps[i].description}
${steps[i].prompt}

`;
    }
    
    enhancedPrompt += `## ORIGINAL REQUEST
${userPrompt}

Please work through each step and provide your final solution with ACTION tags.`;
    
    return enhancedPrompt;
}

// ============================================================================
// Learning from Corrections
// ============================================================================

/**
 * Stores a user correction for future reference
 * @param {string} original - What AI said/did
 * @param {string} correction - What user corrected to
 * @param {string} context - Additional context
 */
function storeCorrection(original, correction, context = "") {
    const corrections = getStoredCorrections();
    
    // Parse the correction to extract the key insight
    const insight = parseCorrection(original, correction);
    
    if (insight) {
        corrections.push({
            id: Date.now(),
            original,
            correction,
            insight,
            context,
            timestamp: new Date().toISOString()
        });
        
        // Keep only recent corrections
        if (corrections.length > AI_CONFIG.MAX_CORRECTIONS) {
            corrections.splice(0, corrections.length - AI_CONFIG.MAX_CORRECTIONS);
        }
        
        localStorage.setItem(AI_CONFIG.CORRECTIONS_KEY, JSON.stringify(corrections));
    }
}

/**
 * Gets stored corrections
 * @returns {Object[]} Array of corrections
 */
function getStoredCorrections() {
    try {
        return JSON.parse(localStorage.getItem(AI_CONFIG.CORRECTIONS_KEY) || "[]");
    } catch {
        return [];
    }
}

/**
 * Parses a correction to extract the key insight
 * @param {string} original - Original text
 * @param {string} correction - Correction text
 * @returns {Object|null} Parsed insight
 */
function parseCorrection(original, correction) {
    const lower = correction.toLowerCase();
    
    // Column corrections: "no, column E not C" or "use column E instead"
    const colMatch = lower.match(/column\s+([a-z])\s+(?:not|instead of|rather than)\s+([a-z])/i);
    if (colMatch) {
        return {
            type: "column_preference",
            wrong: colMatch[2].toUpperCase(),
            correct: colMatch[1].toUpperCase(),
            rule: `Use column ${colMatch[1].toUpperCase()} instead of ${colMatch[2].toUpperCase()}`
        };
    }
    
    // Cell reference corrections: "should be E2 not C2"
    const cellMatch = lower.match(/(?:should be|use)\s+([a-z]+\d+)\s+(?:not|instead of)\s+([a-z]+\d+)/i);
    if (cellMatch) {
        return {
            type: "cell_preference",
            wrong: cellMatch[2].toUpperCase(),
            correct: cellMatch[1].toUpperCase(),
            rule: `Use ${cellMatch[1].toUpperCase()} instead of ${cellMatch[2].toUpperCase()}`
        };
    }
    
    // Header name corrections: "the column is called X not Y"
    const headerMatch = lower.match(/(?:column is called|header is|named)\s+["']?([^"']+)["']?\s+(?:not|instead of)\s+["']?([^"']+)["']?/i);
    if (headerMatch) {
        return {
            type: "header_preference",
            wrong: headerMatch[2].trim(),
            correct: headerMatch[1].trim(),
            rule: `The column "${headerMatch[1].trim()}" should be used (not "${headerMatch[2].trim()}")`
        };
    }
    
    // Format preferences: "use currency format" or "format as percentage"
    const formatMatch = lower.match(/(?:use|format as|should be)\s+(currency|percentage|date|number|text)/i);
    if (formatMatch) {
        return {
            type: "format_preference",
            format: formatMatch[1].toLowerCase(),
            rule: `Format values as ${formatMatch[1].toLowerCase()}`
        };
    }
    
    // Chart preferences: "use bar chart not pie"
    const chartMatch = lower.match(/(?:use|prefer)\s+(bar|line|pie|column|scatter)\s+(?:chart|graph)/i);
    if (chartMatch) {
        return {
            type: "chart_preference",
            chartType: chartMatch[1].toLowerCase(),
            rule: `Prefer ${chartMatch[1].toLowerCase()} charts`
        };
    }
    
    // General preference
    return {
        type: "general",
        rule: correction
    };
}

/**
 * Detects if user message is a correction
 * @param {string} message - User's message
 * @returns {boolean} True if it's a correction
 */
function isCorrection(message) {
    const correctionIndicators = [
        "no,", "not", "wrong", "incorrect", "should be", "instead",
        "actually", "i meant", "that's not", "use column", "the column is"
    ];
    
    const lower = message.toLowerCase().trim();
    return correctionIndicators.some(ind => lower.startsWith(ind) || lower.includes(ind));
}

/**
 * Gets correction context for AI prompt
 * @returns {Object} Corrections organized by type
 */
function getCorrectionContext() {
    const corrections = getStoredCorrections();
    const context = {};
    
    // Group by type and get most recent
    for (const corr of corrections) {
        if (corr.insight) {
            const key = corr.insight.type;
            if (!context[key]) {
                context[key] = [];
            }
            context[key].push(corr.insight.rule);
        }
    }
    
    // Deduplicate and format
    const formatted = {};
    for (const [type, rules] of Object.entries(context)) {
        formatted[type] = [...new Set(rules)].slice(-5).join("; ");
    }
    
    return formatted;
}

/**
 * Clears all stored corrections
 */
function clearCorrections() {
    localStorage.removeItem(AI_CONFIG.CORRECTIONS_KEY);
}

// ============================================================================
// Main AI Engine Interface
// ============================================================================

/**
 * Enhances a user prompt with all AI features
 * @param {string} userPrompt - Original user prompt
 * @param {Object} dataContext - Excel data context
 * @returns {Object} Enhanced prompt and system prompt
 */
function enhancePrompt(userPrompt, dataContext) {
    // Detect task type
    const taskType = detectTaskType(userPrompt);
    
    // Check if this is a correction
    const isCorrectionMsg = isCorrection(userPrompt);
    
    // Get corrections context
    const corrections = getCorrectionContext();
    
    // Get task-specific system prompt
    let systemPrompt = getTaskSpecificPrompt(taskType, corrections);
    
    // Add function calling context for formula tasks
    if (taskType === TASK_TYPES.FORMULA) {
        systemPrompt += getFunctionCallingContext();
    }
    
    // Get RAG context
    const ragContext = getRAGContext(userPrompt);
    
    // Apply multi-step reasoning if needed
    let enhancedUserPrompt = userPrompt;
    if (!isCorrectionMsg && requiresMultiStep(userPrompt)) {
        enhancedUserPrompt = generateReasoningPrompt(userPrompt, dataContext);
    }
    
    // Add RAG context to user prompt
    if (ragContext) {
        enhancedUserPrompt = ragContext + "\n\n" + enhancedUserPrompt;
    }
    
    return {
        systemPrompt,
        userPrompt: enhancedUserPrompt,
        taskType,
        isCorrection: isCorrectionMsg,
        hasMultiStep: requiresMultiStep(userPrompt)
    };
}

/**
 * Processes AI response and extracts any function calls
 * @param {string} response - AI response
 * @returns {Object} Processed response with actions
 */
function processResponse(response) {
    // Parse any function calls
    const functionCalls = parseFunctionCalls(response);
    
    // Clean response (remove function call syntax)
    const cleanedResponse = response.replace(/CALL_FUNCTION\s*\([^)]+\)/gi, "");
    
    return {
        response: cleanedResponse,
        additionalActions: functionCalls
    };
}

/**
 * Handles a user correction
 * @param {string} userMessage - User's correction message
 * @param {string} previousAIResponse - What AI said before
 */
function handleCorrection(userMessage, previousAIResponse) {
    if (isCorrection(userMessage)) {
        storeCorrection(previousAIResponse, userMessage);
        return true;
    }
    return false;
}

// ============================================================================
// Exports
// ============================================================================
export {
    // Task detection
    detectTaskType,
    TASK_TYPES,
    
    // Prompts
    getTaskSpecificPrompt,
    enhancePrompt,
    
    // Function calling
    EXCEL_FUNCTIONS,
    getFunctionCallingContext,
    parseFunctionCalls,
    
    // RAG
    searchPatterns,
    getRAGContext,
    addCustomPattern,
    getAllPatterns,
    FORMULA_PATTERNS,
    
    // Multi-step reasoning
    requiresMultiStep,
    decomposeTask,
    generateReasoningPrompt,
    
    // Corrections
    isCorrection,
    storeCorrection,
    getStoredCorrections,
    getCorrectionContext,
    clearCorrections,
    handleCorrection,
    
    // Main interface
    processResponse
};
