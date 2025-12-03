/*
 * Excel AI Copilot - Accurate Data Understanding
 */

/* global document, Excel, Office, fetch, localStorage */

const CONFIG = {
    GEMINI_MODEL: "gemini-2.0-flash",
    API_ENDPOINT: "https://generativelanguage.googleapis.com/v1beta/models/",
    STORAGE_KEY: "excel_copilot_api_key",
    MAX_HISTORY: 10
};

const state = {
    apiKey: "",
    pendingActions: [],
    currentData: null,
    conversationHistory: [],
    isFirstMessage: true,
    // Preview state
    preview: {
        selections: [],      // boolean[] - selection state for each action
        expandedIndex: -1,   // number - index of expanded action (-1 if none)
        highlightedIndex: -1 // number - index of highlighted action (-1 if none)
    },
    // History state for undo functionality
    history: {
        entries: [],         // HistoryEntry[] - all history entries, newest first
        panelVisible: false, // boolean - whether history panel is shown
        maxEntries: 20       // number - maximum entries to retain
    }
};

// ============================================================================
// Initialize
// ============================================================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initApp();
    }
});

function initApp() {
    state.apiKey = localStorage.getItem(CONFIG.STORAGE_KEY) || "";
    bindEvents();
    readExcelData();
}

function bindEvents() {
    const sendBtn = document.getElementById("sendBtn");
    const input = document.getElementById("promptInput");
    
    sendBtn?.addEventListener("click", handleSend);
    input?.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            handleSend();
        }
    });
    input?.addEventListener("input", () => {
        sendBtn.disabled = !input.value.trim();
        input.style.height = "auto";
        input.style.height = Math.min(input.scrollHeight, 120) + "px";
    });
    
    document.getElementById("applyBtn")?.addEventListener("click", handleApply);
    
    document.getElementById("refreshBtn")?.addEventListener("click", async () => {
        const btn = document.getElementById("refreshBtn");
        btn.classList.add("loading");
        await readExcelData();
        btn.classList.remove("loading");
        toast("Refreshed");
    });
    
    document.getElementById("settingsBtn")?.addEventListener("click", () => {
        document.getElementById("apiKeyInput").value = state.apiKey;
        document.getElementById("modal").classList.add("open");
    });
    
    document.getElementById("closeModal")?.addEventListener("click", closeModal);
    document.getElementById("cancelBtn")?.addEventListener("click", closeModal);
    document.getElementById("saveBtn")?.addEventListener("click", () => {
        state.apiKey = document.getElementById("apiKeyInput").value.trim();
        localStorage.setItem(CONFIG.STORAGE_KEY, state.apiKey);
        closeModal();
        toast("Saved");
    });
    
    document.getElementById("modal")?.addEventListener("click", (e) => {
        if (e.target.id === "modal") closeModal();
    });
    
    document.getElementById("clearBtn")?.addEventListener("click", clearChat);
    
    // History and Undo buttons
    document.getElementById("historyBtn")?.addEventListener("click", toggleHistoryPanel);
    document.getElementById("undoBtn")?.addEventListener("click", performUndo);
    
    document.querySelectorAll("[data-prompt]").forEach(el => {
        el.addEventListener("click", () => {
            document.getElementById("promptInput").value = el.dataset.prompt;
            document.getElementById("sendBtn").disabled = false;
            handleSend();
        });
    });
    
    document.getElementById("togglePwd")?.addEventListener("click", () => {
        const inp = document.getElementById("apiKeyInput");
        inp.type = inp.type === "password" ? "text" : "password";
    });
    
    setupSelectionListener();
}

function closeModal() {
    document.getElementById("modal").classList.remove("open");
}

async function setupSelectionListener() {
    try {
        await Excel.run(async (ctx) => {
            ctx.workbook.worksheets.getActiveWorksheet().onSelectionChanged.add(readExcelData);
            await ctx.sync();
        });
    } catch (e) { /* ignore */ }
}

// ============================================================================
// Column Letter Helper
// ============================================================================
function colIndexToLetter(index) {
    let letter = "";
    while (index >= 0) {
        letter = String.fromCharCode((index % 26) + 65) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}

// ============================================================================
// Read Excel Data with Column Headers
// ============================================================================
async function readExcelData() {
    const infoEl = document.getElementById("contextInfo");
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            
            sheet.load("name");
            usedRange.load(["address", "values", "rowCount", "columnCount", "columnIndex", "rowIndex"]);
            
            await ctx.sync();
            
            const sheetName = sheet.name;
            const values = usedRange.values;
            const startCol = usedRange.columnIndex;
            const startRow = usedRange.rowIndex;
            const rowCount = usedRange.rowCount;
            const colCount = usedRange.columnCount;
            
            // Detect headers (first row)
            const headers = values[0] || [];
            
            // Build column mapping
            const columnMap = [];
            for (let c = 0; c < colCount; c++) {
                const colLetter = colIndexToLetter(startCol + c);
                const headerName = headers[c] || `Column ${colLetter}`;
                columnMap.push({
                    letter: colLetter,
                    index: c,
                    header: headerName
                });
            }
            
            state.currentData = {
                sheetName,
                address: usedRange.address,
                values,
                headers,
                columnMap,
                startRow: startRow + 1, // 1-based
                startCol: colIndexToLetter(startCol),
                rowCount,
                colCount,
                dataStartRow: startRow + 2 // Data starts after header (1-based)
            };
            
            infoEl.textContent = `${sheetName}: ${rowCount} rows × ${colCount} cols`;
        });
    } catch (e) {
        infoEl.textContent = "No data";
        state.currentData = null;
    }
}

// ============================================================================
// Chat UI
// ============================================================================
function showChat() {
    if (state.isFirstMessage) {
        document.getElementById("welcome").style.display = "none";
        document.getElementById("chat").style.display = "flex";
        state.isFirstMessage = false;
    }
}

function addMessage(role, content, type = "") {
    showChat();
    const chat = document.getElementById("chat");
    const msg = document.createElement("div");
    msg.className = `msg ${role} ${type}`;
    msg.innerHTML = `
        <div class="msg-avatar">${role === "user" ? "U" : "AI"}</div>
        <div class="msg-body">${formatText(content)}</div>
    `;
    chat.appendChild(msg);
    chat.scrollTop = chat.scrollHeight;
}

function formatText(text) {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/```(\w*)\n?([\s\S]*?)```/g, '<pre>$2</pre>')
        .replace(/`([^`]+)`/g, '<code>$1</code>')
        .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
        .replace(/\n/g, '<br>');
}

function showTyping() {
    showChat();
    const chat = document.getElementById("chat");
    const el = document.createElement("div");
    el.className = "msg ai";
    el.id = "typing";
    el.innerHTML = `<div class="msg-avatar">AI</div><div class="msg-body"><span class="dot"></span><span class="dot"></span><span class="dot"></span></div>`;
    chat.appendChild(el);
    chat.scrollTop = chat.scrollHeight;
}

function hideTyping() {
    document.getElementById("typing")?.remove();
}

function clearChat() {
    state.conversationHistory = [];
    state.pendingActions = [];
    state.isFirstMessage = true;
    state.preview.selections = [];
    state.preview.expandedIndex = -1;
    document.getElementById("chat").innerHTML = "";
    document.getElementById("chat").style.display = "none";
    document.getElementById("welcome").style.display = "flex";
    document.getElementById("applyBtn").disabled = true;
    hidePreviewPanel();
    toast("Cleared");
}

function toast(msg) {
    const t = document.getElementById("toast");
    t.textContent = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 2000);
}

// ============================================================================
// AI Communication
// ============================================================================
async function handleSend() {
    const input = document.getElementById("promptInput");
    const prompt = input.value.trim();
    if (!prompt) return;
    
    if (!state.apiKey) {
        document.getElementById("settingsBtn").click();
        toast("Enter API key");
        return;
    }
    
    await readExcelData();
    
    addMessage("user", prompt);
    
    input.value = "";
    input.style.height = "auto";
    document.getElementById("sendBtn").disabled = true;
    
    showTyping();
    
    try {
        const response = await callAI(prompt);
        hideTyping();
        
        const { message, actions } = parseResponse(response);
        state.pendingActions = actions;
        
        addMessage("ai", message, actions.length ? "has-action" : "");
        
        if (actions.length) {
            // Initialize preview state and show preview panel
            state.preview.selections = actions.map(() => true);
            state.preview.expandedIndex = -1;
            showPreviewPanel();
        } else {
            hidePreviewPanel();
        }
        
        state.conversationHistory.push(
            { role: "user", parts: [{ text: prompt }] },
            { role: "model", parts: [{ text: response }] }
        );
        
        if (state.conversationHistory.length > CONFIG.MAX_HISTORY * 2) {
            state.conversationHistory = state.conversationHistory.slice(-CONFIG.MAX_HISTORY * 2);
        }
    } catch (err) {
        hideTyping();
        addMessage("ai", "Error: " + err.message, "error");
    }
}

async function callAI(userPrompt) {
    const dataContext = buildDataContext();
    const systemPrompt = getSystemPrompt();
    
    const fullUserMessage = `${dataContext}\n\n---\nUSER REQUEST: ${userPrompt}`;
    
    const contents = [...state.conversationHistory];
    contents.push({ role: "user", parts: [{ text: fullUserMessage }] });
    
    const res = await fetch(
        `${CONFIG.API_ENDPOINT}${CONFIG.GEMINI_MODEL}:generateContent?key=${state.apiKey}`,
        {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                systemInstruction: { parts: [{ text: systemPrompt }] },
                contents,
                generationConfig: { temperature: 0.1, maxOutputTokens: 4096 }
            })
        }
    );
    
    if (!res.ok) throw new Error(`API Error: ${res.status}`);
    
    const data = await res.json();
    return data?.candidates?.[0]?.content?.parts?.[0]?.text || "No response";
}

function buildDataContext() {
    if (!state.currentData) {
        return "ERROR: No Excel data available.";
    }
    
    const { sheetName, values, columnMap, rowCount, colCount, dataStartRow, address } = state.currentData;
    
    let context = `## EXCEL WORKBOOK DATA\n\n`;
    context += `**Sheet:** ${sheetName}\n`;
    context += `**Data Range:** ${address}\n`;
    context += `**Total Rows:** ${rowCount} (including header)\n`;
    context += `**Total Columns:** ${colCount}\n\n`;
    
    // Column structure - CRITICAL for AI to understand
    context += `## COLUMN STRUCTURE\n`;
    context += `| Column Letter | Header Name | Sample Values |\n`;
    context += `|---------------|-------------|---------------|\n`;
    
    for (const col of columnMap) {
        // Get sample values from first few data rows
        const samples = [];
        for (let r = 1; r < Math.min(4, values.length); r++) {
            const val = values[r]?.[col.index];
            if (val !== null && val !== undefined && val !== "") {
                samples.push(String(val).substring(0, 20));
            }
        }
        context += `| ${col.letter} | ${col.header} | ${samples.join(", ")} |\n`;
    }
    
    context += `\n## DATA PREVIEW (First 30 rows)\n\n`;
    
    // Header row
    context += `| Row |`;
    for (const col of columnMap) {
        context += ` ${col.letter}: ${col.header} |`;
    }
    context += `\n|-----|`;
    for (let c = 0; c < colCount; c++) {
        context += `------------|`;
    }
    context += `\n`;
    
    // Data rows
    const maxRows = Math.min(30, values.length);
    for (let r = 0; r < maxRows; r++) {
        const rowNum = state.currentData.startRow + r;
        context += `| ${rowNum} |`;
        for (let c = 0; c < colCount; c++) {
            let val = values[r]?.[c];
            if (val === null || val === undefined) val = "";
            val = String(val).substring(0, 25);
            context += ` ${val} |`;
        }
        context += `\n`;
    }
    
    if (rowCount > 30) {
        context += `\n... and ${rowCount - 30} more rows\n`;
    }
    
    // Add unique values for key columns (for dropdowns)
    context += `\n## UNIQUE VALUES IN EACH COLUMN (for dropdowns)\n`;
    for (const col of columnMap) {
        const uniqueVals = new Set();
        for (let r = 1; r < values.length; r++) {
            const val = values[r]?.[col.index];
            if (val !== null && val !== undefined && val !== "") {
                uniqueVals.add(val);
            }
        }
        if (uniqueVals.size > 0 && uniqueVals.size <= 50) {
            context += `**${col.letter} (${col.header}):** ${Array.from(uniqueVals).slice(0, 20).join(", ")}`;
            if (uniqueVals.size > 20) context += ` ... (${uniqueVals.size} total)`;
            context += `\n`;
        }
    }
    
    return context;
}

function getSystemPrompt() {
    return `You are Excel Copilot, an expert Excel assistant. You have COMPLETE access to the user's Excel data.

## CRITICAL RULES

1. **CHECK THE COLUMN STRUCTURE TABLE** - Find the exact column letter for each header name
2. **State column might be E, F, or any letter** - ALWAYS verify by looking at the table
3. **Data starts at row 2** (row 1 is headers)

## HOW TO CREATE A DROPDOWN

For a dropdown of values from a column:
1. Look at COLUMN STRUCTURE to find the correct column letter
2. Use the validation action with the source range

Example: If "State" is in column E with data from row 2 to row 100:

<ACTION type="validation" target="K10" source="E2:E100">
</ACTION>

That's it! Just one action. The source should be the actual data range (e.g., E2:E100).

## ACTION TYPES

**Dropdown/Validation:**
<ACTION type="validation" target="CELL" source="DATARANGE">
</ACTION>

**Formula:**
<ACTION type="formula" target="CELL_OR_RANGE">
=YOUR_FORMULA
</ACTION>

**Values:**
<ACTION type="values" target="RANGE">
[["val1","val2"],["val3","val4"]]
</ACTION>

**Format:**
<ACTION type="format" target="RANGE">
{"bold":true,"fill":"#4472C4","fontColor":"#FFFFFF"}
</ACTION>

**Chart:**
<ACTION type="chart" target="DATARANGE" chartType="TYPE" title="TITLE" position="CELL">
</ACTION>

## CHART TYPES
- **column** - Vertical bar chart (default, good for comparing categories)
- **bar** - Horizontal bar chart (good for long category names)
- **line** - Line chart (good for trends over time)
- **pie** - Pie chart (good for showing parts of a whole, use with 1 data series)
- **area** - Area chart (good for cumulative totals over time)
- **scatter** - XY Scatter plot (good for correlation between 2 variables)
- **doughnut** - Like pie but with hole in center
- **radar** - Spider/radar chart (good for comparing multiple variables)

## CHART EXAMPLES

**Sales by Region (Column Chart):**
<ACTION type="chart" target="A1:B10" chartType="column" title="Sales by Region" position="H2">
</ACTION>

**Trend Over Time (Line Chart):**
<ACTION type="chart" target="A1:C20" chartType="line" title="Monthly Trend" position="H2">
</ACTION>

**Market Share (Pie Chart):**
<ACTION type="chart" target="A1:B5" chartType="pie" title="Market Share" position="H2">
</ACTION>

**Comparison (Bar Chart):**
<ACTION type="chart" target="A1:D10" chartType="bar" title="Product Comparison" position="H2">
</ACTION>

## CHART TIPS
- Include headers in the data range (first row/column as labels)
- For pie charts, use only 2 columns (labels + values)
- Position is where the top-left of chart will be placed
- Choose chart type based on what story the data tells

## IMPORTANT

- ALWAYS check COLUMN STRUCTURE first
- For dropdowns, source is the data column range (e.g., E2:E100)
- Don't use UNIQUE formula for dropdowns - just use the source range directly
- For charts, include the header row in the target range`;
}

function parseResponse(text) {
    const actions = [];
    const actionRegex = /<ACTION\s+([^>]*)>([\s\S]*?)<\/ACTION>/g;
    
    let match;
    while ((match = actionRegex.exec(text)) !== null) {
        const attrs = match[1];
        const content = match[2].trim();
        
        const type = attrs.match(/type="([^"]+)"/)?.[1] || "formula";
        const target = attrs.match(/target="([^"]+)"/)?.[1] || "";
        const source = attrs.match(/source="([^"]+)"/)?.[1] || "";
        const chartType = attrs.match(/chartType="([^"]+)"/)?.[1] || "column";
        const title = attrs.match(/title="([^"]+)"/)?.[1] || "";
        const position = attrs.match(/position="([^"]+)"/)?.[1] || "H2";
        
        actions.push({ type, target, source, chartType, title, position, data: content });
    }
    
    const message = text.replace(/<ACTION[\s\S]*?<\/ACTION>/g, "").trim();
    return { message: message || "Ready to apply.", actions };
}

// ============================================================================
// Preview Panel Functions
// ============================================================================

/**
 * Gets the SVG icon for an action type
 * @param {string} type - Action type
 * @returns {string} SVG icon markup
 */
function getActionIcon(type) {
    const icons = {
        formula: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7h16M4 12h16M4 17h10"/><text x="18" y="18" font-size="8" fill="currentColor">fx</text></svg>',
        values: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/></svg>',
        format: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 20h16M6 16l6-12 6 12M8 12h8"/></svg>',
        chart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>',
        validation: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>',
        sort: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h12M3 18h6"/></svg>',
        autofill: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>'
    };
    return icons[type] || icons.formula;
}

/**
 * Gets a summary string for an action
 * @param {Object} action - The action to summarize
 * @returns {string} Human-readable summary
 */
function getActionSummary(action) {
    const typeLabels = {
        formula: "Formula",
        values: "Values",
        format: "Format",
        chart: "Chart",
        validation: "Dropdown",
        sort: "Sort",
        autofill: "Autofill"
    };
    return typeLabels[action.type] || action.type;
}

/**
 * Gets detailed description for an action
 * @param {Object} action - The action to describe
 * @returns {string} Detailed description
 */
function getActionDetails(action) {
    switch (action.type) {
        case "formula":
            return action.data || "No formula";
        case "values":
            try {
                const vals = JSON.parse(action.data);
                return JSON.stringify(vals, null, 2);
            } catch {
                return action.data || "No values";
            }
        case "format":
            try {
                const fmt = JSON.parse(action.data);
                const parts = [];
                if (fmt.bold) parts.push("Bold");
                if (fmt.italic) parts.push("Italic");
                if (fmt.fill) parts.push(`Fill: ${fmt.fill}`);
                if (fmt.fontColor) parts.push(`Color: ${fmt.fontColor}`);
                if (fmt.fontSize) parts.push(`Size: ${fmt.fontSize}`);
                if (fmt.numberFormat) parts.push(`Format: ${fmt.numberFormat}`);
                return parts.join(", ") || "No formatting";
            } catch {
                return action.data || "No format data";
            }
        case "chart":
            return `Type: ${action.chartType}\nData: ${action.target}\nPosition: ${action.position}${action.title ? `\nTitle: ${action.title}` : ""}`;
        case "validation":
            return `Source: ${action.source}`;
        case "sort":
            return action.data || "Default sort";
        case "autofill":
            return `Source: ${action.source}`;
        default:
            return action.data || "No details";
    }
}

/**
 * Filters actions based on selection state
 * @param {Object[]} actions - All pending actions
 * @param {boolean[]} selections - Selection state for each action
 * @returns {Object[]} Only the selected actions
 */
function filterSelectedActions(actions, selections) {
    if (!actions || !selections) return [];
    return actions.filter((_, index) => selections[index] === true);
}

/**
 * Checks if any actions are selected
 * @param {boolean[]} selections - Selection state array
 * @returns {boolean} True if at least one action is selected
 */
function hasSelectedActions(selections) {
    if (!selections || selections.length === 0) return false;
    return selections.some(s => s === true);
}

/**
 * Renders a single preview item HTML
 */
function renderPreviewItemHTML(action, index, isExpanded, isSelected, hasWarning) {
    const icon = getActionIcon(action.type);
    const summary = getActionSummary(action);
    const details = getActionDetails(action);
    const expandedClass = isExpanded ? 'expanded' : '';
    const warningClass = hasWarning ? 'warning' : '';
    
    return `
        <div class="preview-item ${expandedClass} ${warningClass}" data-index="${index}">
            <input type="checkbox" class="preview-checkbox" ${isSelected ? 'checked' : ''} data-index="${index}">
            <div class="preview-icon ${action.type}">${icon}</div>
            <div class="preview-content">
                <div class="preview-summary">
                    ${summary}
                    <span class="preview-target">${action.target}</span>
                    ${hasWarning ? '<svg class="preview-warning" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2L1 21h22L12 2zm0 3.5L19.5 19h-15L12 5.5zM11 10v4h2v-4h-2zm0 6v2h2v-2h-2z"/></svg>' : ''}
                </div>
                <div class="preview-details">${details}</div>
            </div>
            <div class="preview-expand">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 9l6 6 6-6"/></svg>
            </div>
        </div>
    `;
}

/**
 * Shows the preview panel with pending actions
 */
function showPreviewPanel() {
    const panel = document.getElementById("previewPanel");
    const list = document.getElementById("previewList");
    
    if (!state.pendingActions.length) {
        panel.style.display = "none";
        return;
    }
    
    // Initialize selections if needed (all selected by default)
    if (state.preview.selections.length !== state.pendingActions.length) {
        state.preview.selections = state.pendingActions.map(() => true);
    }
    
    // Render preview items
    const html = state.pendingActions.map((action, index) => {
        const isExpanded = index === state.preview.expandedIndex;
        const isSelected = state.preview.selections[index];
        return renderPreviewItemHTML(action, index, isExpanded, isSelected, false);
    }).join('');
    
    list.innerHTML = html;
    panel.style.display = "block";
    
    // Bind events
    bindPreviewEvents();
    updateApplyButtonState();
}

/**
 * Hides the preview panel
 */
function hidePreviewPanel() {
    const panel = document.getElementById("previewPanel");
    panel.style.display = "none";
    state.preview.selections = [];
    state.preview.expandedIndex = -1;
}

/**
 * Binds event handlers to preview panel elements
 */
function bindPreviewEvents() {
    // Checkbox changes
    document.querySelectorAll(".preview-checkbox").forEach(cb => {
        cb.addEventListener("change", (e) => {
            const index = parseInt(e.target.dataset.index);
            state.preview.selections[index] = e.target.checked;
            updateApplyButtonState();
        });
        // Stop propagation to prevent expand/collapse when clicking checkbox
        cb.addEventListener("click", (e) => e.stopPropagation());
    });
    
    // Expand/collapse on item click
    document.querySelectorAll(".preview-item").forEach(item => {
        item.addEventListener("click", (e) => {
            if (e.target.classList.contains("preview-checkbox")) return;
            const index = parseInt(item.dataset.index);
            toggleExpand(index);
        });
        
        // Hover highlighting
        item.addEventListener("mouseenter", () => {
            const index = parseInt(item.dataset.index);
            highlightRange(state.pendingActions[index]?.target);
        });
        
        item.addEventListener("mouseleave", () => {
            clearHighlight();
        });
    });
    
    // Select all button
    document.getElementById("selectAllBtn")?.addEventListener("click", toggleSelectAll);
}

/**
 * Toggles expand/collapse for a preview item
 */
function toggleExpand(index) {
    state.preview.expandedIndex = state.preview.expandedIndex === index ? -1 : index;
    showPreviewPanel(); // Re-render
}

/**
 * Toggles select all / deselect all
 */
function toggleSelectAll() {
    const allSelected = state.preview.selections.every(s => s);
    state.preview.selections = state.preview.selections.map(() => !allSelected);
    showPreviewPanel(); // Re-render
}

/**
 * Updates the Apply button state based on selections
 */
function updateApplyButtonState() {
    const applyBtn = document.getElementById("applyBtn");
    const hasSelected = hasSelectedActions(state.preview.selections);
    applyBtn.disabled = !hasSelected;
}

/**
 * Highlights a range in Excel
 */
async function highlightRange(rangeAddress) {
    if (!rangeAddress) return;
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(rangeAddress);
            range.select();
            await ctx.sync();
        });
    } catch (e) {
        // Silently fail - range might be invalid
        console.warn("Could not highlight range:", rangeAddress, e);
    }
}

/**
 * Clears any active highlighting (no-op for now, selection persists)
 */
async function clearHighlight() {
    // Excel doesn't have a "clear selection" API
    // The selection will change when user interacts with Excel
}

// ============================================================================
// History and Undo Functions
// ============================================================================

/**
 * Generates a unique ID for history entries
 */
function generateHistoryId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

/**
 * Captures the current state of a range for undo
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {string} rangeAddress - The range to capture
 * @returns {Promise<Object>} The captured undo data
 */
async function captureUndoData(ctx, sheet, rangeAddress) {
    try {
        const range = sheet.getRange(rangeAddress);
        range.load(["values", "formulas", "address"]);
        await ctx.sync();
        
        return {
            values: range.values,
            formulas: range.formulas,
            address: range.address
        };
    } catch (e) {
        console.warn("Could not capture undo data:", e);
        return null;
    }
}

/**
 * Adds an action to history
 */
function addActionToHistory(action, undoData) {
    const entry = {
        id: generateHistoryId(),
        type: action.type,
        target: action.target,
        timestamp: Date.now(),
        undoData: undoData
    };
    
    // Prepend to history
    state.history.entries = [entry, ...state.history.entries];
    
    // Enforce max limit
    if (state.history.entries.length > state.history.maxEntries) {
        state.history.entries = state.history.entries.slice(0, state.history.maxEntries);
    }
    
    updateUndoButtonState();
    if (state.history.panelVisible) {
        renderHistoryPanel();
    }
}

/**
 * Performs undo of the most recent action
 */
async function performUndo() {
    if (!state.history.entries.length) {
        toast("Nothing to undo");
        return;
    }
    
    const entry = state.history.entries[0];
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(entry.undoData.address);
            
            // Restore formulas (which also restores values for non-formula cells)
            range.formulas = entry.undoData.formulas;
            await ctx.sync();
        });
        
        // Remove from history
        state.history.entries = state.history.entries.slice(1);
        
        updateUndoButtonState();
        if (state.history.panelVisible) {
            renderHistoryPanel();
        }
        
        toast("Undone");
        await readExcelData();
    } catch (e) {
        console.error("Undo failed:", e);
        toast("Undo failed");
        // Keep entry in history on failure
    }
}

/**
 * Updates the Undo button state
 */
function updateUndoButtonState() {
    const undoBtn = document.getElementById("undoBtn");
    if (undoBtn) {
        undoBtn.disabled = state.history.entries.length === 0;
    }
}

/**
 * Formats a timestamp as relative time
 */
function formatRelativeTime(timestamp) {
    const now = Date.now();
    const diff = now - timestamp;
    
    const seconds = Math.floor(diff / 1000);
    const minutes = Math.floor(seconds / 60);
    const hours = Math.floor(minutes / 60);
    const days = Math.floor(hours / 24);
    
    if (seconds < 60) return 'just now';
    if (minutes < 60) return `${minutes} min ago`;
    if (hours < 24) return `${hours} hr ago`;
    if (days === 1) return 'yesterday';
    return `${days} days ago`;
}

/**
 * Renders the history panel
 */
function renderHistoryPanel() {
    const list = document.getElementById("historyList");
    if (!list) return;
    
    if (!state.history.entries.length) {
        list.innerHTML = '<div class="history-empty">No actions yet</div>';
        return;
    }
    
    const typeLabels = {
        formula: "Formula",
        values: "Values",
        format: "Format",
        chart: "Chart",
        validation: "Dropdown",
        sort: "Sort",
        autofill: "Autofill"
    };
    
    const html = state.history.entries.map(entry => {
        const icon = getActionIcon(entry.type);
        const label = typeLabels[entry.type] || entry.type;
        const timeStr = formatRelativeTime(entry.timestamp);
        
        return `
            <div class="history-entry" data-id="${entry.id}">
                <div class="history-icon ${entry.type}">${icon}</div>
                <div class="history-content">
                    <span class="history-label">${label}</span>
                    <span class="history-target">${entry.target}</span>
                </div>
                <span class="history-time">${timeStr}</span>
            </div>
        `;
    }).join('');
    
    list.innerHTML = html;
}

/**
 * Toggles history panel visibility
 */
function toggleHistoryPanel() {
    state.history.panelVisible = !state.history.panelVisible;
    const panel = document.getElementById("historyPanel");
    const btn = document.getElementById("historyBtn");
    
    if (panel) {
        panel.style.display = state.history.panelVisible ? "block" : "none";
        if (state.history.panelVisible) {
            renderHistoryPanel();
        }
    }
    
    if (btn) {
        btn.classList.toggle("active", state.history.panelVisible);
    }
}

// ============================================================================
// Apply Actions
// ============================================================================
async function handleApply() {
    // Get only selected actions
    const selectedActions = filterSelectedActions(state.pendingActions, state.preview.selections);
    
    if (!selectedActions.length) {
        toast("Nothing to apply");
        return;
    }
    
    const applyBtn = document.getElementById("applyBtn");
    applyBtn.disabled = true;
    applyBtn.textContent = "Applying...";
    
    let successCount = 0;
    let errorMsg = "";
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            
            for (const action of selectedActions) {
                try {
                    // Capture undo data before applying (skip for charts - can't undo)
                    let undoData = null;
                    if (action.type !== 'chart') {
                        undoData = await captureUndoData(ctx, sheet, action.target);
                    }
                    
                    await executeAction(ctx, sheet, action);
                    await ctx.sync();
                    successCount++;
                    
                    // Add to history if we have undo data
                    if (undoData) {
                        addActionToHistory(action, undoData);
                    }
                } catch (e) {
                    errorMsg = e.message;
                    console.error("Action failed:", e);
                }
            }
        });
        
        if (successCount === selectedActions.length) {
            addMessage("ai", `${successCount} change${successCount > 1 ? 's' : ''} applied successfully.`, "success");
            toast("Applied");
        } else if (successCount > 0) {
            addMessage("ai", `${successCount}/${selectedActions.length} changes applied. Error: ${errorMsg}`, "error");
        } else {
            addMessage("ai", `Failed: ${errorMsg}`, "error");
        }
        
        // Clear pending actions and hide preview
        state.pendingActions = [];
        hidePreviewPanel();
        await readExcelData();
    } catch (err) {
        addMessage("ai", "Failed: " + err.message, "error");
        toast("Failed");
    }
    
    applyBtn.disabled = true;
    applyBtn.textContent = "Apply Changes";
}

async function executeAction(ctx, sheet, action) {
    const { type, target, source, validationType, chartType, data } = action;
    
    if (!target) throw new Error("No target specified");
    
    const range = sheet.getRange(target);
    range.load(["rowCount", "columnCount"]);
    await ctx.sync();
    
    switch (type) {
        case "formula":
            await applyFormula(range, data);
            break;
            
        case "values":
            applyValues(range, data);
            break;
            
        case "format":
            applyFormat(range, data);
            break;
            
        case "validation":
            await applyValidation(ctx, sheet, range, source);
            break;
            
        case "chart":
            createChart(sheet, range, action);
            break;
            
        case "sort":
            applySort(range, data);
            break;
            
        case "autofill":
            if (source) {
                const sourceRange = sheet.getRange(source);
                sourceRange.autoFill(range, Excel.AutoFillType.fillDefault);
            }
            break;
            
        default:
            if (data) range.values = [[data]];
    }
}

async function applyFormula(range, formula) {
    const rows = range.rowCount;
    const cols = range.columnCount;
    
    if (rows === 1 && cols === 1) {
        range.formulas = [[formula]];
        return;
    }
    
    const formulas = [];
    for (let r = 0; r < rows; r++) {
        const rowFormulas = [];
        for (let c = 0; c < cols; c++) {
            let f = formula;
            if (r > 0) {
                f = formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (m, d1, col, d2, row) => {
                    if (d2 === "$") return m;
                    return `${d1}${col}${d2}${parseInt(row) + r}`;
                });
            }
            rowFormulas.push(f);
        }
        formulas.push(rowFormulas);
    }
    
    range.formulas = formulas;
}

function applyValues(range, data) {
    let values;
    try {
        values = JSON.parse(data);
        if (!Array.isArray(values)) values = [[values]];
        if (!Array.isArray(values[0])) values = [values];
    } catch {
        values = [[data]];
    }
    range.values = values;
}

function applyFormat(range, data) {
    let fmt;
    try { fmt = JSON.parse(data); } catch { fmt = {}; }
    
    if (fmt.bold) range.format.font.bold = true;
    if (fmt.italic) range.format.font.italic = true;
    if (fmt.fill) range.format.fill.color = fmt.fill;
    if (fmt.fontColor) range.format.font.color = fmt.fontColor;
    if (fmt.fontSize) range.format.font.size = fmt.fontSize;
    if (fmt.numberFormat) range.numberFormat = [[fmt.numberFormat]];
    if (fmt.border) {
        range.format.borders.getItem("EdgeTop").style = "Continuous";
        range.format.borders.getItem("EdgeBottom").style = "Continuous";
        range.format.borders.getItem("EdgeLeft").style = "Continuous";
        range.format.borders.getItem("EdgeRight").style = "Continuous";
    }
}

async function applyValidation(ctx, sheet, range, source) {
    if (source) {
        // Clear any existing validation first
        range.dataValidation.clear();
        await ctx.sync();
        
        // Get the source range to extract unique values
        const sourceRange = sheet.getRange(source);
        sourceRange.load("values");
        await ctx.sync();
        
        // Extract unique non-empty values
        const uniqueValues = [];
        const seen = new Set();
        for (const row of sourceRange.values) {
            const val = row[0];
            if (val !== null && val !== undefined && val !== "" && !seen.has(val)) {
                seen.add(val);
                uniqueValues.push(String(val));
            }
        }
        
        // Create comma-separated list for validation
        const listSource = uniqueValues.join(",");
        
        // Set the validation rule with explicit list
        range.dataValidation.rule = {
            list: {
                inCellDropDown: true,
                source: listSource
            }
        };
    }
}

function createChart(sheet, dataRange, action) {
    const { chartType, data } = action;
    const ct = (chartType || "column").toLowerCase();
    
    // Parse additional options from data if provided
    let title = "Chart";
    let position = "H2";
    
    // Try to extract title and position from action attributes or data
    if (action.title) title = action.title;
    if (action.position) position = action.position;
    
    // Determine chart type
    let type = Excel.ChartType.columnClustered;
    
    if (ct.includes("line")) {
        type = Excel.ChartType.line;
    } else if (ct.includes("pie")) {
        type = Excel.ChartType.pie;
    } else if (ct.includes("doughnut") || ct.includes("donut")) {
        type = Excel.ChartType.doughnut;
    } else if (ct.includes("bar")) {
        type = Excel.ChartType.barClustered;
    } else if (ct.includes("area")) {
        type = Excel.ChartType.area;
    } else if (ct.includes("scatter") || ct.includes("xy")) {
        type = Excel.ChartType.xyscatter;
    } else if (ct.includes("radar") || ct.includes("spider")) {
        type = Excel.ChartType.radar;
    } else if (ct.includes("stacked")) {
        if (ct.includes("bar")) {
            type = Excel.ChartType.barStacked;
        } else {
            type = Excel.ChartType.columnStacked;
        }
    }
    
    // Create the chart
    const chart = sheet.charts.add(type, dataRange, Excel.ChartSeriesBy.auto);
    
    // Calculate end position (chart size roughly 8 cols x 15 rows)
    const startCol = position.match(/[A-Z]+/)?.[0] || "H";
    const startRow = parseInt(position.match(/\d+/)?.[0] || "2");
    const endCol = String.fromCharCode(startCol.charCodeAt(0) + 8);
    const endRow = startRow + 15;
    const endPosition = `${endCol}${endRow}`;
    
    chart.setPosition(position, endPosition);
    
    // Set title
    chart.title.text = title;
    chart.title.visible = true;
    
    // Style the chart
    chart.legend.visible = true;
    chart.legend.position = Excel.ChartLegendPosition.bottom;
    
    // For pie charts, show data labels
    if (ct.includes("pie") || ct.includes("doughnut")) {
        chart.legend.position = Excel.ChartLegendPosition.right;
    }
}

function applySort(range, data) {
    const opts = typeof data === "string" ? JSON.parse(data) : (data || {});
    range.sort.apply([{ key: opts.column || 0, ascending: opts.ascending !== false }]);
}

export { handleSend, handleApply, readExcelData, clearChat };
