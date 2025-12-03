/*
 * Excel AI Copilot - Fixed Data Reading
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
    isFirstMessage: true
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
        toast("Data refreshed");
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
    
    // Auto-refresh on selection change
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
// Read Excel Data - ALWAYS reads data
// ============================================================================
async function readExcelData() {
    const infoEl = document.getElementById("contextInfo");
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const selection = ctx.workbook.getSelectedRange();
            const usedRange = sheet.getUsedRange();
            
            sheet.load("name");
            selection.load(["address", "values", "formulas", "rowCount", "columnCount"]);
            usedRange.load(["address", "values", "rowCount", "columnCount"]);
            
            await ctx.sync();
            
            const sheetName = sheet.name;
            
            // Get selection data
            const selectionData = {
                address: selection.address,
                values: selection.values,
                formulas: selection.formulas,
                rows: selection.rowCount,
                cols: selection.columnCount
            };
            
            // Get full sheet data
            const fullSheetData = {
                address: usedRange.address,
                values: usedRange.values,
                rows: usedRange.rowCount,
                cols: usedRange.columnCount
            };
            
            // Determine if selection is meaningful (more than 1 cell or has data)
            const hasSelection = selectionData.rows > 1 || selectionData.cols > 1 || 
                                 (selectionData.values[0]?.[0] !== null && selectionData.values[0]?.[0] !== "");
            
            state.currentData = {
                sheetName,
                selection: selectionData,
                fullSheet: fullSheetData,
                hasSelection
            };
            
            // Update UI
            if (hasSelection) {
                infoEl.textContent = `${selectionData.address} (${selectionData.rows}×${selectionData.cols})`;
            } else {
                infoEl.textContent = `Sheet: ${sheetName} (${fullSheetData.rows}×${fullSheetData.cols})`;
            }
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
    document.getElementById("chat").innerHTML = "";
    document.getElementById("chat").style.display = "none";
    document.getElementById("welcome").style.display = "flex";
    document.getElementById("applyBtn").disabled = true;
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
    
    // Always refresh data before sending
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
            document.getElementById("applyBtn").disabled = false;
        }
        
        // Save to history
        state.conversationHistory.push(
            { role: "user", parts: [{ text: prompt }] },
            { role: "model", parts: [{ text: response }] }
        );
        
        // Trim history
        if (state.conversationHistory.length > CONFIG.MAX_HISTORY * 2) {
            state.conversationHistory = state.conversationHistory.slice(-CONFIG.MAX_HISTORY * 2);
        }
    } catch (err) {
        hideTyping();
        addMessage("ai", "Error: " + err.message, "error");
    }
}

async function callAI(userPrompt) {
    // Build the full prompt with data
    const dataContext = buildDataContext();
    const systemPrompt = getSystemPrompt();
    
    // Create message with data included
    const fullUserMessage = `${dataContext}\n\nUser request: ${userPrompt}`;
    
    // Build contents
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
                generationConfig: { temperature: 0.2, maxOutputTokens: 4096 }
            })
        }
    );
    
    if (!res.ok) {
        const errText = await res.text();
        throw new Error(`API Error: ${res.status}`);
    }
    
    const data = await res.json();
    return data?.candidates?.[0]?.content?.parts?.[0]?.text || "No response";
}

function buildDataContext() {
    if (!state.currentData) {
        return "NO DATA AVAILABLE - Please ensure Excel has data.";
    }
    
    const { sheetName, selection, fullSheet, hasSelection } = state.currentData;
    
    let context = `=== EXCEL DATA ===\nSheet: ${sheetName}\n\n`;
    
    // Always include selection info
    context += `SELECTED RANGE: ${selection.address} (${selection.rows} rows × ${selection.cols} cols)\n`;
    context += `SELECTED DATA:\n${formatDataAsTable(selection.values)}\n\n`;
    
    // Also include full sheet summary if different from selection
    if (!hasSelection || fullSheet.rows > selection.rows || fullSheet.cols > selection.cols) {
        context += `FULL SHEET RANGE: ${fullSheet.address} (${fullSheet.rows} rows × ${fullSheet.cols} cols)\n`;
        
        // Include full sheet data (limited to first 100 rows)
        const limitedValues = fullSheet.values.slice(0, 100);
        context += `FULL SHEET DATA (first ${Math.min(100, fullSheet.rows)} rows):\n${formatDataAsTable(limitedValues)}\n`;
    }
    
    return context;
}

function formatDataAsTable(values) {
    if (!values || !values.length) return "(empty)";
    
    let table = "";
    const maxCols = Math.min(values[0]?.length || 0, 20);
    
    for (let r = 0; r < values.length; r++) {
        const row = [];
        for (let c = 0; c < maxCols; c++) {
            let val = values[r]?.[c];
            if (val === null || val === undefined) val = "";
            if (typeof val === "number") val = val.toString();
            if (typeof val === "string" && val.length > 50) val = val.substring(0, 50) + "...";
            row.push(val);
        }
        table += `Row ${r + 1}: ${row.join(" | ")}\n`;
    }
    
    return table;
}

function getSystemPrompt() {
    return `You are Excel Copilot, an expert Excel assistant.

## YOUR CAPABILITIES
- Analyze data and provide insights
- Create formulas (SUM, AVERAGE, VLOOKUP, IF, etc.)
- Format cells and create tables
- Create charts
- Sort and filter data
- Find patterns, trends, outliers

## IMPORTANT RULES
1. You ALWAYS receive the current Excel data in the user's message
2. Analyze the ACTUAL data provided - never ask for data again
3. Be specific - reference actual cell values, column names, row numbers
4. When creating formulas, specify exact cell ranges

## RESPONSE FORMAT
For analysis: Provide clear insights based on the actual data
For actions: Include action blocks like this:

<ACTION type="formula" target="E2:E10">
=SUM(B2:D2)
</ACTION>

<ACTION type="values" target="A1:C1">
[["Name","Age","City"]]
</ACTION>

<ACTION type="format" target="A1:E1">
{"bold":true,"fill":"#4472C4","fontColor":"#FFFFFF"}
</ACTION>

<ACTION type="chart" target="A1:D10">
column
</ACTION>

## ACTION TYPES
- formula: Apply Excel formula
- values: Set cell values (JSON 2D array)
- format: Apply formatting (JSON with bold, fill, fontColor, fontSize, border, numberFormat)
- chart: Create chart (line, bar, pie, column, area)
- autofill: Fill formula from source to target
- sort: Sort data

Always analyze the provided data directly. Never say you can't see the data.`;
}

function parseResponse(text) {
    const actions = [];
    const actionRegex = /<ACTION\s+([^>]+)>([\s\S]*?)<\/ACTION>/g;
    
    let match;
    while ((match = actionRegex.exec(text)) !== null) {
        const attrs = match[1];
        const content = match[2].trim();
        
        const type = attrs.match(/type="([^"]+)"/)?.[1] || "formula";
        const target = attrs.match(/target="([^"]+)"/)?.[1] || "selection";
        const source = attrs.match(/source="([^"]+)"/)?.[1] || "";
        
        actions.push({ type, target, source, data: content });
    }
    
    const message = text.replace(/<ACTION[\s\S]*?<\/ACTION>/g, "").trim();
    return { message: message || "Ready to apply.", actions };
}

// ============================================================================
// Apply Actions
// ============================================================================
async function handleApply() {
    if (!state.pendingActions.length) {
        toast("Nothing to apply");
        return;
    }
    
    const applyBtn = document.getElementById("applyBtn");
    applyBtn.disabled = true;
    applyBtn.textContent = "Applying...";
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            
            for (const action of state.pendingActions) {
                await executeAction(ctx, sheet, action);
            }
            
            await ctx.sync();
        });
        
        addMessage("ai", "Changes applied successfully.", "success");
        toast("Applied");
        state.pendingActions = [];
        await readExcelData();
    } catch (err) {
        addMessage("ai", "Failed: " + err.message, "error");
        toast("Failed");
        applyBtn.disabled = false;
    }
    
    applyBtn.textContent = "Apply Changes";
}

async function executeAction(ctx, sheet, action) {
    const { type, target, source, data } = action;
    
    let range;
    if (target === "selection") {
        range = ctx.workbook.getSelectedRange();
    } else {
        range = sheet.getRange(target);
    }
    
    range.load(["rowCount", "columnCount", "address"]);
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
        case "autofill":
            await applyAutofill(sheet, source, target);
            break;
        case "chart":
            createChart(sheet, range, data);
            break;
        case "sort":
            applySort(range, data);
            break;
        default:
            range.values = [[data]];
    }
}

async function applyFormula(range, formula) {
    const rows = range.rowCount;
    const cols = range.columnCount;
    
    if (rows === 1 && cols === 1) {
        range.formulas = [[formula]];
        return;
    }
    
    // Build formula array with adjusted references
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

async function applyAutofill(sheet, source, target) {
    const sourceRange = sheet.getRange(source);
    const targetRange = sheet.getRange(target);
    sourceRange.autoFill(targetRange, Excel.AutoFillType.fillDefault);
}

function createChart(sheet, dataRange, data) {
    let chartType = Excel.ChartType.columnClustered;
    const d = (data || "").toLowerCase();
    
    if (d.includes("line")) chartType = Excel.ChartType.line;
    else if (d.includes("pie")) chartType = Excel.ChartType.pie;
    else if (d.includes("bar")) chartType = Excel.ChartType.barClustered;
    else if (d.includes("area")) chartType = Excel.ChartType.area;
    
    const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
    chart.setPosition("H2", "P17");
    chart.title.text = "Chart";
}

function applySort(range, data) {
    const opts = typeof data === "string" ? { column: 0, ascending: true } : data;
    range.sort.apply([{ key: opts.column || 0, ascending: opts.ascending !== false }]);
}

export { handleSend, handleApply, readExcelData, clearChat };
