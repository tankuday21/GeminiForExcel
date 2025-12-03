/*
 * Excel AI Copilot - Professional Grade
 */

/* global document, Excel, Office, fetch, localStorage */

// ============================================================================
// Configuration
// ============================================================================
const CONFIG = {
    GEMINI_MODEL: "gemini-2.0-flash",
    API_ENDPOINT: "https://generativelanguage.googleapis.com/v1beta/models/",
    STORAGE_KEY: "excel_copilot_api_key",
    MAX_HISTORY: 10
};

// ============================================================================
// State
// ============================================================================
const state = {
    apiKey: "",
    pendingActions: [],
    selectionData: null,
    selectionAddress: "",
    sheetName: "",
    conversationHistory: [],
    isFirstMessage: true
};

// ============================================================================
// Initialize
// ============================================================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener("DOMContentLoaded", initApp);
        if (document.readyState === "complete" || document.readyState === "interactive") {
            initApp();
        }
    }
});

function initApp() {
    state.apiKey = localStorage.getItem(CONFIG.STORAGE_KEY) || "";
    bindEvents();
    refreshContext();
    setupAutoRefresh();
}

function bindEvents() {
    // Send
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
        autoResize(input);
    });
    
    // Apply
    document.getElementById("applyBtn")?.addEventListener("click", handleApply);
    
    // Refresh
    document.getElementById("refreshBtn")?.addEventListener("click", () => {
        const btn = document.getElementById("refreshBtn");
        btn.classList.add("loading");
        refreshContext().finally(() => btn.classList.remove("loading"));
    });
    
    // Context toggle
    document.getElementById("useContext")?.addEventListener("change", refreshContext);
    
    // Settings
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
        toast("API key saved");
    });
    document.getElementById("modal")?.addEventListener("click", (e) => {
        if (e.target.id === "modal") closeModal();
    });
    
    // Clear
    document.getElementById("clearBtn")?.addEventListener("click", clearChat);
    
    // Suggestions
    document.querySelectorAll("[data-prompt]").forEach(el => {
        el.addEventListener("click", () => {
            document.getElementById("promptInput").value = el.dataset.prompt;
            document.getElementById("sendBtn").disabled = false;
            handleSend();
        });
    });
    
    // Password toggle
    document.getElementById("togglePwd")?.addEventListener("click", () => {
        const inp = document.getElementById("apiKeyInput");
        inp.type = inp.type === "password" ? "text" : "password";
    });
}

function autoResize(el) {
    el.style.height = "auto";
    el.style.height = Math.min(el.scrollHeight, 120) + "px";
}

function closeModal() {
    document.getElementById("modal").classList.remove("open");
}

async function setupAutoRefresh() {
    try {
        await Excel.run(async (ctx) => {
            ctx.workbook.worksheets.getActiveWorksheet().onSelectionChanged.add(refreshContext);
            await ctx.sync();
        });
    } catch (e) { /* ignore */ }
}

// ============================================================================
// Context Management
// ============================================================================
async function refreshContext() {
    const useContext = document.getElementById("useContext")?.checked;
    const infoEl = document.getElementById("contextInfo");
    
    if (!useContext) {
        infoEl.textContent = "Context off";
        state.selectionData = null;
        state.selectionAddress = "";
        return;
    }
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const range = ctx.workbook.getSelectedRange();
            
            sheet.load("name");
            range.load(["address", "values", "formulas", "numberFormat", "rowCount", "columnCount"]);
            await ctx.sync();
            
            state.sheetName = sheet.name;
            state.selectionAddress = range.address;
            state.selectionData = {
                values: range.values,
                formulas: range.formulas,
                formats: range.numberFormat,
                rows: range.rowCount,
                cols: range.columnCount
            };
            
            infoEl.textContent = `${range.address} · ${range.rowCount}×${range.columnCount}`;
        });
    } catch (e) {
        infoEl.textContent = "No selection";
        state.selectionData = null;
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
    
    addMessage("user", prompt);
    state.conversationHistory.push({ role: "user", parts: [{ text: prompt }] });
    
    input.value = "";
    input.style.height = "auto";
    document.getElementById("sendBtn").disabled = true;
    
    showTyping();
    
    try {
        const response = await callAI();
        hideTyping();
        
        state.conversationHistory.push({ role: "model", parts: [{ text: response }] });
        
        // Parse actions
        const { message, actions } = parseResponse(response);
        state.pendingActions = actions;
        
        addMessage("ai", message, actions.length ? "has-action" : "");
        
        if (actions.length) {
            document.getElementById("applyBtn").disabled = false;
        }
    } catch (err) {
        hideTyping();
        addMessage("ai", "Error: " + err.message, "error");
    }
}

async function callAI() {
    const systemPrompt = getSystemPrompt();
    const contextMsg = getContextMessage();
    
    // Build conversation
    const contents = [];
    
    // Add context at start
    if (contextMsg && state.conversationHistory.length <= 2) {
        contents.push({ role: "user", parts: [{ text: contextMsg }] });
        contents.push({ role: "model", parts: [{ text: "I can see your data. What would you like me to do?" }] });
    }
    
    // Add history
    for (const msg of state.conversationHistory.slice(-CONFIG.MAX_HISTORY * 2)) {
        contents.push(msg);
    }
    
    // Update last user message with current context
    if (contents.length && state.selectionData) {
        const last = contents[contents.length - 1];
        if (last.role === "user") {
            last.parts[0].text = `[Selection: ${state.selectionAddress}]\n${last.parts[0].text}`;
        }
    }
    
    const res = await fetch(
        `${CONFIG.API_ENDPOINT}${CONFIG.GEMINI_MODEL}:generateContent?key=${state.apiKey}`,
        {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                systemInstruction: { parts: [{ text: systemPrompt }] },
                contents,
                generationConfig: { temperature: 0.3, maxOutputTokens: 4096 }
            })
        }
    );
    
    if (!res.ok) throw new Error("API request failed");
    
    const data = await res.json();
    return data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
}

function getSystemPrompt() {
    return `You are Excel Copilot, an expert Excel assistant. You help users with formulas, data analysis, formatting, and automation.

## RULES
1. Be concise and professional
2. Remember conversation context
3. When user asks to apply something to multiple cells, ALWAYS specify the full range
4. For formulas that need to be filled down, provide the formula AND specify the full target range

## ACTIONS
When you need to modify Excel, include action blocks:

<ACTION type="formula" target="E2:E10">
=SUM($B$2:B2)
</ACTION>

<ACTION type="values" target="A1:B2">
[["Name","Age"],["John",25]]
</ACTION>

<ACTION type="format" target="A1:D1">
{"bold":true,"fill":"#4472C4","fontColor":"#FFFFFF"}
</ACTION>

<ACTION type="autofill" source="E2" target="E2:E10">
</ACTION>

## ACTION TYPES
- formula: Apply formula to range (will auto-adjust relative refs)
- values: Set cell values (JSON 2D array)
- format: Apply formatting (JSON object)
- autofill: Fill formula from source to target range
- chart: Create chart (specify type: line/bar/pie/column)
- sort: Sort range (specify column and order)
- filter: Apply filter

## IMPORTANT
- Always use full range like "E2:E10" not just "E2"
- For running totals/cumulative, use autofill action
- Provide clear explanation before action blocks`;
}

function getContextMessage() {
    if (!state.selectionData) return "";
    
    const { values, rows, cols } = state.selectionData;
    const maxRows = Math.min(rows, 50);
    const maxCols = Math.min(cols, 20);
    
    let data = "";
    for (let r = 0; r < maxRows; r++) {
        const row = [];
        for (let c = 0; c < maxCols; c++) {
            row.push(values[r]?.[c] ?? "");
        }
        data += row.join("\t") + "\n";
    }
    
    return `Sheet: ${state.sheetName}
Selection: ${state.selectionAddress} (${rows} rows × ${cols} cols)

Data:
${data}`;
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
        await refreshContext();
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
            await applyFormula(ctx, range, data);
            break;
            
        case "values":
            applyValues(range, data);
            break;
            
        case "format":
            applyFormat(range, data);
            break;
            
        case "autofill":
            await applyAutofill(ctx, sheet, source, target);
            break;
            
        case "chart":
            createChart(sheet, range, data);
            break;
            
        case "sort":
            await applySort(range, data);
            break;
            
        default:
            range.values = [[data]];
    }
}

async function applyFormula(ctx, range, formula) {
    const rows = range.rowCount;
    const cols = range.columnCount;
    
    if (rows === 1 && cols === 1) {
        range.formulas = [[formula]];
        return;
    }
    
    // Parse the starting cell from range address
    const addr = range.address.split("!").pop();
    const startMatch = addr.match(/([A-Z]+)(\d+)/);
    if (!startMatch) {
        range.formulas = [[formula]];
        return;
    }
    
    const startRow = parseInt(startMatch[2]);
    
    // Build formula array with adjusted row references
    const formulas = [];
    for (let r = 0; r < rows; r++) {
        const rowFormulas = [];
        for (let c = 0; c < cols; c++) {
            let f = formula;
            if (r > 0) {
                // Adjust non-absolute row references
                f = formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (m, d1, col, d2, row) => {
                    if (d2 === "$") return m; // Absolute row
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
    try {
        fmt = JSON.parse(data);
    } catch {
        fmt = {};
    }
    
    if (fmt.bold) range.format.font.bold = true;
    if (fmt.italic) range.format.font.italic = true;
    if (fmt.fill) range.format.fill.color = fmt.fill;
    if (fmt.fontColor) range.format.font.color = fmt.fontColor;
    if (fmt.fontSize) range.format.font.size = fmt.fontSize;
    if (fmt.numberFormat) range.numberFormat = [[fmt.numberFormat]];
    if (fmt.align) range.format.horizontalAlignment = fmt.align;
    if (fmt.border) {
        range.format.borders.getItem("EdgeTop").style = "Continuous";
        range.format.borders.getItem("EdgeBottom").style = "Continuous";
        range.format.borders.getItem("EdgeLeft").style = "Continuous";
        range.format.borders.getItem("EdgeRight").style = "Continuous";
    }
}

async function applyAutofill(ctx, sheet, source, target) {
    const sourceRange = sheet.getRange(source);
    const targetRange = sheet.getRange(target);
    sourceRange.autoFill(targetRange, Excel.AutoFillType.fillDefault);
}

function createChart(sheet, dataRange, data) {
    let chartType = Excel.ChartType.columnClustered;
    const d = data.toLowerCase();
    
    if (d.includes("line")) chartType = Excel.ChartType.line;
    else if (d.includes("pie")) chartType = Excel.ChartType.pie;
    else if (d.includes("bar")) chartType = Excel.ChartType.barClustered;
    else if (d.includes("area")) chartType = Excel.ChartType.area;
    else if (d.includes("scatter")) chartType = Excel.ChartType.xyscatter;
    
    const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
    chart.setPosition("H2", "P17");
    chart.title.text = "Chart";
    chart.legend.position = Excel.ChartLegendPosition.bottom;
}

async function applySort(range, data) {
    const opts = typeof data === "string" ? { column: 0, ascending: true } : data;
    range.sort.apply([{
        key: opts.column || 0,
        ascending: opts.ascending !== false
    }]);
}

// ============================================================================
// Exports
// ============================================================================
export { handleSend, handleApply, refreshContext, clearChat };
