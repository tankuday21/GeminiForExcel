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

1. **ALWAYS USE THE CORRECT COLUMN LETTERS** - Check the COLUMN STRUCTURE table carefully
2. **VERIFY column names match column letters** - If user says "State column", find which letter has "State" header
3. **Use exact cell references** - e.g., E2:E100, not generic references
4. **Data starts at row 2** (row 1 is headers) unless otherwise shown

## BEFORE ANY ACTION

1. Identify the correct column letter by matching the header name
2. Determine the data range (usually row 2 to last row)
3. Double-check your formula references

## ACTION FORMAT

Use XML-style action blocks:

**For Data Validation (Dropdown):**
<ACTION type="validation" target="K10" source="E2:E100" validationType="list">
</ACTION>

**For Formulas:**
<ACTION type="formula" target="F2:F100">
=E2*D2
</ACTION>

**For UNIQUE formula (spill):**
<ACTION type="formula" target="L1">
=UNIQUE(E2:E100)
</ACTION>

**For Values:**
<ACTION type="values" target="A1:B2">
[["Header1","Header2"],["Value1","Value2"]]
</ACTION>

**For Formatting:**
<ACTION type="format" target="A1:F1">
{"bold":true,"fill":"#4472C4","fontColor":"#FFFFFF"}
</ACTION>

**For Charts:**
<ACTION type="chart" target="A1:D20" chartType="column">
</ACTION>

## VALIDATION DROPDOWN EXAMPLE

If user wants dropdown of "State" values in cell K10, and State is in column E:
1. Find State column = E
2. Data range = E2:E[lastrow]
3. Create validation:

<ACTION type="validation" target="K10" source="E2:E100" validationType="list">
</ACTION>

## COMMON MISTAKES TO AVOID

- Using wrong column letter (e.g., C instead of E for State)
- Forgetting to check the COLUMN STRUCTURE table
- Using row 1 in data ranges (row 1 is headers)
- Not specifying the full data range

Always explain what you're doing and which columns you're using.`;
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
        const validationType = attrs.match(/validationType="([^"]+)"/)?.[1] || "";
        const chartType = attrs.match(/chartType="([^"]+)"/)?.[1] || "column";
        
        actions.push({ type, target, source, validationType, chartType, data: content });
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
    
    let successCount = 0;
    let errorMsg = "";
    
    try {
        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            
            for (const action of state.pendingActions) {
                try {
                    await executeAction(ctx, sheet, action);
                    await ctx.sync();
                    successCount++;
                } catch (e) {
                    errorMsg = e.message;
                    console.error("Action failed:", e);
                }
            }
        });
        
        if (successCount === state.pendingActions.length) {
            addMessage("ai", "All changes applied successfully.", "success");
            toast("Applied");
        } else if (successCount > 0) {
            addMessage("ai", `${successCount}/${state.pendingActions.length} changes applied. Error: ${errorMsg}`, "error");
        } else {
            addMessage("ai", `Failed: ${errorMsg}`, "error");
        }
        
        state.pendingActions = [];
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
            await applyValidation(ctx, range, source, validationType);
            break;
            
        case "chart":
            createChart(sheet, range, chartType);
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

async function applyValidation(ctx, range, source, validationType) {
    if (validationType === "list" && source) {
        range.dataValidation.rule = {
            list: {
                inCellDropDown: true,
                source: source
            }
        };
    }
}

function createChart(sheet, dataRange, chartType) {
    let type = Excel.ChartType.columnClustered;
    const ct = (chartType || "").toLowerCase();
    
    if (ct.includes("line")) type = Excel.ChartType.line;
    else if (ct.includes("pie")) type = Excel.ChartType.pie;
    else if (ct.includes("bar")) type = Excel.ChartType.barClustered;
    else if (ct.includes("area")) type = Excel.ChartType.area;
    else if (ct.includes("scatter")) type = Excel.ChartType.xyscatter;
    
    const chart = sheet.charts.add(type, dataRange, Excel.ChartSeriesBy.auto);
    chart.setPosition("H2", "P17");
    chart.title.text = "Chart";
}

function applySort(range, data) {
    const opts = typeof data === "string" ? JSON.parse(data) : (data || {});
    range.sort.apply([{ key: opts.column || 0, ascending: opts.ascending !== false }]);
}

export { handleSend, handleApply, readExcelData, clearChat };
