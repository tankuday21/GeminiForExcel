/*
 * Excel AI Copilot - With Chat Memory
 */

/* global document, Excel, Office, fetch, localStorage */

// ============================================================================
// Configuration
// ============================================================================
const CONFIG = {
    GEMINI_MODEL: "gemini-2.0-flash",
    API_ENDPOINT: "https://generativelanguage.googleapis.com/v1beta/models/",
    STORAGE_KEY: "excel_copilot_api_key",
    MAX_CONTEXT_ROWS: 100,
    MAX_CONTEXT_COLS: 26,
    MAX_HISTORY: 20 // Keep last 20 messages for context
};

// ============================================================================
// State
// ============================================================================
let state = {
    apiKey: "",
    lastResponse: null,
    pendingAction: null,
    selectionData: null,
    selectionAddress: "",
    conversationHistory: [], // For Gemini API memory
    isFirstMessage: true
};

// ============================================================================
// DOM Elements
// ============================================================================
const elements = {};

// ============================================================================
// Initialization
// ============================================================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeApp();
    }
});

function initializeApp() {
    cacheElements();
    state.apiKey = localStorage.getItem(CONFIG.STORAGE_KEY) || "";
    bindEvents();
    refreshSelection();
    setupSelectionHandler();
    setupTextareaAutoResize();
}

function cacheElements() {
    elements.welcomeScreen = document.getElementById("welcomeScreen");
    elements.chatContainer = document.getElementById("chatContainer");
    elements.promptInput = document.getElementById("promptInput");
    elements.sendBtn = document.getElementById("sendBtn");
    elements.applyBtn = document.getElementById("applyBtn");
    elements.refreshBtn = document.getElementById("refreshBtn");
    elements.useSelection = document.getElementById("useSelection");
    elements.selectionText = document.getElementById("selectionText");
    elements.settingsBtn = document.getElementById("settingsBtn");
    elements.settingsModal = document.getElementById("settingsModal");
    elements.closeSettings = document.getElementById("closeSettings");
    elements.cancelSettings = document.getElementById("cancelSettings");
    elements.saveSettings = document.getElementById("saveSettings");
    elements.apiKeyInput = document.getElementById("apiKeyInput");
    elements.togglePassword = document.getElementById("togglePassword");
    elements.toast = document.getElementById("toast");
    elements.toastMessage = document.getElementById("toastMessage");
    elements.clearChat = document.getElementById("clearChat");
}

function bindEvents() {
    elements.sendBtn?.addEventListener("click", sendMessage);
    elements.promptInput?.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });
    
    elements.promptInput?.addEventListener("input", () => {
        elements.sendBtn.disabled = !elements.promptInput.value.trim();
    });
    
    elements.applyBtn?.addEventListener("click", applyChanges);
    
    elements.refreshBtn?.addEventListener("click", () => {
        elements.refreshBtn.classList.add("spinning");
        refreshSelection().then(() => {
            setTimeout(() => elements.refreshBtn.classList.remove("spinning"), 500);
        });
    });
    
    elements.useSelection?.addEventListener("change", refreshSelection);
    
    elements.settingsBtn?.addEventListener("click", openSettings);
    elements.closeSettings?.addEventListener("click", closeSettings);
    elements.cancelSettings?.addEventListener("click", closeSettings);
    elements.saveSettings?.addEventListener("click", saveSettings);
    
    elements.clearChat?.addEventListener("click", clearConversation);
    
    elements.togglePassword?.addEventListener("click", () => {
        const input = elements.apiKeyInput;
        const icon = elements.togglePassword.querySelector("i");
        if (input.type === "password") {
            input.type = "text";
            icon.classList.replace("fa-eye", "fa-eye-slash");
        } else {
            input.type = "password";
            icon.classList.replace("fa-eye-slash", "fa-eye");
        }
    });
    
    elements.settingsModal?.addEventListener("click", (e) => {
        if (e.target === elements.settingsModal) closeSettings();
    });
    
    document.querySelectorAll(".suggestion-chip").forEach(chip => {
        chip.addEventListener("click", () => {
            elements.promptInput.value = chip.dataset.prompt;
            elements.sendBtn.disabled = false;
            sendMessage();
        });
    });
}

function setupTextareaAutoResize() {
    const textarea = elements.promptInput;
    if (!textarea) return;
    textarea.addEventListener("input", () => {
        textarea.style.height = "auto";
        textarea.style.height = Math.min(textarea.scrollHeight, 120) + "px";
    });
}

async function setupSelectionHandler() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.onSelectionChanged.add(onSelectionChanged);
            await context.sync();
        });
    } catch (err) {
        console.log("Selection handler setup skipped:", err.message);
    }
}

async function onSelectionChanged() {
    if (elements.useSelection?.checked) {
        await refreshSelection();
    }
}

// ============================================================================
// Selection Management
// ============================================================================
async function refreshSelection() {
    const useSelection = elements.useSelection?.checked;
    
    if (!useSelection) {
        elements.selectionText.textContent = "Selection disabled";
        state.selectionData = null;
        state.selectionAddress = "";
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["address", "values", "rowCount", "columnCount"]);
            await context.sync();
            
            state.selectionAddress = range.address;
            state.selectionData = range.values;
            
            elements.selectionText.textContent = `${range.address} (${range.rowCount}x${range.columnCount})`;
        });
    } catch (err) {
        elements.selectionText.textContent = "No selection";
    }
}

// ============================================================================
// Chat & Messaging
// ============================================================================
function showChat() {
    if (state.isFirstMessage) {
        elements.welcomeScreen.style.display = "none";
        elements.chatContainer.classList.add("active");
        state.isFirstMessage = false;
    }
}

function addMessage(role, content, type = "normal") {
    showChat();
    
    const msgEl = document.createElement("div");
    msgEl.className = `message ${role} ${type}`;
    
    const avatarIcon = role === "ai" ? "fa-microchip" : "fa-user";
    const statusIcon = type === "action" ? '<i class="fas fa-check-circle status-icon success"></i>' : 
                       type === "error" ? '<i class="fas fa-times-circle status-icon error"></i>' : '';
    
    msgEl.innerHTML = `
        <div class="message-avatar">
            <i class="fas ${avatarIcon}"></i>
        </div>
        <div class="message-content">
            ${statusIcon}
            <div class="message-text">${formatContent(content)}</div>
        </div>
    `;
    
    elements.chatContainer.appendChild(msgEl);
    elements.chatContainer.scrollTop = elements.chatContainer.scrollHeight;
}

function formatContent(content) {
    let formatted = content
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    
    // Remove emojis and replace with nothing
    formatted = formatted.replace(/[\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]|[\u{1F600}-\u{1F64F}]|[\u{1F680}-\u{1F6FF}]/gu, '');
    
    // Code blocks
    formatted = formatted.replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) => {
        return `<pre><code>${code.trim()}</code></pre>`;
    });
    
    // Inline code
    formatted = formatted.replace(/`([^`]+)`/g, '<code>$1</code>');
    
    // Bold
    formatted = formatted.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    
    // Line breaks
    formatted = formatted.replace(/\n/g, '<br>');
    
    return formatted;
}

function showTyping() {
    showChat();
    
    const typingEl = document.createElement("div");
    typingEl.className = "message ai";
    typingEl.id = "typingIndicator";
    typingEl.innerHTML = `
        <div class="message-avatar">
            <i class="fas fa-microchip"></i>
        </div>
        <div class="message-content">
            <div class="typing-indicator">
                <span></span><span></span><span></span>
            </div>
        </div>
    `;
    
    elements.chatContainer.appendChild(typingEl);
    elements.chatContainer.scrollTop = elements.chatContainer.scrollHeight;
}

function hideTyping() {
    document.getElementById("typingIndicator")?.remove();
}

function clearConversation() {
    state.conversationHistory = [];
    elements.chatContainer.innerHTML = "";
    elements.welcomeScreen.style.display = "flex";
    elements.chatContainer.classList.remove("active");
    state.isFirstMessage = true;
    state.pendingAction = null;
    elements.applyBtn.disabled = true;
    showToast("Conversation cleared");
}

// ============================================================================
// AI Communication with Memory
// ============================================================================
async function sendMessage() {
    const prompt = elements.promptInput.value.trim();
    if (!prompt) return;
    
    if (!state.apiKey) {
        openSettings();
        showToast("Please set your API key first", "warning");
        return;
    }
    
    // Add user message to UI
    addMessage("user", prompt);
    
    // Add to conversation history for memory
    state.conversationHistory.push({
        role: "user",
        parts: [{ text: prompt }]
    });
    
    // Trim history if too long
    if (state.conversationHistory.length > CONFIG.MAX_HISTORY * 2) {
        state.conversationHistory = state.conversationHistory.slice(-CONFIG.MAX_HISTORY * 2);
    }
    
    elements.promptInput.value = "";
    elements.promptInput.style.height = "auto";
    elements.sendBtn.disabled = true;
    
    showTyping();
    
    try {
        const response = await callGeminiAPIWithMemory();
        hideTyping();
        
        // Add AI response to history
        state.conversationHistory.push({
            role: "model",
            parts: [{ text: response }]
        });
        
        const parsed = parseAIResponse(response);
        state.lastResponse = parsed;
        
        addMessage("ai", parsed.message, parsed.hasAction ? "action" : "normal");
        
        if (parsed.hasAction) {
            elements.applyBtn.disabled = false;
            state.pendingAction = parsed.action;
        }
    } catch (err) {
        hideTyping();
        addMessage("ai", `Error: ${err.message}`, "error");
    }
}

async function callGeminiAPIWithMemory() {
    const systemInstruction = buildSystemPrompt();
    const contextInfo = buildContextPrompt();
    
    // Build contents array with conversation history
    const contents = [];
    
    // Add context as first user message if we have selection
    if (state.selectionData && state.conversationHistory.length <= 2) {
        contents.push({
            role: "user",
            parts: [{ text: `Current Excel context:\n${contextInfo}` }]
        });
        contents.push({
            role: "model", 
            parts: [{ text: "I understand. I can see your Excel data. How can I help you with it?" }]
        });
    }
    
    // Add conversation history
    contents.push(...state.conversationHistory);
    
    // If latest message needs context update, add it
    if (state.selectionData && state.conversationHistory.length > 2) {
        const lastUserMsg = contents[contents.length - 1];
        if (lastUserMsg.role === "user") {
            lastUserMsg.parts[0].text = `[Current selection: ${state.selectionAddress}]\n${lastUserMsg.parts[0].text}`;
        }
    }
    
    const url = `${CONFIG.API_ENDPOINT}${CONFIG.GEMINI_MODEL}:generateContent?key=${encodeURIComponent(state.apiKey)}`;
    
    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            systemInstruction: { parts: [{ text: systemInstruction }] },
            contents: contents,
            generationConfig: { 
                temperature: 0.7, 
                maxOutputTokens: 2048 
            }
        })
    });
    
    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`API Error: ${response.status}`);
    }
    
    const data = await response.json();
    return data?.candidates?.[0]?.content?.parts?.map(p => p.text || "").join("") || "";
}

function buildSystemPrompt() {
    return `You are Excel AI Copilot, a professional assistant for Microsoft Excel. 

IMPORTANT RULES:
1. Remember the entire conversation context. If user refers to previous messages, use that context.
2. Be concise and professional. No emojis.
3. When user asks to "drag" or "fill" a formula, apply it to all cells in the range automatically.
4. For running calculations, apply the formula to ALL cells in the target range, not just one cell.

CAPABILITIES:
- Analyze data, create formulas, generate charts, format cells
- Apply formulas to multiple cells at once
- Remember what was discussed earlier in the conversation

ACTION FORMAT (when modifying the sheet):
[ACTION]
type: formula|values|format|chart
target: cell reference (e.g., E8:E12) or "selection"
data: the formula or data to apply
[/ACTION]

EXAMPLES:
- For running sum in E8:E12, use target: E8:E12 and apply the formula to all cells
- When user says "drag it" or "fill down", apply to the full range mentioned earlier

Always provide the complete solution. If user asks to fill/drag, do it automatically.`;
}

function buildContextPrompt() {
    if (!state.selectionData || !state.selectionAddress) {
        return "No data selected.";
    }
    
    const rows = Math.min(state.selectionData.length, CONFIG.MAX_CONTEXT_ROWS);
    const cols = Math.min(state.selectionData[0]?.length || 0, CONFIG.MAX_CONTEXT_COLS);
    
    let dataStr = "";
    for (let r = 0; r < rows; r++) {
        const rowData = [];
        for (let c = 0; c < cols; c++) {
            rowData.push(state.selectionData[r][c] ?? "");
        }
        dataStr += rowData.join("\t") + "\n";
    }
    
    return `Selected: ${state.selectionAddress} (${state.selectionData.length} rows x ${state.selectionData[0]?.length || 0} cols)\n\nData:\n${dataStr}`;
}

function parseAIResponse(response) {
    const actionMatch = response.match(/\[ACTION\]([\s\S]*?)\[\/ACTION\]/);
    
    if (actionMatch) {
        const actionBlock = actionMatch[1];
        const typeMatch = actionBlock.match(/type:\s*(\w+)/);
        const targetMatch = actionBlock.match(/target:\s*([^\n]+)/);
        const dataMatch = actionBlock.match(/data:\s*([\s\S]*?)(?=\n\w+:|$)/);
        
        const message = response.replace(/\[ACTION\][\s\S]*?\[\/ACTION\]/, "").trim();
        
        return {
            message: message || "Ready to apply changes.",
            hasAction: true,
            action: {
                type: typeMatch?.[1]?.trim() || "value",
                target: targetMatch?.[1]?.trim() || "selection",
                data: dataMatch?.[1]?.trim() || ""
            }
        };
    }
    
    return { message: response, hasAction: false, action: null };
}

// ============================================================================
// Apply Changes
// ============================================================================
async function applyChanges() {
    if (!state.pendingAction) {
        showToast("No changes to apply", "error");
        return;
    }
    
    const action = state.pendingAction;
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            let targetRange;
            
            if (action.target === "selection" || !action.target) {
                targetRange = context.workbook.getSelectedRange();
            } else {
                targetRange = sheet.getRange(action.target);
            }
            
            targetRange.load(["rowCount", "columnCount", "address"]);
            await context.sync();
            
            switch (action.type) {
                case "formula":
                    await applyFormulaToRange(context, sheet, targetRange, action.data);
                    break;
                case "values":
                case "value":
                    await applyValues(targetRange, action.data);
                    break;
                case "format":
                    applyFormat(targetRange, action.data);
                    break;
                case "chart":
                    createChart(sheet, targetRange, action.data);
                    break;
                default:
                    targetRange.values = [[action.data]];
            }
            
            await context.sync();
        });
        
        addMessage("ai", "Changes applied successfully.", "action");
        showToast("Applied");
        
        state.pendingAction = null;
        elements.applyBtn.disabled = true;
        
        await refreshSelection();
    } catch (err) {
        addMessage("ai", `Failed to apply: ${err.message}`, "error");
        showToast("Failed", "error");
    }
}

async function applyFormulaToRange(context, sheet, targetRange, formula) {
    const rowCount = targetRange.rowCount;
    const colCount = targetRange.columnCount;
    
    // If single cell, just apply formula
    if (rowCount === 1 && colCount === 1) {
        targetRange.formulas = [[formula]];
        return;
    }
    
    // For multiple cells, we need to apply formula with relative references
    // Get the starting cell address
    const address = targetRange.address;
    const match = address.match(/([A-Z]+)(\d+)/);
    if (!match) {
        targetRange.formulas = [[formula]];
        return;
    }
    
    const startCol = match[1];
    const startRow = parseInt(match[2]);
    
    // Create formula array for all cells
    const formulas = [];
    for (let r = 0; r < rowCount; r++) {
        const rowFormulas = [];
        for (let c = 0; c < colCount; c++) {
            // Adjust formula for each row
            let adjustedFormula = formula;
            
            // Replace row numbers in formula (simple adjustment)
            if (r > 0) {
                adjustedFormula = formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, dollarCol, col, dollarRow, row) => {
                    if (dollarRow === '$') {
                        return match; // Absolute row reference, don't change
                    }
                    const newRow = parseInt(row) + r;
                    return `${dollarCol}${col}${dollarRow}${newRow}`;
                });
            }
            rowFormulas.push(adjustedFormula);
        }
        formulas.push(rowFormulas);
    }
    
    targetRange.formulas = formulas;
}

async function applyValues(targetRange, data) {
    let values;
    try {
        values = JSON.parse(data);
        if (!Array.isArray(values)) values = [[values]];
        else if (!Array.isArray(values[0])) values = [values];
    } catch {
        values = [[data]];
    }
    
    if (values.length > 1 || values[0]?.length > 1) {
        targetRange = targetRange.getCell(0, 0).getResizedRange(values.length - 1, values[0].length - 1);
    }
    targetRange.values = values;
}

function applyFormat(range, formatData) {
    const data = formatData.toLowerCase();
    if (data.includes("bold")) range.format.font.bold = true;
    if (data.includes("italic")) range.format.font.italic = true;
    if (data.includes("currency")) range.numberFormat = [["$#,##0.00"]];
    if (data.includes("percent")) range.numberFormat = [["0.00%"]];
    if (data.includes("header")) {
        range.format.font.bold = true;
        range.format.fill.color = "#1a1a2e";
        range.format.font.color = "#FFFFFF";
    }
}

function createChart(sheet, dataRange, chartData) {
    const data = chartData.toLowerCase();
    let chartType = Excel.ChartType.columnClustered;
    
    if (data.includes("line")) chartType = Excel.ChartType.line;
    else if (data.includes("pie")) chartType = Excel.ChartType.pie;
    else if (data.includes("bar")) chartType = Excel.ChartType.barClustered;
    else if (data.includes("area")) chartType = Excel.ChartType.area;
    
    const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
    chart.setPosition("H2", "O15");
    chart.title.text = "Chart";
}

// ============================================================================
// Settings
// ============================================================================
function openSettings() {
    elements.apiKeyInput.value = state.apiKey;
    elements.settingsModal.classList.add("active");
}

function closeSettings() {
    elements.settingsModal.classList.remove("active");
}

function saveSettings() {
    const key = elements.apiKeyInput.value.trim();
    state.apiKey = key;
    localStorage.setItem(CONFIG.STORAGE_KEY, key);
    closeSettings();
    showToast("Saved");
}

// ============================================================================
// Toast
// ============================================================================
function showToast(message, type = "success") {
    const icon = elements.toast.querySelector("i");
    icon.className = type === "success" ? "fas fa-check-circle" : "fas fa-exclamation-triangle";
    
    elements.toastMessage.textContent = message;
    elements.toast.classList.add("show");
    elements.toast.classList.toggle("error", type !== "success");
    
    setTimeout(() => elements.toast.classList.remove("show"), 2500);
}

// ============================================================================
// Exports
// ============================================================================
export { sendMessage, applyChanges, refreshSelection, clearConversation };
