/*
 * Excel AI Copilot - Modern UI
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
    MAX_CONTEXT_COLS: 26
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
    chatHistory: [],
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
    // Cache DOM elements
    cacheElements();
    
    // Load saved API key
    state.apiKey = localStorage.getItem(CONFIG.STORAGE_KEY) || "";
    
    // Bind events
    bindEvents();
    
    // Initial selection refresh
    refreshSelection();
    
    // Setup selection change handler
    setupSelectionHandler();
    
    // Auto-resize textarea
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
    elements.contextInfo = document.getElementById("contextInfo");
    elements.settingsBtn = document.getElementById("settingsBtn");
    elements.settingsModal = document.getElementById("settingsModal");
    elements.closeSettings = document.getElementById("closeSettings");
    elements.cancelSettings = document.getElementById("cancelSettings");
    elements.saveSettings = document.getElementById("saveSettings");
    elements.apiKeyInput = document.getElementById("apiKeyInput");
    elements.togglePassword = document.getElementById("togglePassword");
    elements.toast = document.getElementById("toast");
    elements.toastMessage = document.getElementById("toastMessage");
}

function bindEvents() {
    // Send message
    elements.sendBtn?.addEventListener("click", sendMessage);
    elements.promptInput?.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            sendMessage();
        }
    });
    
    // Enable/disable send button based on input
    elements.promptInput?.addEventListener("input", () => {
        elements.sendBtn.disabled = !elements.promptInput.value.trim();
    });
    
    // Apply changes
    elements.applyBtn?.addEventListener("click", applyChanges);
    
    // Refresh selection
    elements.refreshBtn?.addEventListener("click", () => {
        elements.refreshBtn.classList.add("spinning");
        refreshSelection().then(() => {
            setTimeout(() => elements.refreshBtn.classList.remove("spinning"), 500);
        });
    });
    
    // Selection toggle
    elements.useSelection?.addEventListener("change", refreshSelection);
    
    // Settings modal
    elements.settingsBtn?.addEventListener("click", openSettings);
    elements.closeSettings?.addEventListener("click", closeSettings);
    elements.cancelSettings?.addEventListener("click", closeSettings);
    elements.saveSettings?.addEventListener("click", saveSettings);
    
    // Toggle password visibility
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
    
    // Close modal on overlay click
    elements.settingsModal?.addEventListener("click", (e) => {
        if (e.target === elements.settingsModal) closeSettings();
    });
    
    // Suggestion chips
    document.querySelectorAll(".suggestion-chip").forEach(chip => {
        chip.addEventListener("click", () => {
            const prompt = chip.dataset.prompt;
            elements.promptInput.value = prompt;
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
            
            const rows = range.rowCount;
            const cols = range.columnCount;
            elements.selectionText.textContent = `${range.address} (${rows}×${cols})`;
        });
    } catch (err) {
        elements.selectionText.textContent = "No selection";
        console.error("Selection error:", err);
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
    
    const avatarIcon = role === "ai" ? "fa-sparkles" : "fa-user";
    
    msgEl.innerHTML = `
        <div class="message-avatar">
            <i class="fas ${avatarIcon}"></i>
        </div>
        <div class="message-content">${formatContent(content)}</div>
    `;
    
    elements.chatContainer.appendChild(msgEl);
    elements.chatContainer.scrollTop = elements.chatContainer.scrollHeight;
    
    state.chatHistory.push({ role, content, type });
}

function formatContent(content) {
    let formatted = content
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    
    // Code blocks
    formatted = formatted.replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) => {
        return `<pre>${code.trim()}</pre>`;
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
            <i class="fas fa-sparkles"></i>
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

// ============================================================================
// AI Communication
// ============================================================================
async function sendMessage() {
    const prompt = elements.promptInput.value.trim();
    if (!prompt) return;
    
    // Check API key
    if (!state.apiKey) {
        openSettings();
        showToast("Please set your API key first", "warning");
        return;
    }
    
    // Add user message
    addMessage("user", prompt);
    elements.promptInput.value = "";
    elements.promptInput.style.height = "auto";
    elements.sendBtn.disabled = true;
    
    // Show typing
    showTyping();
    
    try {
        const response = await callGeminiAPI(prompt);
        hideTyping();
        
        const parsed = parseAIResponse(response);
        state.lastResponse = parsed;
        
        addMessage("ai", parsed.message, parsed.hasAction ? "action" : "normal");
        
        if (parsed.hasAction) {
            elements.applyBtn.disabled = false;
            state.pendingAction = parsed.action;
        }
    } catch (err) {
        hideTyping();
        addMessage("ai", `Sorry, I encountered an error: ${err.message}`, "error");
    }
}

async function callGeminiAPI(userPrompt) {
    const systemPrompt = buildSystemPrompt();
    const contextPrompt = buildContextPrompt();
    const fullPrompt = `${systemPrompt}\n\n${contextPrompt}\n\nUser: ${userPrompt}`;
    
    const url = `${CONFIG.API_ENDPOINT}${CONFIG.GEMINI_MODEL}:generateContent?key=${encodeURIComponent(state.apiKey)}`;
    
    const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            contents: [{ parts: [{ text: fullPrompt }] }],
            generationConfig: { temperature: 0.7, maxOutputTokens: 2048 }
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
    return `You are Excel AI Copilot, a helpful assistant for Microsoft Excel. You help users analyze data, create formulas, generate charts, and automate tasks.

Be concise and friendly. When providing formulas, use exact Excel syntax.

When you want to modify the sheet, include an ACTION block:
[ACTION]
type: formula|values|format|chart
target: cell reference or "selection"
data: the data to apply
[/ACTION]

Only include ACTION blocks when the user wants to modify the sheet.`;
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
    
    return `Selected: ${state.selectionAddress} (${state.selectionData.length} rows × ${state.selectionData[0]?.length || 0} cols)\n\nData:\n${dataStr}`;
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
            message: message || "Ready to apply changes. Click 'Apply to Sheet' when ready.",
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
            
            switch (action.type) {
                case "formula":
                    targetRange.formulas = [[action.data]];
                    break;
                case "values":
                case "value":
                    let values;
                    try {
                        values = JSON.parse(action.data);
                        if (!Array.isArray(values)) values = [[values]];
                        else if (!Array.isArray(values[0])) values = [values];
                    } catch {
                        values = [[action.data]];
                    }
                    if (values.length > 1 || values[0]?.length > 1) {
                        targetRange = targetRange.getCell(0, 0).getResizedRange(values.length - 1, values[0].length - 1);
                    }
                    targetRange.values = values;
                    break;
                case "format":
                    applyFormat(targetRange, action.data);
                    break;
                case "chart":
                    createChart(context, sheet, targetRange, action.data);
                    break;
                default:
                    targetRange.values = [[action.data]];
            }
            
            await context.sync();
        });
        
        addMessage("ai", "✅ Changes applied successfully!", "action");
        showToast("Changes applied!");
        
        state.pendingAction = null;
        elements.applyBtn.disabled = true;
        
        await refreshSelection();
    } catch (err) {
        addMessage("ai", `❌ Failed to apply: ${err.message}`, "error");
        showToast("Failed to apply changes", "error");
    }
}

function applyFormat(range, formatData) {
    const data = formatData.toLowerCase();
    if (data.includes("bold")) range.format.font.bold = true;
    if (data.includes("italic")) range.format.font.italic = true;
    if (data.includes("currency")) range.numberFormat = [["$#,##0.00"]];
    if (data.includes("percent")) range.numberFormat = [["0.00%"]];
    if (data.includes("header")) {
        range.format.font.bold = true;
        range.format.fill.color = "#7C3AED";
        range.format.font.color = "#FFFFFF";
    }
}

function createChart(context, sheet, dataRange, chartData) {
    const data = chartData.toLowerCase();
    let chartType = Excel.ChartType.columnClustered;
    
    if (data.includes("line")) chartType = Excel.ChartType.line;
    else if (data.includes("pie")) chartType = Excel.ChartType.pie;
    else if (data.includes("bar")) chartType = Excel.ChartType.barClustered;
    else if (data.includes("area")) chartType = Excel.ChartType.area;
    
    const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
    chart.setPosition("H2", "O15");
    chart.title.text = "Generated Chart";
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
    showToast("Settings saved!");
}

// ============================================================================
// Toast
// ============================================================================
function showToast(message, type = "success") {
    const icon = elements.toast.querySelector("i");
    icon.className = type === "success" ? "fas fa-check-circle" : "fas fa-exclamation-circle";
    icon.style.color = type === "success" ? "var(--success)" : "var(--error)";
    
    elements.toastMessage.textContent = message;
    elements.toast.classList.add("show");
    
    setTimeout(() => {
        elements.toast.classList.remove("show");
    }, 3000);
}

// ============================================================================
// Exports
// ============================================================================
export { sendMessage, applyChanges, refreshSelection };
