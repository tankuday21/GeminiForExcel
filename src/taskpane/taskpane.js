/*
 * Excel AI Copilot - Accurate Data Understanding
 * Enhanced with: Task-specific prompts, Function calling, RAG, Multi-step reasoning, Learning
 */

/* global document, Excel, Office, fetch, localStorage */

// Version number - increment with each update
const VERSION = "2.6.2";

import {
    detectTaskType,
    TASK_TYPES,
    enhancePrompt,
    isCorrection,
    handleCorrection,
    processResponse,
    getRAGContext,
    getCorrectionContext,
    clearCorrections
} from "./ai-engine.js";

const CONFIG = {
    GEMINI_MODEL: "gemini-2.0-flash",
    API_ENDPOINT: "https://generativelanguage.googleapis.com/v1beta/models/",
    STORAGE_KEY: "excel_copilot_api_key",
    THEME_KEY: "excel_copilot_theme",
    MAX_HISTORY: 10,
    MAX_RETRIES: 3,
    RETRY_DELAY: 1000,
    VERSION: VERSION
};

const state = {
    apiKey: "",
    pendingActions: [],
    currentData: null,
    conversationHistory: [],
    isFirstMessage: true,
    lastAIResponse: "",      // Track last AI response for corrections
    currentTaskType: null,   // Track current task type
    mode: "edit",            // "edit" or "readonly"
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
    
    // Update version badge and add click handler
    const versionBadge = document.getElementById("versionBadge");
    if (versionBadge) {
        versionBadge.textContent = `v${VERSION}`;
        versionBadge.style.cursor = "pointer";
        versionBadge.addEventListener("click", checkForUpdates);
    }
    
    // Load saved theme
    const savedTheme = localStorage.getItem(CONFIG.THEME_KEY);
    if (savedTheme) {
        document.documentElement.setAttribute('data-theme', savedTheme);
    }
    // Load saved mode preference
    const savedMode = localStorage.getItem("excel_copilot_mode");
    if (savedMode) {
        state.mode = savedMode;
    }
    bindEvents();
    initModeButtons();
    readExcelData().then(() => {
        // Smart suggestions disabled for cleaner UI
    });
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
        // Set mode radio button
        document.getElementById(state.mode === "readonly" ? "modeReadOnly" : "modeEdit").checked = true;
        document.getElementById("modal").classList.add("open");
    });
    
    document.getElementById("closeModal")?.addEventListener("click", closeModal);
    document.getElementById("cancelBtn")?.addEventListener("click", closeModal);
    document.getElementById("saveBtn")?.addEventListener("click", () => {
        state.apiKey = document.getElementById("apiKeyInput").value.trim();
        localStorage.setItem(CONFIG.STORAGE_KEY, state.apiKey);
        
        // Save mode preference
        const selectedMode = document.querySelector('input[name="mode"]:checked')?.value || "edit";
        state.mode = selectedMode;
        localStorage.setItem("excel_copilot_mode", selectedMode);
        
        // Update apply button state
        updateApplyButtonState();
        
        closeModal();
        toast("Saved");
    });
    
    document.getElementById("modal")?.addEventListener("click", (e) => {
        if (e.target.id === "modal") closeModal();
    });
    
    // Clear learned preferences button
    document.getElementById("clearPrefsBtn")?.addEventListener("click", () => {
        clearLearnedCorrections();
    });
    
    document.getElementById("clearBtn")?.addEventListener("click", clearChat);
    
    // History and Undo buttons
    document.getElementById("historyBtn")?.addEventListener("click", toggleHistoryPanel);
    document.getElementById("undoBtn")?.addEventListener("click", performUndo);
    
    // Theme toggle
    document.getElementById("themeBtn")?.addEventListener("click", toggleTheme);
    
    // Mode switch buttons
    document.getElementById("editModeBtn")?.addEventListener("click", () => setMode("edit"));
    document.getElementById("readOnlyModeBtn")?.addEventListener("click", () => setMode("readonly"));
    
    // Keyboard shortcuts
    document.addEventListener("keydown", handleKeyboardShortcuts);
    
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
            
            // Smart suggestions disabled for cleaner UI
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
    
    // Smooth scroll to bottom
    setTimeout(() => {
        chat.scrollTo({
            top: chat.scrollHeight,
            behavior: 'smooth'
        });
    }, 100);
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

// ============================================================================
// Loading Skeleton
// ============================================================================

function showLoadingSkeleton() {
    showChat();
    const chat = document.getElementById("chat");
    const el = document.createElement("div");
    el.className = "skeleton-msg";
    el.id = "loading-skeleton";
    el.innerHTML = `
        <div class="skeleton skeleton-avatar"></div>
        <div class="skeleton-content">
            <div class="skeleton skeleton-line" style="width: 90%"></div>
            <div class="skeleton skeleton-line" style="width: 75%"></div>
            <div class="skeleton skeleton-line" style="width: 60%"></div>
        </div>
    `;
    chat.appendChild(el);
    chat.scrollTop = chat.scrollHeight;
}

function hideLoadingSkeleton() {
    document.getElementById("loading-skeleton")?.remove();
}

// ============================================================================
// Smart Suggestions
// ============================================================================

function generateSmartSuggestions() {
    const suggestions = [];
    const data = state.currentData;
    
    if (!data) return suggestions;
    
    // Based on data characteristics
    if (data.rowCount > 5) {
        suggestions.push({ icon: '📊', text: 'Summarize this data', prompt: 'Summarize this data and give me key statistics' });
    }
    
    // Check for numeric columns
    const hasNumbers = data.values.some(row => row.some(cell => typeof cell === 'number'));
    if (hasNumbers) {
        suggestions.push({ icon: '➕', text: 'Add totals', prompt: 'Add SUM formulas for all numeric columns' });
        suggestions.push({ icon: '📈', text: 'Create chart', prompt: 'Create a chart to visualize this data' });
    }
    
    // Check for potential duplicates
    if (data.rowCount > 10) {
        suggestions.push({ icon: '🔍', text: 'Find duplicates', prompt: 'Check for duplicate values in this data' });
    }
    
    // Check column headers for common patterns
    const headers = data.headers.map(h => String(h).toLowerCase());
    if (headers.some(h => h.includes('date') || h.includes('time'))) {
        suggestions.push({ icon: '📅', text: 'Sort by date', prompt: 'Sort this data by date column' });
    }
    if (headers.some(h => h.includes('price') || h.includes('amount') || h.includes('cost'))) {
        suggestions.push({ icon: '💰', text: 'Calculate totals', prompt: 'Calculate the total of all monetary values' });
    }
    if (headers.some(h => h.includes('email'))) {
        suggestions.push({ icon: '✉️', text: 'Validate emails', prompt: 'Check if all email addresses are valid' });
    }
    
    return suggestions.slice(0, 4); // Max 4 suggestions
}

function showSmartSuggestions() {
    const container = document.getElementById("smartSuggestions");
    if (!container) return;
    
    const suggestions = generateSmartSuggestions();
    
    if (suggestions.length === 0) {
        container.style.display = "none";
        return;
    }
    
    container.innerHTML = suggestions.map(s => `
        <button class="smart-suggestion" data-prompt="${s.prompt}">
            <span>${s.icon}</span> ${s.text}
        </button>
    `).join('');
    
    container.style.display = "flex";
    
    // Bind click events
    container.querySelectorAll(".smart-suggestion").forEach(btn => {
        btn.addEventListener("click", () => {
            document.getElementById("promptInput").value = btn.dataset.prompt;
            document.getElementById("sendBtn").disabled = false;
            container.style.display = "none";
            handleSend();
        });
    });
}

function hideSmartSuggestions() {
    const container = document.getElementById("smartSuggestions");
    if (container) container.style.display = "none";
}

// ============================================================================
// Formula Explanation
// ============================================================================

function explainFormula(formula) {
    const explanations = {
        'SUM': 'Adds all numbers in a range',
        'AVERAGE': 'Calculates the average of numbers',
        'COUNT': 'Counts cells with numbers',
        'COUNTA': 'Counts non-empty cells',
        'MAX': 'Returns the largest value',
        'MIN': 'Returns the smallest value',
        'IF': 'Returns one value if true, another if false',
        'VLOOKUP': 'Looks up a value in the first column and returns a value in the same row',
        'XLOOKUP': 'Searches a range and returns a matching item',
        'INDEX': 'Returns a value at a given position',
        'MATCH': 'Returns the position of a value in a range',
        'CONCATENATE': 'Joins text strings together',
        'LEFT': 'Returns characters from the start of text',
        'RIGHT': 'Returns characters from the end of text',
        'MID': 'Returns characters from the middle of text',
        'LEN': 'Returns the length of text',
        'TRIM': 'Removes extra spaces from text',
        'UPPER': 'Converts text to uppercase',
        'LOWER': 'Converts text to lowercase',
        'ROUND': 'Rounds a number to specified digits',
        'SUMIF': 'Adds cells that meet a condition',
        'COUNTIF': 'Counts cells that meet a condition',
        'IFERROR': 'Returns a value if there is an error',
        'TODAY': 'Returns the current date',
        'NOW': 'Returns the current date and time',
        'YEAR': 'Returns the year from a date',
        'MONTH': 'Returns the month from a date',
        'DAY': 'Returns the day from a date'
    };
    
    // Extract function name from formula
    const match = formula.match(/=?([A-Z]+)\(/i);
    if (match) {
        const funcName = match[1].toUpperCase();
        return explanations[funcName] || `${funcName} function`;
    }
    return 'Excel formula';
}

function clearChat() {
    state.conversationHistory = [];
    state.pendingActions = [];
    state.isFirstMessage = true;
    state.lastAIResponse = "";
    state.currentTaskType = null;
    state.preview.selections = [];
    state.preview.expandedIndex = -1;
    document.getElementById("chat").innerHTML = "";
    document.getElementById("chat").style.display = "none";
    document.getElementById("welcome").style.display = "flex";
    document.getElementById("applyBtn").disabled = true;
    hidePreviewPanel();
    hideTaskTypeIndicator();
    toast("Cleared");
}

/**
 * Clears learned corrections (accessible via settings or command)
 */
function clearLearnedCorrections() {
    clearCorrections();
    toast("Learned preferences cleared");
}

function toast(msg) {
    const t = document.getElementById("toast");
    t.textContent = msg;
    t.classList.add("show");
    setTimeout(() => t.classList.remove("show"), 2000);
}

/**
 * Checks for updates by fetching the deployed version
 */
async function checkForUpdates() {
    const versionBadge = document.getElementById("versionBadge");
    const originalText = versionBadge.textContent;
    
    try {
        versionBadge.textContent = "Checking...";
        
        // Fetch the deployed taskpane.js with cache-busting
        const response = await fetch(`https://tankuday21.github.io/GeminiForExcel/taskpane.js?t=${Date.now()}`);
        const code = await response.text();
        
        // Extract version from the deployed code
        const versionMatch = code.match(/const VERSION = "([^"]+)"/);
        
        if (versionMatch) {
            const deployedVersion = versionMatch[1];
            const currentVersion = VERSION;
            
            if (deployedVersion === currentVersion) {
                toast("✓ You're on the latest version!");
                versionBadge.textContent = originalText;
            } else {
                toast(`Update available: v${deployedVersion}`);
                versionBadge.textContent = `v${currentVersion} → v${deployedVersion}`;
                versionBadge.style.color = "#ff9800";
                
                // Reset after 5 seconds
                setTimeout(() => {
                    versionBadge.textContent = originalText;
                    versionBadge.style.color = "";
                }, 5000);
            }
        } else {
            throw new Error("Could not parse version");
        }
    } catch (error) {
        console.error("Update check failed:", error);
        toast("Failed to check for updates");
        versionBadge.textContent = originalText;
    }
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
    
    // Always refresh data before sending to ensure latest data
    await readExcelData();
    
    // Clear conversation history in read-only mode to ensure fresh data context
    if (state.mode === "readonly") {
        state.conversationHistory = [];
    }
    
    // Detect task type and show indicator
    const taskType = detectTaskType(prompt);
    const isCorrectionMsg = isCorrection(prompt);
    
    // Add user message with task type badge
    addMessage("user", prompt, isCorrectionMsg ? "correction" : "");
    
    // Show task type indicator
    if (!isCorrectionMsg) {
        showTaskTypeIndicator(taskType);
    }
    
    input.value = "";
    input.style.height = "auto";
    document.getElementById("sendBtn").disabled = true;
    
    showTyping();
    showLoadingSkeleton();
    
    try {
        const response = await callAI(prompt);
        hideTyping();
        hideLoadingSkeleton();
        hideTaskTypeIndicator();
        
        // In read-only mode, don't parse actions - just show the response
        let { message, actions } = parseResponse(response);
        
        if (state.mode === "readonly") {
            // Strip out ACTION tags in read-only mode
            actions = [];
            message = response.replace(/<ACTION[\s\S]*?<\/ACTION>/g, "").trim();
        }
        
        state.pendingActions = actions;
        
        // Add task type badge to AI response
        const taskBadge = getTaskTypeBadge(state.currentTaskType);
        const enhancedMessage = taskBadge + message;
        
        addMessage("ai", enhancedMessage, actions.length ? "has-action" : "");
        
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
        hideLoadingSkeleton();
        hideTaskTypeIndicator();
        addMessage("ai", getErrorMessage(err), "error");
    }
}

/**
 * Shows task type indicator during processing
 */
function showTaskTypeIndicator(taskType) {
    const labels = {
        [TASK_TYPES.FORMULA]: "🔢 Formula Mode",
        [TASK_TYPES.CHART]: "📊 Chart Mode",
        [TASK_TYPES.ANALYSIS]: "📈 Analysis Mode",
        [TASK_TYPES.FORMAT]: "🎨 Format Mode",
        [TASK_TYPES.DATA_ENTRY]: "✏️ Data Entry Mode",
        [TASK_TYPES.VALIDATION]: "✅ Validation Mode",
        [TASK_TYPES.GENERAL]: "💡 General Mode"
    };
    
    const indicator = document.getElementById("taskTypeIndicator");
    if (indicator) {
        indicator.textContent = labels[taskType] || labels[TASK_TYPES.GENERAL];
        indicator.className = `task-indicator ${taskType}`;
        indicator.style.display = "block";
    }
}

/**
 * Hides task type indicator
 */
function hideTaskTypeIndicator() {
    const indicator = document.getElementById("taskTypeIndicator");
    if (indicator) {
        indicator.style.display = "none";
    }
}

/**
 * Gets a badge string for the task type
 */
function getTaskTypeBadge(taskType) {
    const badges = {
        [TASK_TYPES.FORMULA]: "**[Formula]** ",
        [TASK_TYPES.CHART]: "**[Chart]** ",
        [TASK_TYPES.ANALYSIS]: "**[Analysis]** ",
        [TASK_TYPES.FORMAT]: "**[Format]** ",
        [TASK_TYPES.DATA_ENTRY]: "**[Data]** ",
        [TASK_TYPES.VALIDATION]: "**[Validation]** ",
        [TASK_TYPES.GENERAL]: ""
    };
    return badges[taskType] || "";
}

async function callAI(userPrompt) {
    const dataContext = buildDataContext();
    
    // Use AI engine to enhance the prompt with task-specific context
    const enhanced = enhancePrompt(userPrompt, state.currentData);
    state.currentTaskType = enhanced.taskType;
    
    // Handle corrections - learn from user feedback
    if (enhanced.isCorrection && state.lastAIResponse) {
        handleCorrection(userPrompt, state.lastAIResponse);
    }
    
    // Build the enhanced system prompt - modify for read-only mode
    let systemPrompt = enhanced.systemPrompt;
    
    if (state.mode === "readonly") {
        systemPrompt = getReadOnlySystemPrompt();
    }
    
    // Build the enhanced user message
    const fullUserMessage = `${dataContext}\n\n---\nUSER REQUEST: ${enhanced.userPrompt}`;
    
    const contents = [...state.conversationHistory];
    contents.push({ role: "user", parts: [{ text: fullUserMessage }] });
    
    // Use retry logic for transient errors
    const response = await withRetry(async () => {
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
    });
    
    // Store response for potential correction learning
    state.lastAIResponse = response;
    
    // Process response for any function calls
    const processed = processResponse(response);
    
    return processed.response;
}

/**
 * Gets the system prompt for read-only mode
 * In this mode, AI analyzes data and gives direct answers without formulas/actions
 */
function getReadOnlySystemPrompt() {
    return `You are Excel Copilot in READ-ONLY mode. You are a data analyst assistant.

## YOUR ROLE
- Analyze the Excel data provided and give DIRECT ANSWERS
- Do NOT provide formulas or ACTION tags
- Do NOT suggest modifications to the spreadsheet
- Calculate and compute answers yourself from the data provided
- Give clear, concise answers with the actual values/numbers

## CRITICAL ACCURACY RULES
1. **COUNT CAREFULLY**: When counting occurrences, examine EVERY cell in the data preview
2. **CASE SENSITIVITY**: "d" and "D" are different - count only exact matches unless told otherwise
3. **EMPTY CELLS**: Ignore empty/null cells in counts
4. **VERIFY YOUR COUNT**: Double-check your answer before responding
5. **BE PRECISE**: If you count 5 occurrences, say exactly 5, not approximately

## EXAMPLES
- "How many times does 'd' appear?" → Examine each cell, count lowercase 'd' only → "There are exactly 5 occurrences of 'd'"
- "What is the total of column B?" → Add all numbers in column B → "The total is 1,234"
- "How many rows have value X?" → Count rows where value equals X → "There are 15 rows with value X"

## IMPORTANT
- You have access to ALL the data in the DATA PREVIEW section
- Count EVERY occurrence manually from the data shown
- Be EXACT with your counts - accuracy is critical
- Do NOT use ACTION tags - just provide text answers`;
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

// Note: getSystemPrompt is now handled by ai-engine.js via enhancePrompt()
// This function is kept as a fallback but is no longer the primary source
function getSystemPrompt() {
    return `You are Excel Copilot, an expert Excel assistant.
Check COLUMN STRUCTURE first. Data starts at row 2.
Use ACTION tags for: formula, values, format, chart, validation.`;
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
    const isReadOnly = state.mode === "readonly";
    applyBtn.disabled = !hasSelected || isReadOnly;
    
    // Update button text to indicate read-only mode
    if (isReadOnly) {
        applyBtn.textContent = "Read-Only Mode";
    } else {
        applyBtn.textContent = "Apply Changes";
    }
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
// Theme Toggle
// ============================================================================

/**
 * Toggles between light and dark theme
 */
function toggleTheme() {
    const html = document.documentElement;
    const currentTheme = html.getAttribute('data-theme');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    html.setAttribute('data-theme', newTheme);
    localStorage.setItem(CONFIG.THEME_KEY, newTheme);
    toast(newTheme === 'dark' ? 'Dark mode' : 'Light mode');
}

/**
 * Sets the mode (edit or readonly)
 */
function setMode(mode) {
    state.mode = mode;
    localStorage.setItem("excel_copilot_mode", mode);
    
    // Update button states
    const editBtn = document.getElementById("editModeBtn");
    const readOnlyBtn = document.getElementById("readOnlyModeBtn");
    
    if (editBtn && readOnlyBtn) {
        editBtn.classList.toggle("active", mode === "edit");
        readOnlyBtn.classList.toggle("active", mode === "readonly");
    }
    
    // Update apply button
    updateApplyButtonState();
    
    toast(mode === "edit" ? "Edit mode" : "Read-only mode");
}

/**
 * Initializes mode buttons based on saved state
 */
function initModeButtons() {
    const editBtn = document.getElementById("editModeBtn");
    const readOnlyBtn = document.getElementById("readOnlyModeBtn");
    
    if (editBtn && readOnlyBtn) {
        editBtn.classList.toggle("active", state.mode === "edit");
        readOnlyBtn.classList.toggle("active", state.mode === "readonly");
    }
}

// ============================================================================
// Keyboard Shortcuts
// ============================================================================

/**
 * Handles keyboard shortcuts
 */
function handleKeyboardShortcuts(e) {
    // Ctrl+Enter or Cmd+Enter to send message
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        const input = document.getElementById("promptInput");
        if (document.activeElement === input && input.value.trim()) {
            e.preventDefault();
            handleSend();
        }
    }
    
    // Ctrl+Z or Cmd+Z to undo (when not in input)
    if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
        const activeEl = document.activeElement;
        if (activeEl.tagName !== 'INPUT' && activeEl.tagName !== 'TEXTAREA') {
            if (state.history.entries.length > 0) {
                e.preventDefault();
                performUndo();
            }
        }
    }
    
    // Escape to close modal or clear input
    if (e.key === 'Escape') {
        const modal = document.getElementById("modal");
        if (modal.classList.contains('open')) {
            closeModal();
        } else {
            const input = document.getElementById("promptInput");
            if (input.value) {
                input.value = '';
                document.getElementById("sendBtn").disabled = true;
            }
        }
    }
    
    // Ctrl+D or Cmd+D to toggle dark mode
    if ((e.ctrlKey || e.metaKey) && e.key === 'd') {
        e.preventDefault();
        toggleTheme();
    }
}

// ============================================================================
// Better Error Handling
// ============================================================================

/**
 * Validates a range address
 */
function isValidRange(address) {
    if (!address) return false;
    // Basic Excel range pattern: A1, A1:B10, Sheet1!A1:B10
    const pattern = /^([A-Za-z_][A-Za-z0-9_]*!)?(\$?[A-Z]+\$?\d+)(:\$?[A-Z]+\$?\d+)?$/i;
    return pattern.test(address);
}

/**
 * Gets a user-friendly error message
 */
function getErrorMessage(error, context = '') {
    const msg = error.message || String(error);
    
    // API errors
    if (msg.includes('401') || msg.includes('403')) {
        return 'Invalid API key. Please check your settings.';
    }
    if (msg.includes('429')) {
        return 'Rate limit exceeded. Please wait a moment and try again.';
    }
    if (msg.includes('500') || msg.includes('502') || msg.includes('503')) {
        return 'AI service temporarily unavailable. Retrying...';
    }
    if (msg.includes('network') || msg.includes('fetch')) {
        return 'Network error. Please check your connection.';
    }
    
    // Excel errors
    if (msg.includes('InvalidReference') || msg.includes('invalid range')) {
        return `Invalid cell reference${context ? ': ' + context : ''}. Please check the range.`;
    }
    if (msg.includes('RichApi')) {
        return 'Excel error. Please try again.';
    }
    
    // Generic
    return msg.length > 100 ? msg.substring(0, 100) + '...' : msg;
}

/**
 * Retries a function with exponential backoff
 */
async function withRetry(fn, maxRetries = CONFIG.MAX_RETRIES) {
    let lastError;
    for (let i = 0; i < maxRetries; i++) {
        try {
            return await fn();
        } catch (e) {
            lastError = e;
            const msg = e.message || '';
            // Only retry on transient errors
            if (msg.includes('429') || msg.includes('500') || msg.includes('502') || msg.includes('503')) {
                const delay = CONFIG.RETRY_DELAY * Math.pow(2, i);
                await new Promise(r => setTimeout(r, delay));
                continue;
            }
            throw e;
        }
    }
    throw lastError;
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
    // Check if in read-only mode
    if (state.mode === "readonly") {
        toast("Read-only mode: Cannot apply changes");
        return;
    }
    
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
    
    // Sheet creation doesn't need a range
    if (type === "sheet") {
        await createSheet(ctx, target, data);
        return;
    }
    
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
            await applyFormat(ctx, range, data);
            break;
            
        case "conditionalFormat":
            await applyConditionalFormat(ctx, range, data);
            break;
            
        case "clearFormat":
            await clearConditionalFormat(ctx, range);
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
            
        case "copy":
            await applyCopy(ctx, sheet, source, target);
            break;
            
        case "copyValues":
            await applyCopyValues(ctx, sheet, source, target);
            break;
            
        case "sheet":
            await createSheet(ctx, target, data);
            break;
            
        case "filter":
            await applyFilter(ctx, sheet, range, data);
            break;
            
        case "clearFilter":
            await clearFilter(ctx, sheet);
            break;
            
        default:
            if (data) range.values = [[data]];
    }
}

/**
 * Creates a new sheet with optional name
 */
async function createSheet(ctx, sheetName, data) {
    if (!sheetName) {
        throw new Error("Sheet name is required");
    }
    
    const sheets = ctx.workbook.worksheets;
    // Use add() with proper parameters - name is optional, position is optional
    const newSheet = sheets.add();
    newSheet.name = sheetName;
    
    // If data is provided, populate it
    if (data) {
        try {
            const values = JSON.parse(data);
            if (Array.isArray(values) && values.length > 0) {
                const range = newSheet.getRange(`A1:${String.fromCharCode(64 + values[0].length)}${values.length}`);
                range.values = values;
            }
        } catch (e) {
            // Data parsing failed, just create empty sheet
        }
    }
}

/**
 * Copies data from source range to target range
 */
async function applyCopy(ctx, sheet, source, target) {
    if (!source || !target) {
        throw new Error("Copy requires both source and target ranges");
    }
    
    const sourceRange = sheet.getRange(source);
    sourceRange.load(["values", "formulas", "rowCount", "columnCount"]);
    await ctx.sync();
    
    // Get source dimensions
    const rowCount = sourceRange.rowCount;
    const colCount = sourceRange.columnCount;
    
    // If target is a single cell, resize it to match source dimensions
    const targetCell = sheet.getRange(target);
    const targetRange = targetCell.getResizedRange(rowCount - 1, colCount - 1);
    
    // Copy formulas (preserves both formulas and values)
    targetRange.formulas = sourceRange.formulas;
}

/**
 * Copies only values (not formulas) from source range to target range
 */
async function applyCopyValues(ctx, sheet, source, target) {
    if (!source || !target) {
        throw new Error("Copy requires both source and target ranges");
    }
    
    const sourceRange = sheet.getRange(source);
    sourceRange.load(["values", "rowCount", "columnCount"]);
    await ctx.sync();
    
    // Get source dimensions
    const rowCount = sourceRange.rowCount;
    const colCount = sourceRange.columnCount;
    
    // If target is a single cell, resize it to match source dimensions
    const targetCell = sheet.getRange(target);
    const targetRange = targetCell.getResizedRange(rowCount - 1, colCount - 1);
    
    // Copy only values (converts formulas to their results)
    targetRange.values = sourceRange.values;
}

async function applyFormula(range, formula) {
    const rows = range.rowCount;
    const cols = range.columnCount;
    
    // For single cell, just set the formula
    if (rows === 1 && cols === 1) {
        range.formulas = [[formula]];
        return;
    }
    
    // For multi-row ranges, use autofill (most reliable method)
    // This works for any size and Excel handles the formula adjustment automatically
    if (rows > 1) {
        // Set formula in first cell only
        const firstCell = range.getCell(0, 0);
        firstCell.formulas = [[formula]];
        
        // Use autofill to copy down (Excel handles this efficiently)
        firstCell.autoFill(range, Excel.AutoFillType.fillDefault);
        return;
    }
    
    // For single-row, multi-column ranges, build the formula array
    const formulas = [[]];
    for (let c = 0; c < cols; c++) {
        let f = formula;
        if (c > 0) {
            // Adjust cell references for each column
            // Match pattern: optional$, column letters, optional$, row number
            f = formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colAbs, col, rowAbs, row) => {
                // If column has $ before it (absolute reference), don't adjust
                if (colAbs === "$") return match;
                // Otherwise, increment the column letter
                const colCode = col.charCodeAt(0);
                const newCol = String.fromCharCode(colCode + c);
                return `${colAbs}${newCol}${rowAbs}${row}`;
            });
        }
        formulas[0].push(f);
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

async function applyFormat(ctx, range, data) {
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

/**
 * Applies conditional formatting to a range
 * Supports multiple rules in a single action
 */
async function applyConditionalFormat(ctx, range, data) {
    let rules;
    try { 
        const parsed = JSON.parse(data);
        // Support both single rule and array of rules
        rules = Array.isArray(parsed) ? parsed : [parsed];
    } catch { 
        rules = []; 
    }
    
    // Clear existing conditional formats only once
    range.conditionalFormats.clearAll();
    await ctx.sync();
    
    // Add each conditional format rule
    for (const rule of rules) {
        if (rule.type === "cellValue" && rule.operator && rule.value) {
            const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            cf.cellValue.format.fill.color = rule.fill || "#FFFF00";
            
            // Handle different operators
            let operator = rule.operator;
            if (operator === "GreaterThan") operator = Excel.ConditionalCellValueOperator.greaterThan;
            else if (operator === "LessThan") operator = Excel.ConditionalCellValueOperator.lessThan;
            else if (operator === "EqualTo") operator = Excel.ConditionalCellValueOperator.equalTo;
            else if (operator === "GreaterThanOrEqual") operator = Excel.ConditionalCellValueOperator.greaterThanOrEqual;
            else if (operator === "LessThanOrEqual") operator = Excel.ConditionalCellValueOperator.lessThanOrEqual;
            else if (operator === "Between") operator = Excel.ConditionalCellValueOperator.between;
            
            cf.cellValue.rule = {
                formula1: String(rule.value),
                formula2: rule.value2 ? String(rule.value2) : undefined,
                operator: operator
            };
        }
    }
}

/**
 * Clears all conditional formatting from a range
 */
async function clearConditionalFormat(ctx, range) {
    range.conditionalFormats.clearAll();
    await ctx.sync();
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
    
    // Handle the data range - check if it's a valid contiguous range
    let chartDataRange = dataRange;
    const targetAddress = action.target;
    
    // Check if target contains comma (non-contiguous) - not supported directly
    if (targetAddress && targetAddress.includes(",")) {
        // For non-contiguous ranges, we need to use the first range
        // and add series manually, or use a contiguous subset
        const ranges = targetAddress.split(",").map(r => r.trim());
        // Use the full data range if available, otherwise first range
        chartDataRange = sheet.getRange(ranges[0]);
    }
    
    // Create the chart
    const chart = sheet.charts.add(type, chartDataRange, Excel.ChartSeriesBy.auto);
    
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
    
    // For line/trend charts, improve readability
    if (ct.includes("line") || ct.includes("trend")) {
        chart.legend.position = Excel.ChartLegendPosition.bottom;
    }
}

function applySort(range, data) {
    let opts = {};
    
    // Parse data - can be JSON or simple format
    if (typeof data === "string") {
        try {
            opts = JSON.parse(data);
        } catch {
            // Try to parse simple format like "column:1,ascending:true"
            const parts = data.split(",");
            for (const part of parts) {
                const [key, value] = part.split(":").map(s => s.trim());
                if (key === "column") opts.column = parseInt(value) || 0;
                if (key === "ascending") opts.ascending = value !== "false";
                if (key === "hasHeaders") opts.hasHeaders = value === "true";
            }
        }
    } else {
        opts = data || {};
    }
    
    // Default to first column, ascending, with headers
    const columnIndex = opts.column || 0;
    const ascending = opts.ascending !== false;
    const hasHeaders = opts.hasHeaders !== false; // Default to true (has headers)
    
    range.sort.apply(
        [{ 
            key: columnIndex, 
            ascending: ascending 
        }],
        false, // matchCase
        hasHeaders, // hasHeaders - true means first row is header and won't be sorted
        Excel.SortOrientation.rows
    );
}

/**
 * Applies AutoFilter to a range
 * @param {Object} ctx - Excel context
 * @param {Object} sheet - Excel worksheet
 * @param {Object} range - Excel range
 * @param {string} data - Filter criteria as JSON string
 */
async function applyFilter(ctx, sheet, range, data) {
    let filterOpts = {};
    
    // Parse filter options
    if (typeof data === "string") {
        try {
            filterOpts = JSON.parse(data);
        } catch {
            throw new Error("Invalid filter data format");
        }
    } else {
        filterOpts = data || {};
    }
    
    // Try to clear any existing autofilter (ignore errors if none exists)
    try {
        sheet.autoFilter.clearCriteria();
        await ctx.sync();
    } catch (e) {
        // No existing filter, continue
    }
    
    // Apply AutoFilter to the range
    sheet.autoFilter.apply(range);
    await ctx.sync();
    
    // If specific column filters are provided, apply them
    if (filterOpts.column !== undefined && filterOpts.values) {
        // Get the filter criteria for the specified column
        const criteria = {
            filterOn: Excel.FilterOn.values,
            values: filterOpts.values
        };
        
        // Apply the filter criteria to the column
        sheet.autoFilter.apply(range, filterOpts.column, criteria);
        await ctx.sync();
    }
}

/**
 * Clears all filters from the worksheet
 * @param {Object} ctx - Excel context
 * @param {Object} sheet - Excel worksheet
 */
async function clearFilter(ctx, sheet) {
    try {
        sheet.autoFilter.clearCriteria();
        await ctx.sync();
    } catch (e) {
        // No filter to clear, ignore
    }
}

export { handleSend, handleApply, readExcelData, clearChat };
