/*
 * Excel AI Copilot - Accurate Data Understanding
 * Enhanced with: Task-specific prompts, Function calling, RAG, Multi-step reasoning, Learning
 */

/* global document, Excel, Office, fetch, localStorage */

// Version number - increment with each update
const VERSION = "3.6.1";

import {
    detectTaskType,
    TASK_TYPES,
    enhancePrompt,
    isCorrection,
    handleCorrection,
    processResponse,
    extractResponseText,
    getRAGContext,
    getCorrectionContext,
    clearCorrections
} from "./ai-engine.js";

import {
    colIndexToLetter,
    colLetterToIndex,
    buildDataContext as buildDataContextFromModule,
    setupSelectionListener as setupSelectionListenerFromModule,
    removeSelectionListener,
    validateHeaders
} from "./excel-data.js";

import {
    setDiagnosticLogger,
    executeAction as executeActionFromModule,
    adjustFormulaReferences
} from "./action-executor.js";

import {
    initDiagnostics,
    logInfo,
    logWarn,
    logError,
    logDebug,
    getLogs,
    clearLogs,
    toggleDebugMode,
    isDebugMode,
    renderDiagnosticsPanel
} from "./diagnostics.js";

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
    allSheetsData: [],       // Data from all sheets in workbook
    conversationHistory: [],
    isFirstMessage: true,
    lastAIResponse: "",      // Track last AI response for corrections
    currentTaskType: null,   // Track current task type
    worksheetScope: "single", // "single" or "all" - controls multi-sheet access
    selectionHandler: null,  // Reference to selection change event handler
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
    },
    // Diagnostics state
    logs: [],                // Diagnostic log entries
    diagnosticsPanelVisible: false
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
    // Initialize diagnostics system
    initDiagnostics((logs) => {
        state.logs = logs;
        if (state.diagnosticsPanelVisible) {
            updateDiagnosticsPanel();
        }
    });
    
    // Set up diagnostic logger for action executor
    setDiagnosticLogger((msg) => logDebug(msg));
    
    logInfo("Excel Copilot initializing", { version: VERSION });
    
    // Comment 8: Load API key with de-obfuscation
    // Note: Key is stored with basic obfuscation for backward compatibility
    try {
        const stored = localStorage.getItem(CONFIG.STORAGE_KEY) || "";
        if (stored) {
            // Try to decode (new format) or use as-is (old format)
            try {
                state.apiKey = atob(stored);
            } catch (e) {
                // Old format - use as-is
                state.apiKey = stored;
            }
        }
    } catch (e) {
        state.apiKey = "";
        logWarn("Could not load API key");
    }
    
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
    
    // Load saved worksheet scope preference
    const savedScope = localStorage.getItem("excel_copilot_worksheet_scope");
    if (savedScope) {
        state.worksheetScope = savedScope;
    }
    bindEvents();
    initModeButtons();
    readExcelData().then(() => {
        logInfo("Initial data load complete");
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
        // Set worksheet scope radio button
        document.getElementById(state.worksheetScope === "all" ? "scopeAll" : "scopeSingle").checked = true;
        document.getElementById("modal").classList.add("open");
    });
    
    document.getElementById("closeModal")?.addEventListener("click", closeModal);
    document.getElementById("cancelBtn")?.addEventListener("click", closeModal);
    document.getElementById("saveBtn")?.addEventListener("click", async () => {
        state.apiKey = document.getElementById("apiKeyInput").value.trim();
        
        // Comment 8: Store API key with minimal obfuscation
        // Note: For better security, consider not persisting at all
        if (state.apiKey) {
            // Simple base64 encoding (not true encryption, but prevents casual viewing)
            const obfuscated = btoa(state.apiKey);
            localStorage.setItem(CONFIG.STORAGE_KEY, obfuscated);
        } else {
            localStorage.removeItem(CONFIG.STORAGE_KEY);
        }
        
        // Save worksheet scope preference
        const selectedScope = document.querySelector('input[name="worksheetScope"]:checked')?.value || "single";
        const scopeChanged = state.worksheetScope !== selectedScope;
        state.worksheetScope = selectedScope;
        localStorage.setItem("excel_copilot_worksheet_scope", selectedScope);
        
        // Comment 6: Re-attach selection listener and refresh data when scope changes
        if (scopeChanged) {
            await reattachSelectionListener();
        }
        
        // Refresh data to apply new scope
        await readExcelData();
        
        closeModal();
        toast("Saved");
        logInfo("Settings saved", { scope: selectedScope });
    });
    
    // Comment 8: Add "Remove API key" functionality
    document.getElementById("removeApiKeyBtn")?.addEventListener("click", () => {
        state.apiKey = "";
        localStorage.removeItem(CONFIG.STORAGE_KEY);
        document.getElementById("apiKeyInput").value = "";
        toast("API key removed");
        logInfo("API key removed");
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
    
    // Comment 10: Diagnostics panel buttons
    document.getElementById("diagnosticsBtn")?.addEventListener("click", toggleDiagnosticsPanel);
    document.getElementById("clearLogsBtn")?.addEventListener("click", () => {
        clearLogs();
        updateDiagnosticsPanel();
        toast("Logs cleared");
    });
    document.getElementById("toggleDebugBtn")?.addEventListener("click", () => {
        const newState = toggleDebugMode();
        toast(newState ? "Debug mode enabled" : "Debug mode disabled");
        updateDebugModeCheckbox();
    });
    document.getElementById("debugModeCheckbox")?.addEventListener("change", (e) => {
        if (e.target.checked) {
            toggleDebugMode();
        } else {
            toggleDebugMode();
        }
    });
    
    // Update buttons
    document.getElementById("checkUpdateBtn")?.addEventListener("click", checkForUpdatesInSettings);
    document.getElementById("updateNowBtn")?.addEventListener("click", performUpdate);
    
    // Update current version text when settings opens
    document.getElementById("settingsBtn")?.addEventListener("click", () => {
        const versionText = document.getElementById("currentVersionText");
        if (versionText) versionText.textContent = `v${VERSION}`;
    });
    
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
        // Remove existing handler if any (Comment 6)
        if (state.selectionHandler) {
            await removeSelectionListener(state.selectionHandler);
            state.selectionHandler = null;
        }
        
        await Excel.run(async (ctx) => {
            const worksheet = ctx.workbook.worksheets.getActiveWorksheet();
            state.selectionHandler = worksheet.onSelectionChanged.add(readExcelData);
            await ctx.sync();
            logDebug("Selection listener attached");
        });
    } catch (e) {
        // Log warning instead of silently ignoring (Comment 1)
        console.warn("Could not attach selection listener:", e);
        logWarn(`Selection listener failed: ${e.message}`);
        toast("Selection auto-refresh unavailable");
    }
}

/**
 * Re-attaches selection listener (called when worksheet scope changes)
 */
async function reattachSelectionListener() {
    await setupSelectionListener();
}

// ============================================================================
// Read Excel Data with Column Headers
// Note: colIndexToLetter and colLetterToIndex are imported from excel-data.js
// ============================================================================
async function readExcelData() {
    const infoEl = document.getElementById("contextInfo");
    
    try {
        await Excel.run(async (ctx) => {
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            // Read sheets based on scope setting
            const allSheetsData = [];
            const shouldReadAllSheets = state.worksheetScope === "all";
            
            // Get active sheet first
            const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
            activeSheet.load("name");
            await ctx.sync();
            
            const activeSheetName = activeSheet.name;
            
            // Determine which sheets to read
            const sheetsToRead = shouldReadAllSheets 
                ? sheets.items.slice(0, 10) // All sheets (max 10)
                : [sheets.items.find(s => s.name === activeSheetName) || sheets.items[0]]; // Just active sheet
            
            for (const sheet of sheetsToRead) {
                try {
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
                    
                    // Comment 2: Handle empty sheets
                    if (rowCount === 0 || !values || values.length === 0) {
                        logDebug(`Sheet "${sheetName}" has no data, skipping`);
                        continue;
                    }
                    
                    // Comment 2: Guard for empty values array before reading headers
                    if (values.length === 0) {
                        logDebug(`Sheet "${sheetName}" has empty values array, skipping`);
                        continue;
                    }
                    
                    // Detect headers (first row) with validation (Comment 2)
                    const headers = values[0] || [];
                    const headerValidation = validateHeaders(headers);
                    
                    // Build column mapping with validated headers
                    const columnMap = [];
                    for (let c = 0; c < colCount; c++) {
                        const colLetter = colIndexToLetter(startCol + c);
                        let headerName;
                        if (headerValidation.isValid && headers[c]) {
                            headerName = headers[c];
                        } else {
                            // Use generic column names if headers don't look valid
                            headerName = `Column ${colLetter}`;
                        }
                        columnMap.push({
                            letter: colLetter,
                            index: c,
                            header: headerName
                        });
                    }
                    
                    allSheetsData.push({
                        sheetName,
                        address: usedRange.address,
                        values,
                        headers: headerValidation.isValid ? headers : columnMap.map(c => c.header),
                        columnMap,
                        startRow: startRow + 1,
                        startCol: colIndexToLetter(startCol),
                        rowCount,
                        colCount,
                        dataStartRow: startRow + 2,
                        headerValidation
                    });
                } catch (e) {
                    // Sheet might be empty, log and skip it (Comment 1)
                    const sheetName = sheet.name || "Unknown";
                    console.warn(`Skipping sheet ${sheetName}:`, e);
                    logWarn(`Failed to read sheet "${sheetName}": ${e.message}`);
                }
            }
            
            // Comment 2: Handle case where no sheets have usable data
            if (allSheetsData.length === 0) {
                state.currentData = null;
                state.allSheetsData = [];
                infoEl.textContent = "No usable data found in any sheet";
                logWarn("No usable data found in any sheet");
                return;
            }
            
            // Set current data to active sheet
            const activeSheetData = allSheetsData.find(s => s.sheetName === activeSheetName);
            state.currentData = activeSheetData || allSheetsData[0] || null;
            state.allSheetsData = shouldReadAllSheets ? allSheetsData : [];
            
            if (state.currentData) {
                const scopeText = shouldReadAllSheets ? ` (${allSheetsData.length} sheets)` : "";
                infoEl.textContent = `${state.currentData.sheetName}: ${state.currentData.rowCount} rows × ${state.currentData.colCount} cols${scopeText}`;
            } else {
                infoEl.textContent = "No data";
            }
        });
    } catch (e) {
        // Comment 1: Log actual error and show meaningful message
        console.error("Failed to read Excel data:", e);
        const errorReason = e.message || "Unknown error";
        infoEl.textContent = `Failed to read data: ${errorReason.substring(0, 50)}`;
        state.currentData = null;
        state.allSheetsData = [];
        logError(`readExcelData error: ${errorReason}`);
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
        
        // Extract version from the deployed code (handles minified Ae="x.x.x" and source const VERSION = "x.x.x")
        const versionMatch = code.match(/(?:[A-Za-z]{1,2}="|const VERSION\s*=\s*")(\d+\.\d+\.\d+)"/);
        
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

// Store latest available version for update
let latestAvailableVersion = null;

/**
 * Checks for updates from the settings modal
 */
async function checkForUpdatesInSettings() {
    const checkBtn = document.getElementById("checkUpdateBtn");
    const updateBtn = document.getElementById("updateNowBtn");
    const statusEl = document.getElementById("updateStatus");
    
    const originalBtnText = checkBtn.innerHTML;
    
    try {
        checkBtn.innerHTML = `<svg class="spin" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M21 12a9 9 0 11-6.219-8.56"/></svg> Checking...`;
        checkBtn.disabled = true;
        statusEl.className = "hint update-checking";
        statusEl.textContent = "Checking for updates...";
        
        // Fetch the deployed taskpane.js with cache-busting
        const response = await fetch(`https://tankuday21.github.io/GeminiForExcel/taskpane.js?t=${Date.now()}`);
        const code = await response.text();
        
        // Extract version from the deployed code (handles minified Ae="x.x.x" and source const VERSION = "x.x.x")
        const versionMatch = code.match(/(?:[A-Za-z]{1,2}="|const VERSION\s*=\s*")(\d+\.\d+\.\d+)"/);
        
        if (versionMatch) {
            const deployedVersion = versionMatch[1];
            const currentVersion = VERSION;
            
            if (deployedVersion === currentVersion) {
                statusEl.className = "hint";
                statusEl.innerHTML = `✓ You're on the latest version (<strong>v${currentVersion}</strong>)`;
                updateBtn.style.display = "none";
                latestAvailableVersion = null;
                toast("You're up to date!");
            } else {
                latestAvailableVersion = deployedVersion;
                statusEl.className = "hint update-available";
                statusEl.innerHTML = `Update available! <strong>v${currentVersion}</strong> → <strong>v${deployedVersion}</strong>`;
                updateBtn.style.display = "inline-flex";
                toast(`Update available: v${deployedVersion}`);
                logInfo(`Update available: ${currentVersion} -> ${deployedVersion}`);
            }
        } else {
            throw new Error("Could not parse version");
        }
    } catch (error) {
        console.error("Update check failed:", error);
        statusEl.className = "hint";
        statusEl.textContent = `Failed to check for updates. Current: v${VERSION}`;
        updateBtn.style.display = "none";
        toast("Update check failed");
        logError(`Update check failed: ${error.message}`);
    } finally {
        checkBtn.innerHTML = originalBtnText;
        checkBtn.disabled = false;
    }
}

/**
 * Performs the update by reloading the add-in
 */
async function performUpdate() {
    const updateBtn = document.getElementById("updateNowBtn");
    const statusEl = document.getElementById("updateStatus");
    
    updateBtn.innerHTML = `<svg class="spin" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M21 12a9 9 0 11-6.219-8.56"/></svg> Updating...`;
    updateBtn.disabled = true;
    
    try {
        // Clear Office cache if possible
        statusEl.textContent = "Clearing cache and reloading...";
        
        // Log the update attempt
        logInfo(`Updating to v${latestAvailableVersion}`);
        
        // Clear any cached data
        if ('caches' in window) {
            const cacheNames = await caches.keys();
            await Promise.all(cacheNames.map(name => caches.delete(name)));
        }
        
        // Force reload the page to get the latest version
        // The add-in will reload from the server with the new version
        toast("Reloading add-in...");
        
        // Small delay to show the toast
        setTimeout(() => {
            window.location.reload(true);
        }, 500);
        
    } catch (error) {
        console.error("Update failed:", error);
        statusEl.textContent = `Update failed: ${error.message}`;
        updateBtn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg> Retry Update`;
        updateBtn.disabled = false;
        toast("Update failed");
        logError(`Update failed: ${error.message}`);
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
    
    // Comment 1: Separate readExcelData from AI call with specific error handling
    try {
        await readExcelData();
    } catch (dataError) {
        console.error("Failed to read Excel data:", dataError);
        logError(`Data read failed: ${dataError.message}`);
        addMessage("ai", "Could not load Excel data. Please ensure the workbook is open and the sheet has a used range.", "error");
        toast("Data read failed");
        return; // Short-circuit before AI call
    }
    
    // Check if we have data to work with
    if (!state.currentData) {
        addMessage("ai", "No Excel data available. Please ensure your workbook has data in the active sheet.", "error");
        return;
    }
    
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
        [TASK_TYPES.TABLE]: "� TaGble Mode",
        [TASK_TYPES.PIVOT]: "🔄 Pivot Mode",
        [TASK_TYPES.DATA_MANIPULATION]: "✂️ Data Manipulation Mode",
        [TASK_TYPES.SHAPES]: "🔷 Shapes Mode",
        [TASK_TYPES.COMMENTS]: "💬 Comments Mode",
        [TASK_TYPES.PROTECTION]: "🔒 Protection Mode",
        [TASK_TYPES.PAGE_SETUP]: "📄 Page Setup Mode",
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
        [TASK_TYPES.TABLE]: "**[Table]** ",
        [TASK_TYPES.PIVOT]: "**[Pivot]** ",
        [TASK_TYPES.DATA_MANIPULATION]: "**[Data]** ",
        [TASK_TYPES.SHAPES]: "**[Shapes]** ",
        [TASK_TYPES.COMMENTS]: "**[Comments]** ",
        [TASK_TYPES.PROTECTION]: "**[Protection]** ",
        [TASK_TYPES.PAGE_SETUP]: "**[Page Setup]** ",
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
    
    // Make API call without retry logic - let errors through
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
    
    if (!res.ok) {
        const errorData = await res.json().catch(() => ({}));
        const errorMessage = errorData.error?.message || `HTTP ${res.status}`;
        throw new Error(`API Error: ${errorMessage}`);
    }
    
    const data = await res.json();
    
    // Comment 5: Use robust response extraction
    const extracted = extractResponseText(data);
    
    if (extracted.error) {
        logWarn(`AI response issue: ${extracted.error}`);
        if (extracted.blocked) {
            throw new Error(extracted.error);
        }
        // Show toast for empty responses
        toast(extracted.error);
        return extracted.error;
    }
    
    if (!extracted.text) {
        logWarn("AI returned no content");
        toast("AI returned no content");
        return "No response from AI";
    }
    
    const response = extracted.text;
    
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
    
    // List all sheets in workbook
    if (state.allSheetsData && state.allSheetsData.length > 1) {
        context += `**Available Sheets:** ${state.allSheetsData.map(s => s.sheetName).join(", ")}\n`;
        context += `**Active Sheet:** ${sheetName}\n\n`;
    } else {
        context += `**Sheet:** ${sheetName}\n`;
    }
    
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
    
    // Add information about other sheets
    if (state.allSheetsData && state.allSheetsData.length > 1) {
        context += `\n## OTHER SHEETS IN WORKBOOK\n`;
        for (const sheet of state.allSheetsData) {
            if (sheet.sheetName === sheetName) continue; // Skip current sheet
            
            context += `\n### ${sheet.sheetName}\n`;
            context += `- Columns: ${sheet.headers.join(", ")}\n`;
            context += `- Rows: ${sheet.rowCount}\n`;
            
            // Show first few rows as sample
            if (sheet.values.length > 1) {
                context += `- Sample data (first 3 rows):\n`;
                for (let r = 0; r < Math.min(3, sheet.values.length); r++) {
                    const row = sheet.values[r];
                    context += `  ${r === 0 ? "Headers" : `Row ${r}`}: ${row.slice(0, 5).join(" | ")}\n`;
                }
            }
        }
        context += `\n**Note:** You can reference data from any sheet using sheet name (e.g., DeptManagers!A2:B10)\n`;
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
    
    // API errors - show actual error message instead of generic ones
    if (msg.includes('401') || msg.includes('403')) {
        return 'Invalid API key. Please check your settings.';
    }
    // Removed 429 rate limit check - show actual error
    if (msg.includes('500') || msg.includes('502') || msg.includes('503')) {
        return 'AI service temporarily unavailable. Please try again.';
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
    
    // Generic - show full error message
    return msg;
}

/**
 * Retries a function with exponential backoff
 * Removed 429 rate limiting - only retry on server errors
 */
async function withRetry(fn, maxRetries = CONFIG.MAX_RETRIES) {
    let lastError;
    for (let i = 0; i < maxRetries; i++) {
        try {
            return await fn();
        } catch (e) {
            lastError = e;
            const msg = e.message || '';
            // Only retry on server errors (not rate limits)
            if (msg.includes('500') || msg.includes('502') || msg.includes('503')) {
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
// Comment 10: Diagnostics Panel Functions
// ============================================================================

/**
 * Toggles diagnostics panel visibility
 */
function toggleDiagnosticsPanel() {
    state.diagnosticsPanelVisible = !state.diagnosticsPanelVisible;
    const panel = document.getElementById("diagnosticsPanel");
    const btn = document.getElementById("diagnosticsBtn");
    
    if (panel) {
        panel.style.display = state.diagnosticsPanelVisible ? "block" : "none";
        if (state.diagnosticsPanelVisible) {
            updateDiagnosticsPanel();
        }
    }
    
    if (btn) {
        btn.classList.toggle("active", state.diagnosticsPanelVisible);
    }
}

/**
 * Updates the diagnostics panel content
 */
function updateDiagnosticsPanel() {
    const list = document.getElementById("diagnosticsList");
    if (!list) return;
    
    const logs = getLogs();
    list.innerHTML = renderDiagnosticsPanel(logs.slice(0, 50));
}

/**
 * Updates the debug mode checkbox state
 */
function updateDebugModeCheckbox() {
    const checkbox = document.getElementById("debugModeCheckbox");
    if (checkbox) {
        checkbox.checked = isDebugMode();
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
    
    // Actions that use logical names (table names, PivotTable names) instead of range addresses
    // These should NOT pre-load a range as the target is not a valid range address
    const logicalNameActions = [
        "createPivotTable",      // target is source range, but handler resolves it
        "addPivotField",         // target is PivotTable name
        "configurePivotLayout",  // target is PivotTable name
        "refreshPivotTable",     // target is PivotTable name
        "deletePivotTable",      // target is PivotTable name
        "styleTable",            // target is table name
        "addTableRow",           // target is table name
        "addTableColumn",        // target is table name
        "resizeTable",           // target is table name
        "convertToRange",        // target is table name
        "toggleTableTotals",     // target is table name
        "insertRows",            // target is row number, not range
        "insertColumns",         // target is column letter, not range
        "deleteRows",            // target is row range like "10:15"
        "deleteColumns",         // target is column range like "D:F"
        "createSlicer",          // target is table/pivot name
        "configureSlicer",       // target is slicer name
        "connectSlicerToTable",  // target is slicer name
        "connectSlicerToPivot",  // target is slicer name
        "deleteSlicer",          // target is slicer name
        "deleteNamedRange",      // target is named range name
        "updateNamedRange",      // target is named range name
        "listNamedRanges",       // target is scope option
        "formatShape",           // target is shape name
        "deleteShape",           // target is shape name
        "groupShapes",           // target is shape names (comma-separated)
        "arrangeShapes",         // target is shape name
        "ungroupShapes",         // target is group name
        "addComment",            // target is cell address (comment API handles it)
        "addNote",               // target is cell address (note API handles it)
        "editComment",           // target is cell with comment
        "editNote",              // target is cell with note
        "deleteComment",         // target is cell with comment
        "deleteNote",            // target is cell with note
        "replyToComment",        // target is cell with parent comment
        "resolveComment",        // target is cell with comment
        "createSparkline",       // target is location cell/range
        "configureSparkline",    // target is sparkline location
        "deleteSparkline",       // target is sparkline location
        "renameSheet",           // target is sheet name
        "moveSheet",             // target is sheet name
        "hideSheet",             // target is sheet name
        "unhideSheet",           // target is sheet name
        "unfreezePane",          // target is "current" or sheet name
        "setZoom",               // target is "current" or sheet name
        "createView",            // target is view name
        "setPageSetup",          // target is sheet name
        "setPageMargins",        // target is sheet name
        "setPageOrientation",    // target is sheet name
        "setHeaderFooter",       // target is sheet name
        "setPageBreaks"          // target is sheet name
    ];
    
    // Only pre-load range for actions that actually need it
    let range = null;
    if (!logicalNameActions.includes(type)) {
        range = sheet.getRange(target);
        range.load(["rowCount", "columnCount"]);
        await ctx.sync();
    }
    
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
            await createChart(ctx, sheet, range, action);
            break;
            
        case "pivotChart":
            await createPivotChart(ctx, sheet, range, action);
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
            
        case "removeDuplicates":
            await removeDuplicates(ctx, range, data);
            break;
            
        case "createTable":
            await createTable(ctx, sheet, range, action);
            break;
            
        case "styleTable":
            await styleTable(ctx, sheet, target, data);
            break;
            
        case "addTableRow":
            await addTableRow(ctx, sheet, action);
            break;
            
        case "addTableColumn":
            await addTableColumn(ctx, sheet, action);
            break;
            
        case "resizeTable":
            await resizeTable(ctx, sheet, action);
            break;
            
        case "convertToRange":
            await convertToRange(ctx, sheet, target);
            break;
            
        case "toggleTableTotals":
            await toggleTableTotals(ctx, sheet, action);
            break;
            
        case "insertRows":
            await insertRows(ctx, sheet, action);
            break;
            
        case "insertColumns":
            await insertColumns(ctx, sheet, action);
            break;
            
        case "deleteRows":
            await deleteRows(ctx, sheet, action);
            break;
            
        case "deleteColumns":
            await deleteColumns(ctx, sheet, action);
            break;
            
        case "mergeCells":
            await mergeCells(ctx, sheet, action);
            break;
            
        case "unmergeCells":
            await unmergeCells(ctx, sheet, action);
            break;
            
        case "findReplace":
            await findReplace(ctx, sheet, action);
            break;
            
        case "textToColumns":
            await textToColumns(ctx, sheet, action);
            break;
            
        case "createPivotTable":
            await createPivotTable(ctx, sheet, action);
            break;
            
        case "addPivotField":
            await addPivotField(ctx, sheet, action);
            break;
            
        case "configurePivotLayout":
            await configurePivotLayout(ctx, sheet, action);
            break;
            
        case "refreshPivotTable":
            await refreshPivotTable(ctx, sheet, action);
            break;
            
        case "deletePivotTable":
            await deletePivotTable(ctx, sheet, action);
            break;
            
        case "createSlicer":
            await createSlicer(ctx, sheet, action);
            break;
            
        case "configureSlicer":
            await configureSlicer(ctx, sheet, action);
            break;
            
        case "connectSlicerToTable":
            await connectSlicerToTable(ctx, sheet, action);
            break;
            
        case "connectSlicerToPivot":
            await connectSlicerToPivot(ctx, sheet, action);
            break;
            
        case "deleteSlicer":
            await deleteSlicer(ctx, sheet, action);
            break;
            
        case "createNamedRange":
            await createNamedRange(ctx, sheet, action);
            break;
            
        case "deleteNamedRange":
            await deleteNamedRange(ctx, sheet, action);
            break;
            
        case "updateNamedRange":
            await updateNamedRange(ctx, sheet, action);
            break;
            
        case "listNamedRanges":
            await listNamedRanges(ctx, sheet, action);
            break;
            
        case "protectWorksheet":
            await protectWorksheet(ctx, sheet, action);
            break;
            
        case "unprotectWorksheet":
            await unprotectWorksheet(ctx, sheet, action);
            break;
            
        case "protectRange":
            await protectRange(ctx, sheet, action);
            break;
            
        case "unprotectRange":
            await unprotectRange(ctx, sheet, action);
            break;
            
        case "protectWorkbook":
            await protectWorkbook(ctx, sheet, action);
            break;
            
        case "unprotectWorkbook":
            await unprotectWorkbook(ctx, sheet, action);
            break;
            
        case "insertShape":
            await insertShape(ctx, sheet, action);
            break;
            
        case "insertImage":
            await insertImage(ctx, sheet, action);
            break;
            
        case "insertTextBox":
            await insertTextBox(ctx, sheet, action);
            break;
            
        case "formatShape":
            await formatShape(ctx, sheet, target, data);
            break;
            
        case "deleteShape":
            await deleteShape(ctx, sheet, target);
            break;
            
        case "groupShapes":
            await groupShapes(ctx, sheet, action);
            break;
            
        case "arrangeShapes":
            await arrangeShapes(ctx, sheet, target, data);
            break;
            
        case "ungroupShapes":
            await ungroupShapes(ctx, sheet, target);
            break;
            
        case "addComment":
            await addComment(ctx, sheet, action);
            break;
            
        case "addNote":
            await addNote(ctx, sheet, action);
            break;
            
        case "editComment":
            await editComment(ctx, sheet, action);
            break;
            
        case "editNote":
            await editNote(ctx, sheet, action);
            break;
            
        case "deleteComment":
            await deleteComment(ctx, sheet, action);
            break;
            
        case "deleteNote":
            await deleteNote(ctx, sheet, action);
            break;
            
        case "replyToComment":
            await replyToComment(ctx, sheet, action);
            break;
            
        case "resolveComment":
            await resolveComment(ctx, sheet, action);
            break;
            
        case "createSparkline":
            await createSparkline(ctx, sheet, action);
            break;
            
        case "configureSparkline":
            await configureSparkline(ctx, sheet, action);
            break;
            
        case "deleteSparkline":
            await deleteSparkline(ctx, sheet, action);
            break;
            
        case "renameSheet":
            await renameSheet(ctx, sheet, action);
            break;
            
        case "moveSheet":
            await moveSheet(ctx, sheet, action);
            break;
            
        case "hideSheet":
            await hideSheet(ctx, sheet, action);
            break;
            
        case "unhideSheet":
            await unhideSheet(ctx, sheet, action);
            break;
            
        case "freezePanes":
            await freezePanes(ctx, sheet, action);
            break;
            
        case "unfreezePane":
            await unfreezePane(ctx, sheet, action);
            break;
            
        case "setZoom":
            await setZoom(ctx, sheet, action);
            break;
            
        case "splitPane":
            await splitPane(ctx, sheet, action);
            break;
            
        case "createView":
            await createView(ctx, sheet, action);
            break;
            
        case "addHyperlink":
            await addHyperlink(ctx, range, data);
            break;
            
        case "removeHyperlink":
            await removeHyperlink(ctx, range);
            break;
            
        case "editHyperlink":
            await editHyperlink(ctx, range, data);
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
    
    // Parse target to get the starting cell
    // If target is already a range (e.g., "A2:A51"), extract just the first cell
    const targetAddress = target.includes(":") ? target.split(":")[0] : target;
    
    // Get the starting cell and resize to match source
    const targetCell = sheet.getRange(targetAddress);
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
    
    // For multi-row, single-column ranges, use autofill approach
    if (rows > 1 && cols === 1) {
        // Set formula in first cell only
        const firstCell = range.getCell(0, 0);
        firstCell.formulas = [[formula]];
        
        // Try autofill first (most efficient)
        try {
            firstCell.autoFill(range, Excel.AutoFillType.fillDefault);
            return;
        } catch (autofillError) {
            // Autofill failed, use manual array method
            console.warn("Autofill failed, using formula array:", autofillError);
            
            // Build formula array manually
            const formulas = [];
            for (let r = 0; r < rows; r++) {
                let f = formula;
                if (r > 0) {
                    // Adjust row numbers in cell references (but not absolute references)
                    f = formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colAbs, col, rowAbs, row) => {
                        if (rowAbs === "$") return match; // Skip absolute row references
                        const newRow = parseInt(row) + r;
                        return `${colAbs}${col}${rowAbs}${newRow}`;
                    });
                }
                formulas.push([f]);
            }
            
            // Apply all formulas at once
            range.formulas = formulas;
            return;
        }
    }
    
    // For single-row, multi-column ranges, build the formula array
    if (rows === 1 && cols > 1) {
        const formulas = [[]];
        for (let c = 0; c < cols; c++) {
            let f = formula;
            if (c > 0) {
                // Comment 3: Use robust base-26 conversion for multi-letter columns
                f = formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colAbs, col, rowAbs, row) => {
                    if (colAbs === "$") return match; // Skip absolute column references
                    // Use colLetterToIndex and colIndexToLetter for multi-letter support
                    const colIndex = colLetterToIndex(col);
                    const newCol = colIndexToLetter(colIndex + c);
                    return `${colAbs}${newCol}${rowAbs}${row}`;
                });
            }
            formulas[0].push(f);
        }
        range.formulas = formulas;
        return;
    }
    
    // For multi-row, multi-column ranges, build 2D formula array
    if (rows > 1 && cols > 1) {
        const formulas = [];
        for (let r = 0; r < rows; r++) {
            const rowFormulas = [];
            for (let c = 0; c < cols; c++) {
                let f = formula;
                // Adjust both row and column references
                if (r > 0 || c > 0) {
                    // Comment 3: Use robust base-26 conversion for multi-letter columns
                    f = formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colAbs, col, rowAbs, row) => {
                        let newCol = col;
                        let newRow = parseInt(row);
                        
                        // Adjust column if not absolute - use robust conversion
                        if (colAbs !== "$" && c > 0) {
                            const colIndex = colLetterToIndex(col);
                            newCol = colIndexToLetter(colIndex + c);
                        }
                        
                        // Adjust row if not absolute
                        if (rowAbs !== "$" && r > 0) {
                            newRow = newRow + r;
                        }
                        
                        return `${colAbs}${newCol}${rowAbs}${newRow}`;
                    });
                }
                rowFormulas.push(f);
            }
            formulas.push(rowFormulas);
        }
        range.formulas = formulas;
        return;
    }
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
    
    // ========== Font Properties ==========
    if (fmt.bold !== undefined) range.format.font.bold = fmt.bold;
    if (fmt.italic !== undefined) range.format.font.italic = fmt.italic;
    if (fmt.fontColor) range.format.font.color = fmt.fontColor;
    if (fmt.fontSize) range.format.font.size = fmt.fontSize;
    
    // ========== Fill Properties ==========
    if (fmt.fill) range.format.fill.color = fmt.fill;
    
    // ========== Alignment Properties ==========
    const validHorizontalAlignments = ["General", "Left", "Center", "Right", "Fill", "Justify", "CenterAcrossSelection", "Distributed"];
    const validVerticalAlignments = ["Top", "Center", "Bottom", "Justify", "Distributed"];
    
    if (fmt.horizontalAlignment && validHorizontalAlignments.includes(fmt.horizontalAlignment)) {
        range.format.horizontalAlignment = fmt.horizontalAlignment;
    }
    if (fmt.verticalAlignment && validVerticalAlignments.includes(fmt.verticalAlignment)) {
        range.format.verticalAlignment = fmt.verticalAlignment;
    }
    
    // ========== Text Control Properties ==========
    if (fmt.wrapText !== undefined) range.format.wrapText = fmt.wrapText;
    if (fmt.textOrientation !== undefined) {
        const orientation = parseInt(fmt.textOrientation);
        if ((orientation >= -90 && orientation <= 90) || orientation === 255) {
            range.format.textOrientation = orientation;
        }
    }
    if (fmt.indentLevel !== undefined) {
        const indent = parseInt(fmt.indentLevel);
        if (indent >= 0 && indent <= 250) range.format.indentLevel = indent;
    }
    if (fmt.shrinkToFit !== undefined) range.format.shrinkToFit = fmt.shrinkToFit;
    if (fmt.readingOrder) {
        const validReadingOrders = ["Context", "LeftToRight", "RightToLeft"];
        if (validReadingOrders.includes(fmt.readingOrder)) {
            range.format.readingOrder = fmt.readingOrder;
        }
    }
    
    // ========== Number Format ==========
    const numberFormatPresets = {
        "currency": "$#,##0.00",
        "accounting": "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        "percentage": "0.00%",
        "date": "m/d/yyyy",
        "dateShort": "mm/dd/yy",
        "dateLong": "dddd, mmmm dd, yyyy",
        "time": "h:mm:ss AM/PM",
        "timeShort": "h:mm AM/PM",
        "time24": "hh:mm:ss",
        "fraction": "# ?/?",
        "scientific": "0.00E+00",
        "text": "@",
        "number": "#,##0.00",
        "integer": "#,##0"
    };
    
    if (fmt.numberFormatPreset && numberFormatPresets[fmt.numberFormatPreset]) {
        range.numberFormat = [[numberFormatPresets[fmt.numberFormatPreset]]];
    } else if (fmt.numberFormat) {
        range.numberFormat = [[fmt.numberFormat]];
    }
    
    // ========== Cell Style ==========
    const validStyles = [
        "Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "Title", "Total",
        "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
        "Good", "Bad", "Neutral", "Warning Text",
        "Input", "Output", "Calculation", "Check Cell", "Explanatory Text", "Linked Cell", "Note"
    ];
    
    if (fmt.style && validStyles.includes(fmt.style)) {
        try { range.format.style = fmt.style; } catch (e) { console.warn("Style error:", e); }
    }
    
    // ========== Border Properties ==========
    const validBorderStyles = ["Continuous", "Dash", "DashDot", "DashDotDot", "Dot", "Double", "None"];
    const validBorderWeights = ["Hairline", "Thin", "Medium", "Thick"];
    const borderSides = {
        "top": "EdgeTop", "bottom": "EdgeBottom", "left": "EdgeLeft", "right": "EdgeRight",
        "insideHorizontal": "InsideHorizontal", "insideVertical": "InsideVertical",
        "diagonalDown": "DiagonalDown", "diagonalUp": "DiagonalUp"
    };
    
    // Simple border (backward compatible) - applies continuous black thin borders to all edges
    if (fmt.border === true) {
        const edgeTop = range.format.borders.getItem("EdgeTop");
        edgeTop.style = "Continuous";
        edgeTop.color = "#000000";
        edgeTop.weight = "Thin";
        
        const edgeBottom = range.format.borders.getItem("EdgeBottom");
        edgeBottom.style = "Continuous";
        edgeBottom.color = "#000000";
        edgeBottom.weight = "Thin";
        
        const edgeLeft = range.format.borders.getItem("EdgeLeft");
        edgeLeft.style = "Continuous";
        edgeLeft.color = "#000000";
        edgeLeft.weight = "Thin";
        
        const edgeRight = range.format.borders.getItem("EdgeRight");
        edgeRight.style = "Continuous";
        edgeRight.color = "#000000";
        edgeRight.weight = "Thin";
    }
    
    // Advanced borders (individual sides with style/color/weight)
    if (fmt.borders && typeof fmt.borders === "object") {
        for (const [side, borderConfig] of Object.entries(fmt.borders)) {
            const excelSide = borderSides[side];
            if (!excelSide) continue;
            
            try {
                const border = range.format.borders.getItem(excelSide);
                if (borderConfig.style && validBorderStyles.includes(borderConfig.style)) {
                    border.style = borderConfig.style;
                } else {
                    border.style = "Continuous";
                }
                if (borderConfig.color) border.color = borderConfig.color;
                if (borderConfig.weight && validBorderWeights.includes(borderConfig.weight)) {
                    border.weight = borderConfig.weight;
                }
            } catch (e) { console.warn("Border error:", e); }
        }
    }
}

/**
 * Applies conditional formatting to a range with comprehensive Office.js support
 * Supports multiple rules in a single action
 * 
 * Supported types: cellValue, colorScale, dataBar, iconSet, topBottom, preset, textComparison, custom
 */
async function applyConditionalFormat(ctx, range, data) {
    let rules;
    try { 
        const parsed = JSON.parse(data);
        rules = Array.isArray(parsed) ? parsed : [parsed];
    } catch { 
        rules = []; 
    }
    
    // Validation helper for hex colors
    const isValidHexColor = (color) => /^#[0-9A-Fa-f]{6}$/.test(color);
    
    // Valid icon set styles
    const validIconSets = [
        "threeArrows", "threeArrowsGray", "threeTriangles", "threeFlags",
        "threeTrafficLights1", "threeTrafficLights2", "threeSigns",
        "threeSymbols", "threeSymbols2", "threeStars",
        "fourArrows", "fourArrowsGray", "fourRedToBlack", "fourRating", "fourTrafficLights",
        "fiveArrows", "fiveArrowsGray", "fiveRating", "fiveQuarters", "fiveBoxes"
    ];
    
    // Valid preset criteria
    const validPresetCriteria = [
        "duplicateValues", "uniqueValues", "aboveAverage", "belowAverage",
        "equalOrAboveAverage", "equalOrBelowAverage",
        "oneStdDevAboveAverage", "oneStdDevBelowAverage",
        "twoStdDevAboveAverage", "twoStdDevBelowAverage",
        "threeStdDevAboveAverage", "threeStdDevBelowAverage",
        "yesterday", "today", "tomorrow", "lastSevenDays",
        "lastWeek", "thisWeek", "nextWeek",
        "lastMonth", "thisMonth", "nextMonth"
    ];
    
    range.conditionalFormats.clearAll();
    await ctx.sync();
    
    for (const rule of rules) {
        try {
            const ruleType = rule.type || "cellValue";
            
            // ========== Cell Value ==========
            if (ruleType === "cellValue" && rule.operator && rule.value !== undefined) {
                // Map operator string to Excel enum with validation
                const operatorMap = {
                    "GreaterThan": Excel.ConditionalCellValueOperator.greaterThan,
                    "LessThan": Excel.ConditionalCellValueOperator.lessThan,
                    "EqualTo": Excel.ConditionalCellValueOperator.equalTo,
                    "NotEqualTo": Excel.ConditionalCellValueOperator.notEqual,
                    "GreaterThanOrEqual": Excel.ConditionalCellValueOperator.greaterThanOrEqual,
                    "LessThanOrEqual": Excel.ConditionalCellValueOperator.lessThanOrEqual,
                    "Between": Excel.ConditionalCellValueOperator.between
                };
                
                const operator = operatorMap[rule.operator];
                if (!operator) {
                    console.warn(`Invalid cellValue operator: ${rule.operator}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                
                // Apply fill color with validation
                const fillColor = rule.fill || "#FFFF00";
                cf.cellValue.format.fill.color = isValidHexColor(fillColor) ? fillColor : "#FFFF00";
                
                // Apply font color with validation
                if (rule.fontColor && isValidHexColor(rule.fontColor)) {
                    cf.cellValue.format.font.color = rule.fontColor;
                }
                if (rule.bold) cf.cellValue.format.font.bold = rule.bold;
                
                cf.cellValue.rule = {
                    formula1: String(rule.value),
                    formula2: rule.value2 ? String(rule.value2) : undefined,
                    operator: operator
                };
            }
            
            // ========== Color Scale ==========
            else if (ruleType === "colorScale" && rule.minimum && rule.maximum) {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                
                const mapCriterionType = (type) => {
                    const typeMap = {
                        "lowestValue": Excel.ConditionalFormatColorCriterionType.lowestValue,
                        "highestValue": Excel.ConditionalFormatColorCriterionType.highestValue,
                        "number": Excel.ConditionalFormatColorCriterionType.number,
                        "percent": Excel.ConditionalFormatColorCriterionType.percent,
                        "percentile": Excel.ConditionalFormatColorCriterionType.percentile,
                        "formula": Excel.ConditionalFormatColorCriterionType.formula
                    };
                    return typeMap[type] || Excel.ConditionalFormatColorCriterionType.lowestValue;
                };
                
                const criteria = {
                    minimum: {
                        type: mapCriterionType(rule.minimum.type),
                        color: rule.minimum.color || "#63BE7B",
                        formula: rule.minimum.formula || null
                    },
                    maximum: {
                        type: mapCriterionType(rule.maximum.type),
                        color: rule.maximum.color || "#F8696B",
                        formula: rule.maximum.formula || null
                    }
                };
                
                if (rule.midpoint) {
                    criteria.midpoint = {
                        type: mapCriterionType(rule.midpoint.type),
                        color: rule.midpoint.color || "#FFEB84",
                        formula: rule.midpoint.formula || null
                    };
                }
                
                cf.colorScale.criteria = criteria;
            }
            
            // ========== Data Bar ==========
            else if (ruleType === "dataBar") {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                
                if (rule.barDirection) {
                    const directionMap = {
                        "Context": Excel.ConditionalDataBarDirection.context,
                        "LeftToRight": Excel.ConditionalDataBarDirection.leftToRight,
                        "RightToLeft": Excel.ConditionalDataBarDirection.rightToLeft
                    };
                    cf.dataBar.barDirection = directionMap[rule.barDirection] || Excel.ConditionalDataBarDirection.context;
                }
                
                if (rule.showDataBarOnly !== undefined) cf.dataBar.showDataBarOnly = rule.showDataBarOnly;
                
                if (rule.positiveFormat) {
                    if (rule.positiveFormat.fillColor) cf.dataBar.positiveFormat.fillColor = rule.positiveFormat.fillColor;
                    if (rule.positiveFormat.borderColor) cf.dataBar.positiveFormat.borderColor = rule.positiveFormat.borderColor;
                    if (rule.positiveFormat.gradientFill !== undefined) cf.dataBar.positiveFormat.gradientFill = rule.positiveFormat.gradientFill;
                }
                
                if (rule.negativeFormat) {
                    if (rule.negativeFormat.fillColor) cf.dataBar.negativeFormat.fillColor = rule.negativeFormat.fillColor;
                    if (rule.negativeFormat.borderColor) cf.dataBar.negativeFormat.borderColor = rule.negativeFormat.borderColor;
                }
                
                if (rule.axisColor) cf.dataBar.axisColor = rule.axisColor;
            }
            
            // ========== Icon Set ==========
            else if (ruleType === "iconSet" && rule.style && validIconSets.includes(rule.style)) {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
                
                const styleMap = {
                    "threeArrows": Excel.IconSet.threeArrows,
                    "threeArrowsGray": Excel.IconSet.threeArrowsGray,
                    "threeTriangles": Excel.IconSet.threeTriangles,
                    "threeFlags": Excel.IconSet.threeFlags,
                    "threeTrafficLights1": Excel.IconSet.threeTrafficLights1,
                    "threeTrafficLights2": Excel.IconSet.threeTrafficLights2,
                    "threeSigns": Excel.IconSet.threeSigns,
                    "threeSymbols": Excel.IconSet.threeSymbols,
                    "threeSymbols2": Excel.IconSet.threeSymbols2,
                    "threeStars": Excel.IconSet.threeStars,
                    "fourArrows": Excel.IconSet.fourArrows,
                    "fourArrowsGray": Excel.IconSet.fourArrowsGray,
                    "fourRedToBlack": Excel.IconSet.fourRedToBlack,
                    "fourRating": Excel.IconSet.fourRating,
                    "fourTrafficLights": Excel.IconSet.fourTrafficLights,
                    "fiveArrows": Excel.IconSet.fiveArrows,
                    "fiveArrowsGray": Excel.IconSet.fiveArrowsGray,
                    "fiveRating": Excel.IconSet.fiveRating,
                    "fiveQuarters": Excel.IconSet.fiveQuarters,
                    "fiveBoxes": Excel.IconSet.fiveBoxes
                };
                
                cf.iconSet.style = styleMap[rule.style];
                
                // Determine expected criteria count based on icon set style
                const threeIconSets = ["threeArrows", "threeArrowsGray", "threeTriangles", "threeFlags", "threeTrafficLights1", "threeTrafficLights2", "threeSigns", "threeSymbols", "threeSymbols2", "threeStars"];
                const fourIconSets = ["fourArrows", "fourArrowsGray", "fourRedToBlack", "fourRating", "fourTrafficLights"];
                
                let expectedCriteriaCount = 3;
                if (fourIconSets.includes(rule.style)) expectedCriteriaCount = 4;
                else if (!threeIconSets.includes(rule.style)) expectedCriteriaCount = 5;
                
                if (rule.criteria && Array.isArray(rule.criteria)) {
                    // Validate criteria count matches icon count
                    if (rule.criteria.length !== expectedCriteriaCount) {
                        console.warn(`iconSet criteria count mismatch: expected ${expectedCriteriaCount} for ${rule.style}, got ${rule.criteria.length}`);
                        continue;
                    }
                    
                    const criteriaArray = rule.criteria.map(c => {
                        if (!c || Object.keys(c).length === 0) return {};
                        
                        const criterionTypeMap = {
                            "number": Excel.ConditionalFormatIconRuleType.number,
                            "percent": Excel.ConditionalFormatIconRuleType.percent,
                            "percentile": Excel.ConditionalFormatIconRuleType.percentile,
                            "formula": Excel.ConditionalFormatIconRuleType.formula
                        };
                        
                        const operatorMap = {
                            "greaterThan": Excel.ConditionalIconCriterionOperator.greaterThan,
                            "greaterThanOrEqual": Excel.ConditionalIconCriterionOperator.greaterThanOrEqual
                        };
                        
                        return {
                            type: criterionTypeMap[c.type] || Excel.ConditionalFormatIconRuleType.percent,
                            operator: operatorMap[c.operator] || Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
                            formula: c.formula || "0"
                        };
                    });
                    cf.iconSet.criteria = criteriaArray;
                }
                
                if (rule.showIconOnly !== undefined) cf.iconSet.showIconOnly = rule.showIconOnly;
                if (rule.reverseIconOrder !== undefined) cf.iconSet.reverseIconOrder = rule.reverseIconOrder;
            }
            
            // ========== Top/Bottom ==========
            else if (ruleType === "topBottom" && rule.rule && rule.rank !== undefined) {
                // Validate rank is a positive integer
                const rank = parseInt(rule.rank);
                if (!Number.isInteger(rank) || rank <= 0) {
                    console.warn(`Invalid topBottom rank: ${rule.rank}. Rank must be a positive integer.`);
                    continue;
                }
                
                // Map rule type and validate
                const ruleTypeMap = {
                    "TopItems": Excel.ConditionalTopBottomCriterionType.topItems,
                    "BottomItems": Excel.ConditionalTopBottomCriterionType.bottomItems,
                    "TopPercent": Excel.ConditionalTopBottomCriterionType.topPercent,
                    "BottomPercent": Excel.ConditionalTopBottomCriterionType.bottomPercent
                };
                
                const mappedRuleType = ruleTypeMap[rule.rule];
                if (!mappedRuleType) {
                    console.warn(`Invalid topBottom rule type: ${rule.rule}`);
                    continue;
                }
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
                
                cf.topBottom.rule = {
                    type: mappedRuleType,
                    rank: rank
                };
                
                // Apply formatting with color validation
                if (rule.fill && isValidHexColor(rule.fill)) cf.topBottom.format.fill.color = rule.fill;
                if (rule.fontColor && isValidHexColor(rule.fontColor)) cf.topBottom.format.font.color = rule.fontColor;
                if (rule.bold) cf.topBottom.format.font.bold = rule.bold;
            }
            
            // ========== Preset ==========
            else if (ruleType === "preset" && rule.criterion && validPresetCriteria.includes(rule.criterion)) {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
                
                const criterionMap = {
                    "duplicateValues": Excel.ConditionalFormatPresetCriterion.duplicateValues,
                    "uniqueValues": Excel.ConditionalFormatPresetCriterion.uniqueValues,
                    "aboveAverage": Excel.ConditionalFormatPresetCriterion.aboveAverage,
                    "belowAverage": Excel.ConditionalFormatPresetCriterion.belowAverage,
                    "equalOrAboveAverage": Excel.ConditionalFormatPresetCriterion.equalOrAboveAverage,
                    "equalOrBelowAverage": Excel.ConditionalFormatPresetCriterion.equalOrBelowAverage,
                    "oneStdDevAboveAverage": Excel.ConditionalFormatPresetCriterion.oneStdDevAboveAverage,
                    "oneStdDevBelowAverage": Excel.ConditionalFormatPresetCriterion.oneStdDevBelowAverage,
                    "twoStdDevAboveAverage": Excel.ConditionalFormatPresetCriterion.twoStdDevAboveAverage,
                    "twoStdDevBelowAverage": Excel.ConditionalFormatPresetCriterion.twoStdDevBelowAverage,
                    "threeStdDevAboveAverage": Excel.ConditionalFormatPresetCriterion.threeStdDevAboveAverage,
                    "threeStdDevBelowAverage": Excel.ConditionalFormatPresetCriterion.threeStdDevBelowAverage,
                    "yesterday": Excel.ConditionalFormatPresetCriterion.yesterday,
                    "today": Excel.ConditionalFormatPresetCriterion.today,
                    "tomorrow": Excel.ConditionalFormatPresetCriterion.tomorrow,
                    "lastSevenDays": Excel.ConditionalFormatPresetCriterion.lastSevenDays,
                    "lastWeek": Excel.ConditionalFormatPresetCriterion.lastWeek,
                    "thisWeek": Excel.ConditionalFormatPresetCriterion.thisWeek,
                    "nextWeek": Excel.ConditionalFormatPresetCriterion.nextWeek,
                    "lastMonth": Excel.ConditionalFormatPresetCriterion.lastMonth,
                    "thisMonth": Excel.ConditionalFormatPresetCriterion.thisMonth,
                    "nextMonth": Excel.ConditionalFormatPresetCriterion.nextMonth
                };
                
                cf.preset.rule = { criterion: criterionMap[rule.criterion] };
                
                // Apply formatting with color validation
                if (rule.fill && isValidHexColor(rule.fill)) cf.preset.format.fill.color = rule.fill;
                if (rule.fontColor && isValidHexColor(rule.fontColor)) cf.preset.format.font.color = rule.fontColor;
                if (rule.bold) cf.preset.format.font.bold = rule.bold;
            }
            
            // ========== Text Comparison ==========
            else if (ruleType === "textComparison" && rule.operator && rule.text) {
                const validOperators = ["contains", "notContains", "beginsWith", "endsWith"];
                if (!validOperators.includes(rule.operator)) continue;
                
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);
                
                const operatorMap = {
                    "contains": Excel.ConditionalTextOperator.contains,
                    "notContains": Excel.ConditionalTextOperator.notContains,
                    "beginsWith": Excel.ConditionalTextOperator.beginsWith,
                    "endsWith": Excel.ConditionalTextOperator.endsWith
                };
                
                cf.textComparison.rule = {
                    operator: operatorMap[rule.operator],
                    text: rule.text
                };
                
                // Apply formatting with color validation
                if (rule.fill && isValidHexColor(rule.fill)) cf.textComparison.format.fill.color = rule.fill;
                if (rule.fontColor && isValidHexColor(rule.fontColor)) cf.textComparison.format.font.color = rule.fontColor;
                if (rule.bold) cf.textComparison.format.font.bold = rule.bold;
            }
            
            // ========== Custom Formula ==========
            else if (ruleType === "custom" && rule.formula && rule.formula.startsWith("=")) {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
                
                cf.custom.rule = { formula: rule.formula };
                
                // Apply formatting with color validation
                if (rule.fill && isValidHexColor(rule.fill)) cf.custom.format.fill.color = rule.fill;
                if (rule.fontColor && isValidHexColor(rule.fontColor)) cf.custom.format.font.color = rule.fontColor;
                if (rule.bold) cf.custom.format.font.bold = rule.bold;
                if (rule.italic) cf.custom.format.font.italic = rule.italic;
            }
            
        } catch (e) { console.warn("Conditional format error:", e); }
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

async function createChart(ctx, sheet, dataRange, action) {
    const { chartType, data } = action;
    const ct = (chartType || "column").toLowerCase();
    
    // Load data to analyze it
    dataRange.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await ctx.sync();
    
    const values = dataRange.values;
    const headers = values[0];
    const rowCount = dataRange.rowCount;
    
    // Parse additional options from data if provided
    let title = "Chart";
    let position = "H2";
    
    // Try to extract title and position from action attributes or data
    if (action.title) title = action.title;
    if (action.position) position = action.position;
    
    // SMART DETECTION: If data has many rows with text categories, aggregate it
    let shouldAggregate = false;
    let categoryCol = -1;
    let valueCol = -1;
    
    // Check if we have a text column (categories) and need to count/sum
    if (rowCount > 10 && headers.length >= 2) {
        // Find first text column (likely category)
        for (let c = 0; c < headers.length; c++) {
            const sample = values.slice(1, Math.min(6, values.length)).map(r => r[c]);
            const hasText = sample.some(v => typeof v === "string" && v.length > 0);
            const hasRepeats = new Set(sample).size < sample.length;
            if (hasText && hasRepeats) {
                categoryCol = c;
                break;
            }
        }
        
        // Find meaningful numeric column (not IDs)
        for (let c = 0; c < headers.length; c++) {
            if (c === categoryCol) continue;
            
            const header = String(headers[c] || "").toLowerCase();
            const sample = values.slice(1, Math.min(10, values.length)).map(r => r[c]);
            const hasNumbers = sample.every(v => typeof v === "number" || !isNaN(parseFloat(v)));
            
            if (!hasNumbers) continue;
            
            // Skip if it looks like an ID column (sequential, unique, or has "id" in name)
            const isID = header.includes("id") || header.includes("no") || header.includes("number");
            const numericSample = sample.map(v => parseFloat(v)).filter(v => !isNaN(v));
            const isSequential = numericSample.length > 3 && 
                numericSample.every((v, i) => i === 0 || v > numericSample[i-1]);
            const isUnique = new Set(numericSample).size === numericSample.length;
            
            // Skip ID-like columns
            if (isID || (isSequential && isUnique)) continue;
            
            // This looks like a meaningful numeric column
            valueCol = c;
            break;
        }
        
        shouldAggregate = categoryCol !== -1;
    }
    
    // If we should aggregate, do it
    if (shouldAggregate) {
        const aggregated = {};
        for (let r = 1; r < values.length; r++) {
            const key = String(values[r][categoryCol] || "").trim();
            if (!key) continue;
            if (!aggregated[key]) aggregated[key] = { count: 0, sum: 0 };
            aggregated[key].count++;
            if (valueCol !== -1) {
                const val = parseFloat(values[r][valueCol]);
                if (!isNaN(val)) aggregated[key].sum += val;
            }
        }
        
        // Create aggregated data
        const aggData = Object.entries(aggregated)
            .map(([key, data]) => [key, valueCol !== -1 ? data.sum : data.count])
            .sort((a, b) => b[1] - a[1]);
        
        // Write aggregated data below original
        const aggStartRow = dataRange.rowIndex + rowCount + 2;
        const aggValues = [[headers[categoryCol] || "Category", valueCol !== -1 ? headers[valueCol] : "Count"], ...aggData];
        const aggRange = sheet.getRangeByIndexes(aggStartRow, dataRange.columnIndex, aggValues.length, 2);
        aggRange.values = aggValues;
        await ctx.sync();
        
        // Use aggregated data for chart
        dataRange = aggRange;
    }
    
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
    
    // Comment 4: Check if target contains comma (non-contiguous) - not supported directly
    if (targetAddress && targetAddress.includes(",")) {
        // Log warning and show user feedback
        console.warn("Non-contiguous ranges not fully supported for charts, using first range only");
        logWarn(`Chart: Non-contiguous range "${targetAddress}" - using first range only`);
        toast("Non-contiguous ranges not supported for charts; only the first range was used");
        
        // For non-contiguous ranges, we need to use the first range
        const ranges = targetAddress.split(",").map(r => r.trim());
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
    
    console.log(`Created ${ct} chart at ${position}`);
    
    // Parse advanced chart options from action.data
    // Supports both JSON string (from AI-generated ACTION tags) and plain objects (programmatic calls)
    let advancedOptions = {};
    if (action.data) {
        if (typeof action.data === "string") {
            try {
                advancedOptions = JSON.parse(action.data);
            } catch (e) {
                console.log(`Warning: Could not parse advanced chart options: ${e.message}`);
            }
        } else if (typeof action.data === "object") {
            advancedOptions = action.data;
        }
    }
    
    // Check if any series-level operations are needed (trendlines, dataLabels, comboSeries)
    const needsSeriesLoad = (advancedOptions.trendlines && Array.isArray(advancedOptions.trendlines) && advancedOptions.trendlines.length > 0) ||
                           advancedOptions.dataLabels ||
                           (advancedOptions.comboSeries && Array.isArray(advancedOptions.comboSeries) && advancedOptions.comboSeries.length > 0);
    
    // Load series once if needed for any series-level operations
    if (needsSeriesLoad) {
        chart.series.load("items");
        await ctx.sync();
    }
    
    // ========== Trendline Support ==========
    if (advancedOptions.trendlines && Array.isArray(advancedOptions.trendlines) && advancedOptions.trendlines.length > 0) {
        try {
            const trendlineTypeMap = {
                "Linear": Excel.ChartTrendlineType.linear,
                "Exponential": Excel.ChartTrendlineType.exponential,
                "Polynomial": Excel.ChartTrendlineType.polynomial,
                "MovingAverage": Excel.ChartTrendlineType.movingAverage
            };
            
            for (const trendlineConfig of advancedOptions.trendlines) {
                const seriesIndex = trendlineConfig.seriesIndex || 0;
                const trendlineType = trendlineConfig.type || "Linear";
                
                if (seriesIndex >= 0 && seriesIndex < chart.series.items.length) {
                    const series = chart.series.items[seriesIndex];
                    const trendline = series.trendlines.add(trendlineTypeMap[trendlineType] || Excel.ChartTrendlineType.linear);
                    
                    if (trendlineType === "MovingAverage" && trendlineConfig.period) {
                        trendline.movingAveragePeriod = trendlineConfig.period;
                    }
                    if (trendlineType === "Polynomial" && trendlineConfig.order) {
                        trendline.polynomialOrder = trendlineConfig.order;
                    }
                    
                    console.log(`Added ${trendlineType} trendline to series ${seriesIndex}`);
                } else {
                    console.log(`Warning: Invalid seriesIndex ${seriesIndex} for trendline, skipping`);
                }
            }
        } catch (e) {
            console.log(`Warning: Trendline customization error: ${e.message}`);
        }
    }
    
    // ========== Data Label Customization ==========
    if (advancedOptions.dataLabels) {
        try {
            
            const dataLabelPositionMap = {
                "Center": Excel.ChartDataLabelPosition.center,
                "InsideEnd": Excel.ChartDataLabelPosition.insideEnd,
                "OutsideEnd": Excel.ChartDataLabelPosition.outsideEnd,
                "InsideBase": Excel.ChartDataLabelPosition.insideBase,
                "BestFit": Excel.ChartDataLabelPosition.bestFit,
                "Left": Excel.ChartDataLabelPosition.left,
                "Right": Excel.ChartDataLabelPosition.right,
                "Top": Excel.ChartDataLabelPosition.top,
                "Bottom": Excel.ChartDataLabelPosition.bottom
            };
            
            for (const series of chart.series.items) {
                series.hasDataLabels = true;
                const labels = series.dataLabels;
                
                if (advancedOptions.dataLabels.position && dataLabelPositionMap[advancedOptions.dataLabels.position]) {
                    labels.position = dataLabelPositionMap[advancedOptions.dataLabels.position];
                }
                if (advancedOptions.dataLabels.showValue !== undefined) {
                    labels.showValue = advancedOptions.dataLabels.showValue;
                }
                if (advancedOptions.dataLabels.showSeriesName !== undefined) {
                    labels.showSeriesName = advancedOptions.dataLabels.showSeriesName;
                }
                if (advancedOptions.dataLabels.showCategoryName !== undefined) {
                    labels.showCategoryName = advancedOptions.dataLabels.showCategoryName;
                }
                if (advancedOptions.dataLabels.showLegendKey !== undefined) {
                    labels.showLegendKey = advancedOptions.dataLabels.showLegendKey;
                }
                if (advancedOptions.dataLabels.showPercentage !== undefined) {
                    labels.showPercentage = advancedOptions.dataLabels.showPercentage;
                }
                if (advancedOptions.dataLabels.numberFormat) {
                    labels.numberFormat = advancedOptions.dataLabels.numberFormat;
                }
                
                // Font formatting for data labels
                if (advancedOptions.dataLabels.format && advancedOptions.dataLabels.format.font) {
                    const font = advancedOptions.dataLabels.format.font;
                    if (font.bold !== undefined) labels.format.font.bold = font.bold;
                    if (font.color) labels.format.font.color = font.color;
                    if (font.size) labels.format.font.size = font.size;
                }
            }
            
            console.log(`Applied data labels: position=${advancedOptions.dataLabels.position || 'default'}`);
        } catch (e) {
            console.log(`Warning: Data label customization error: ${e.message}`);
        }
    }
    
    // ========== Axis Formatting ==========
    if (advancedOptions.axes) {
        try {
            const displayUnitMap = {
                "Hundreds": Excel.ChartAxisDisplayUnit.hundreds,
                "Thousands": Excel.ChartAxisDisplayUnit.thousands,
                "TenThousands": Excel.ChartAxisDisplayUnit.tenThousands,
                "HundredThousands": Excel.ChartAxisDisplayUnit.hundredThousands,
                "Millions": Excel.ChartAxisDisplayUnit.millions,
                "TenMillions": Excel.ChartAxisDisplayUnit.tenMillions,
                "HundredMillions": Excel.ChartAxisDisplayUnit.hundredMillions,
                "Billions": Excel.ChartAxisDisplayUnit.billions
            };
            
            // Category axis (X-axis)
            if (advancedOptions.axes.category) {
                const catAxis = chart.axes.categoryAxis;
                const catConfig = advancedOptions.axes.category;
                
                if (catConfig.title) {
                    catAxis.title.text = catConfig.title;
                    catAxis.title.visible = true;
                }
                if (catConfig.gridlines !== undefined) {
                    catAxis.majorGridlines.visible = catConfig.gridlines;
                }
                if (catConfig.format && catConfig.format.font) {
                    const font = catConfig.format.font;
                    if (font.bold !== undefined) catAxis.format.font.bold = font.bold;
                    if (font.color) catAxis.format.font.color = font.color;
                    if (font.size) catAxis.format.font.size = font.size;
                }
                
                console.log(`Applied category axis formatting: title="${catConfig.title || 'none'}"`);
            }
            
            // Value axis (Y-axis)
            if (advancedOptions.axes.value) {
                const valAxis = chart.axes.valueAxis;
                const valConfig = advancedOptions.axes.value;
                
                if (valConfig.title) {
                    valAxis.title.text = valConfig.title;
                    valAxis.title.visible = true;
                }
                if (valConfig.displayUnit && displayUnitMap[valConfig.displayUnit]) {
                    valAxis.displayUnit = displayUnitMap[valConfig.displayUnit];
                }
                if (valConfig.gridlines !== undefined) {
                    valAxis.majorGridlines.visible = valConfig.gridlines;
                }
                if (valConfig.minimum !== undefined) {
                    valAxis.minimum = valConfig.minimum;
                }
                if (valConfig.maximum !== undefined) {
                    valAxis.maximum = valConfig.maximum;
                }
                if (valConfig.format && valConfig.format.font) {
                    const font = valConfig.format.font;
                    if (font.bold !== undefined) valAxis.format.font.bold = font.bold;
                    if (font.color) valAxis.format.font.color = font.color;
                    if (font.size) valAxis.format.font.size = font.size;
                }
                
                console.log(`Applied value axis formatting: title="${valConfig.title || 'none'}", displayUnit="${valConfig.displayUnit || 'none'}"`);
            }
        } catch (e) {
            console.log(`Warning: Axis formatting error: ${e.message}`);
        }
    }
    
    // ========== Chart Element Formatting ==========
    if (advancedOptions.formatting) {
        try {
            // Title formatting
            if (advancedOptions.formatting.title && advancedOptions.formatting.title.font) {
                const font = advancedOptions.formatting.title.font;
                if (font.bold !== undefined) chart.title.format.font.bold = font.bold;
                if (font.color) chart.title.format.font.color = font.color;
                if (font.size) chart.title.format.font.size = font.size;
                if (font.italic !== undefined) chart.title.format.font.italic = font.italic;
                
                console.log(`Applied title formatting: bold=${font.bold}, color=${font.color}, size=${font.size}`);
            }
            
            // Legend formatting
            if (advancedOptions.formatting.legend) {
                const legendConfig = advancedOptions.formatting.legend;
                
                const legendPositionMap = {
                    "Top": Excel.ChartLegendPosition.top,
                    "Bottom": Excel.ChartLegendPosition.bottom,
                    "Left": Excel.ChartLegendPosition.left,
                    "Right": Excel.ChartLegendPosition.right,
                    "Corner": Excel.ChartLegendPosition.corner,
                    "Custom": Excel.ChartLegendPosition.custom
                };
                
                if (legendConfig.position && legendPositionMap[legendConfig.position]) {
                    chart.legend.position = legendPositionMap[legendConfig.position];
                }
                if (legendConfig.font) {
                    if (legendConfig.font.bold !== undefined) chart.legend.format.font.bold = legendConfig.font.bold;
                    if (legendConfig.font.color) chart.legend.format.font.color = legendConfig.font.color;
                    if (legendConfig.font.size) chart.legend.format.font.size = legendConfig.font.size;
                }
                
                console.log(`Applied legend formatting: position=${legendConfig.position || 'default'}`);
            }
            
            // Chart area formatting (fill and border)
            if (advancedOptions.formatting.chartArea) {
                if (advancedOptions.formatting.chartArea.fill) {
                    chart.format.fill.setSolidColor(advancedOptions.formatting.chartArea.fill);
                    console.log(`Applied chart area fill: ${advancedOptions.formatting.chartArea.fill}`);
                }
                
                // Chart area border customization
                if (advancedOptions.formatting.chartArea.border) {
                    const borderConfig = advancedOptions.formatting.chartArea.border;
                    const chartLine = chart.format.border;
                    
                    if (borderConfig.color) {
                        chartLine.color = borderConfig.color;
                    }
                    if (borderConfig.weight !== undefined) {
                        chartLine.weight = borderConfig.weight;
                    }
                    if (borderConfig.lineStyle) {
                        const lineStyleMap = {
                            "Continuous": Excel.ChartLineStyle.continuous,
                            "Dash": Excel.ChartLineStyle.dash,
                            "DashDot": Excel.ChartLineStyle.dashDot,
                            "DashDotDot": Excel.ChartLineStyle.dashDotDot,
                            "Dot": Excel.ChartLineStyle.dot,
                            "Grey25": Excel.ChartLineStyle.grey25,
                            "Grey50": Excel.ChartLineStyle.grey50,
                            "Grey75": Excel.ChartLineStyle.grey75,
                            "Automatic": Excel.ChartLineStyle.automatic,
                            "None": Excel.ChartLineStyle.none
                        };
                        if (lineStyleMap[borderConfig.lineStyle]) {
                            chartLine.lineStyle = lineStyleMap[borderConfig.lineStyle];
                        }
                    }
                    
                    console.log(`Applied chart area border: color=${borderConfig.color || 'default'}, weight=${borderConfig.weight || 'default'}, style=${borderConfig.lineStyle || 'default'}`);
                }
            }
            
            // Plot area formatting (fill and border)
            if (advancedOptions.formatting.plotArea) {
                const plotArea = chart.plotArea;
                
                if (advancedOptions.formatting.plotArea.fill) {
                    plotArea.format.fill.setSolidColor(advancedOptions.formatting.plotArea.fill);
                    console.log(`Applied plot area fill: ${advancedOptions.formatting.plotArea.fill}`);
                }
                
                if (advancedOptions.formatting.plotArea.border) {
                    const borderConfig = advancedOptions.formatting.plotArea.border;
                    const plotLine = plotArea.format.border;
                    
                    if (borderConfig.color) {
                        plotLine.color = borderConfig.color;
                    }
                    if (borderConfig.weight !== undefined) {
                        plotLine.weight = borderConfig.weight;
                    }
                    if (borderConfig.lineStyle) {
                        const lineStyleMap = {
                            "Continuous": Excel.ChartLineStyle.continuous,
                            "Dash": Excel.ChartLineStyle.dash,
                            "DashDot": Excel.ChartLineStyle.dashDot,
                            "DashDotDot": Excel.ChartLineStyle.dashDotDot,
                            "Dot": Excel.ChartLineStyle.dot,
                            "None": Excel.ChartLineStyle.none
                        };
                        if (lineStyleMap[borderConfig.lineStyle]) {
                            plotLine.lineStyle = lineStyleMap[borderConfig.lineStyle];
                        }
                    }
                    
                    console.log(`Applied plot area border: color=${borderConfig.color || 'default'}, weight=${borderConfig.weight || 'default'}`);
                }
            }
        } catch (e) {
            console.log(`Warning: Chart element formatting error: ${e.message}`);
        }
    }
    
    // ========== Combo Chart / Secondary Axis Support ==========
    // Note: Series already loaded above if needsSeriesLoad was true
    if (advancedOptions.comboSeries && Array.isArray(advancedOptions.comboSeries) && advancedOptions.comboSeries.length > 0) {
        try {
            const comboChartTypeMap = {
                "Line": Excel.ChartType.line,
                "ColumnClustered": Excel.ChartType.columnClustered,
                "ColumnStacked": Excel.ChartType.columnStacked,
                "Area": Excel.ChartType.area,
                "AreaStacked": Excel.ChartType.areaStacked,
                "Scatter": Excel.ChartType.xyscatter
            };
            
            const axisGroupMap = {
                "Primary": Excel.ChartAxisGroup.primary,
                "Secondary": Excel.ChartAxisGroup.secondary
            };
            
            for (const comboConfig of advancedOptions.comboSeries) {
                const seriesIndex = comboConfig.seriesIndex;
                
                if (seriesIndex >= 0 && seriesIndex < chart.series.items.length) {
                    const series = chart.series.items[seriesIndex];
                    
                    if (comboConfig.chartType && comboChartTypeMap[comboConfig.chartType]) {
                        series.chartType = comboChartTypeMap[comboConfig.chartType];
                    }
                    if (comboConfig.axisGroup && axisGroupMap[comboConfig.axisGroup]) {
                        series.axisGroup = axisGroupMap[comboConfig.axisGroup];
                    }
                    
                    console.log(`Set series ${seriesIndex} to ${comboConfig.chartType || 'default'} on ${comboConfig.axisGroup || 'Primary'} axis`);
                } else {
                    console.log(`Warning: Invalid seriesIndex ${seriesIndex} for combo series, skipping`);
                }
            }
            
            // Configure secondary value axis if any series uses it
            if (advancedOptions.axes && advancedOptions.axes.value2) {
                try {
                    const secValAxis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
                    const val2Config = advancedOptions.axes.value2;
                    
                    if (val2Config.title) {
                        secValAxis.title.text = val2Config.title;
                        secValAxis.title.visible = true;
                    }
                    
                    console.log(`Applied secondary value axis title: "${val2Config.title || 'none'}"`);
                } catch (secAxisError) {
                    console.log(`Warning: Secondary axis configuration error: ${secAxisError.message}`);
                }
            }
        } catch (e) {
            console.log(`Warning: Combo chart customization error: ${e.message}`);
        }
    }
    
    // Final sync to apply all chart customizations
    await ctx.sync();
}

/**
 * Creates a pivot chart by aggregating data intelligently
 */
async function createPivotChart(ctx, sheet, range, action) {
    range.load(["values", "rowIndex", "columnIndex", "rowCount"]);
    await ctx.sync();
    
    const values = range.values;
    const headers = values[0];
    
    // Parse options
    let options = { groupBy: null, aggregate: null, aggregateFunc: "sum", chartType: "column", title: "Pivot Chart", position: "H2" };
    if (action.data) {
        try { options = { ...options, ...JSON.parse(action.data) }; } catch (e) {}
    }
    if (action.chartType) options.chartType = action.chartType;
    if (action.title) options.title = action.title;
    if (action.position) options.position = action.position;
    
    // Find columns - be more flexible with matching
    let groupByIdx = -1;
    for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i]).toLowerCase().trim();
        const searchTerm = String(options.groupBy || "").toLowerCase().trim();
        if (header === searchTerm || header.includes(searchTerm) || searchTerm.includes(header)) {
            groupByIdx = i;
            break;
        }
    }
    
    if (groupByIdx === -1) {
        console.error("Available headers:", headers);
        throw new Error(`Column "${options.groupBy}" not found. Available: ${headers.join(", ")}`);
    }
    
    let aggregateIdx = -1;
    if (options.aggregate) {
        for (let i = 0; i < headers.length; i++) {
            const header = String(headers[i]).toLowerCase().trim();
            const searchTerm = String(options.aggregate).toLowerCase().trim();
            if (header === searchTerm || header.includes(searchTerm) || searchTerm.includes(header)) {
                aggregateIdx = i;
                break;
            }
        }
    }
    
    // Aggregate data
    const aggregated = {};
    for (let r = 1; r < values.length; r++) {
        const groupValue = values[r][groupByIdx];
        const key = String(groupValue || "").trim();
        if (!key || key === "null" || key === "undefined") continue;
        
        if (!aggregated[key]) aggregated[key] = { count: 0, sum: 0, values: [] };
        aggregated[key].count++;
        
        if (aggregateIdx !== -1) {
            const val = parseFloat(values[r][aggregateIdx]);
            if (!isNaN(val)) {
                aggregated[key].sum += val;
                aggregated[key].values.push(val);
            }
        }
    }
    
    // Calculate final values
    const chartData = [];
    for (const [key, data] of Object.entries(aggregated)) {
        let value;
        const func = (options.aggregateFunc || "count").toLowerCase();
        switch (func) {
            case "count": 
                value = data.count; 
                break;
            case "average": 
            case "avg": 
                value = data.values.length > 0 ? data.sum / data.values.length : data.count; 
                break;
            case "max": 
                value = data.values.length > 0 ? Math.max(...data.values) : data.count; 
                break;
            case "min": 
                value = data.values.length > 0 ? Math.min(...data.values) : data.count; 
                break;
            case "sum":
            default: 
                value = data.values.length > 0 ? data.sum : data.count; 
                break;
        }
        chartData.push([key, value]);
    }
    chartData.sort((a, b) => b[1] - a[1]);
    
    // Write aggregated data
    const chartStartRow = range.rowIndex + range.rowCount + 2;
    const chartValues = [[options.groupBy || "Category", options.aggregate || "Value"], ...chartData];
    const chartDataRange = sheet.getRangeByIndexes(chartStartRow, range.columnIndex, chartValues.length, 2);
    chartDataRange.values = chartValues;
    await ctx.sync();
    
    // Create chart
    let type = Excel.ChartType.columnClustered;
    const ct = options.chartType.toLowerCase();
    if (ct.includes("pie")) type = Excel.ChartType.pie;
    else if (ct.includes("bar")) type = Excel.ChartType.barClustered;
    else if (ct.includes("line")) type = Excel.ChartType.line;
    
    const chart = sheet.charts.add(type, chartDataRange, Excel.ChartSeriesBy.columns);
    const position = options.position || "H2";
    const startCol = position.match(/[A-Z]+/)?.[0] || "H";
    const startRow = parseInt(position.match(/\d+/)?.[0] || "2");
    chart.setPosition(position, `${String.fromCharCode(startCol.charCodeAt(0) + 8)}${startRow + 15}`);
    chart.title.text = options.title;
    chart.legend.visible = true;
    chart.legend.position = ct.includes("pie") ? Excel.ChartLegendPosition.right : Excel.ChartLegendPosition.bottom;
    await ctx.sync();
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

/**
 * Removes duplicate rows from a range
 * @param {Object} ctx - Excel context
 * @param {Object} range - Excel range
 * @param {string} data - JSON string with columns array
 */
async function removeDuplicates(ctx, range, data) {
    // Load the range data and address
    range.load(["values", "rowCount", "columnCount", "address"]);
    await ctx.sync();
    
    const values = range.values;
    const rowCount = range.rowCount;
    const colCount = range.columnCount;
    const rangeAddress = range.address;
    
    // Parse options (which columns to check for duplicates)
    let options = { columns: [] };
    if (data) {
        try {
            options = JSON.parse(data);
        } catch (e) {
            // Default to all columns
            options.columns = Array.from({ length: colCount }, (_, i) => i);
        }
    }
    
    // If no columns specified, use all columns
    if (!options.columns || options.columns.length === 0) {
        options.columns = Array.from({ length: colCount }, (_, i) => i);
    }
    
    // Find unique rows (keep first occurrence)
    const seen = new Set();
    const uniqueRows = [];
    
    for (let r = 0; r < rowCount; r++) {
        const row = values[r];
        
        // Create a key from the specified columns
        const key = options.columns.map(colIdx => {
            const val = row[colIdx];
            return val === null || val === undefined ? "" : String(val);
        }).join("|");
        
        if (!seen.has(key)) {
            seen.add(key);
            uniqueRows.push(row);
        }
    }
    
    // Clear the original range
    range.clear(Excel.ClearApplyTo.contents);
    await ctx.sync();
    
    // Write back only unique rows
    if (uniqueRows.length > 0) {
        // Get the sheet and create a new range reference from the original address
        const sheet = range.worksheet;
        const address = rangeAddress.split("!")[1] || rangeAddress; // Remove sheet name if present
        const startCell = address.split(":")[0]; // Get starting cell (e.g., "A1")
        
        // Create a new range from the start cell
        const targetCell = sheet.getRange(startCell);
        const newRange = targetCell.getResizedRange(uniqueRows.length - 1, colCount - 1);
        newRange.values = uniqueRows;
    }
}

// ============================================================================
// Table Operations
// ============================================================================

/**
 * Valid table styles for validation
 */
const VALID_TABLE_STYLES = [
    // Light styles (1-21)
    ...Array.from({ length: 21 }, (_, i) => `TableStyleLight${i + 1}`),
    // Medium styles (1-28)
    ...Array.from({ length: 28 }, (_, i) => `TableStyleMedium${i + 1}`),
    // Dark styles (1-11)
    ...Array.from({ length: 11 }, (_, i) => `TableStyleDark${i + 1}`)
];

/**
 * Creates an Excel Table from a range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Excel.Range} range - Data range for the table
 * @param {Object} action - Action with table options
 */
async function createTable(ctx, sheet, range, action) {
    logDebug(`Starting createTable at range "${action.target}"`);
    
    let options = { tableName: null, style: "TableStyleMedium2" };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for createTable, using defaults`);
        }
    }
    
    // Validate style with clear error message
    if (options.style && !VALID_TABLE_STYLES.includes(options.style)) {
        logWarn(`Invalid table style "${options.style}". Valid styles: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11. Using TableStyleMedium2.`);
        options.style = "TableStyleMedium2";
    }
    
    try {
        // Create table with headers (true = first row is header)
        const table = sheet.tables.add(range, true);
        
        // Set table name if provided
        if (options.tableName) {
            table.name = options.tableName;
        }
        
        // Apply style
        table.style = options.style;
        
        // Enable default table features
        table.showBandedRows = true;
        table.showFilterButton = true;
        
        await ctx.sync();
        
        const tableName = options.tableName || table.name;
        logInfo(`Successfully created table "${tableName}" at ${action.target} with style ${options.style}`);
    } catch (e) {
        const errorMsg = e.message && e.message.includes("already") 
            ? `Failed to create table: Range ${action.target} already contains a table or overlaps with one.`
            : `Failed to create table at ${action.target}: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Applies styling to an existing table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {string} tableName - Name of the table to style
 * @param {string} data - JSON string with style options
 */
async function styleTable(ctx, sheet, tableName, data) {
    logDebug(`Starting styleTable for table "${tableName}"`);
    
    let options = { style: "TableStyleMedium2" };
    
    if (data) {
        try {
            options = { ...options, ...JSON.parse(data) };
        } catch (e) {
            logWarn(`Failed to parse data for styleTable, using defaults`);
        }
    }
    
    // Validate style with clear error message
    if (options.style && !VALID_TABLE_STYLES.includes(options.style)) {
        logWarn(`Invalid table style "${options.style}". Valid styles: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11. Using TableStyleMedium2.`);
        options.style = "TableStyleMedium2";
    }
    
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Apply style
        table.style = options.style;
        
        // Apply additional style options if provided
        if (options.highlightFirstColumn !== undefined) {
            table.highlightFirstColumn = options.highlightFirstColumn;
        }
        if (options.highlightLastColumn !== undefined) {
            table.highlightLastColumn = options.highlightLastColumn;
        }
        if (options.showBandedRows !== undefined) {
            table.showBandedRows = options.showBandedRows;
        }
        if (options.showBandedColumns !== undefined) {
            table.showBandedColumns = options.showBandedColumns;
        }
        
        await ctx.sync();
        logInfo(`Successfully applied style "${options.style}" to table "${tableName}"`);
    } catch (e) {
        const errorMsg = `Failed to style table "${tableName}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Adds a row to an existing table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with row options
 */
async function addTableRow(ctx, sheet, action) {
    logDebug(`Starting addTableRow for target "${action.target}"`);
    
    let options = { tableName: action.target, position: "end", values: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for addTableRow, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    // Determine index: null for end, 0 for start, or specific number
    let index = null;
    if (options.position === "start" || options.position === 0) {
        index = 0;
    } else if (typeof options.position === "number" && options.position > 0) {
        index = options.position;
    }
    // null means append to end
    
    // Prepare values - should be array of arrays
    let rowValues = null;
    if (options.values) {
        rowValues = Array.isArray(options.values[0]) ? options.values : [options.values];
    }
    
    try {
        table.rows.add(index, rowValues);
        await ctx.sync();
        logInfo(`Successfully added row to table "${tableName}" at position ${options.position || "end"}`);
    } catch (e) {
        const errorMsg = `Failed to add row to table "${tableName}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Adds a column to an existing table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with column options
 */
async function addTableColumn(ctx, sheet, action) {
    logDebug(`Starting addTableColumn for target "${action.target}"`);
    
    let options = { tableName: action.target, columnName: "NewColumn", position: "end", values: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for addTableColumn, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    // Determine index: null for end, 0 for start, or specific number
    let index = null;
    if (options.position === "start" || options.position === 0) {
        index = 0;
    } else if (typeof options.position === "number" && options.position > 0) {
        index = options.position;
    }
    
    // Prepare values - should include header as first element
    let columnValues = null;
    if (options.values) {
        columnValues = Array.isArray(options.values[0]) ? options.values : options.values.map(v => [v]);
    } else if (options.columnName) {
        // Just add header if no values provided
        columnValues = [[options.columnName]];
    }
    
    try {
        table.columns.add(index, columnValues);
        await ctx.sync();
        logInfo(`Successfully added column "${options.columnName}" to table "${tableName}" at position ${options.position || "end"}`);
    } catch (e) {
        const errorMsg = `Failed to add column to table "${tableName}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Resizes an existing table to a new range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with resize options
 */
async function resizeTable(ctx, sheet, action) {
    logDebug(`Starting resizeTable for target "${action.target}"`);
    
    let options = { tableName: action.target, newRange: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for resizeTable, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    
    if (!options.newRange) {
        const errorMsg = `newRange is required for resizeTable operation on table "${tableName}".`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    // Get current range for logging
    const currentRange = table.getRange();
    currentRange.load("address");
    await ctx.sync();
    
    const oldAddress = currentRange.address;
    
    try {
        // Resize the table
        table.resize(options.newRange);
        await ctx.sync();
        logInfo(`Successfully resized table "${tableName}" from ${oldAddress} to ${options.newRange}`);
    } catch (e) {
        const errorMsg = `Failed to resize table "${tableName}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Converts a table back to a normal range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {string} tableName - Name of the table to convert
 */
async function convertToRange(ctx, sheet, tableName) {
    logDebug(`Starting convertToRange for table "${tableName}"`);
    
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Convert table to range - preserves data and formatting
        table.convertToRange();
        await ctx.sync();
        logInfo(`Successfully converted table "${tableName}" to normal range`);
    } catch (e) {
        const errorMsg = `Failed to convert table "${tableName}" to range: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Toggles the total row for a table
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with totals options
 */
/**
 * Gets valid totals calculation functions map
 * Note: Must be called inside Excel.run context, not at module load time
 */
function getValidTotalsFunctions() {
    return {
        "sum": Excel.TotalsCalculation.sum,
        "average": Excel.TotalsCalculation.average,
        "avg": Excel.TotalsCalculation.average,
        "count": Excel.TotalsCalculation.count,
        "countnumbers": Excel.TotalsCalculation.countNumbers,
        "max": Excel.TotalsCalculation.max,
        "min": Excel.TotalsCalculation.min,
        "stddev": Excel.TotalsCalculation.stdDev,
        "var": Excel.TotalsCalculation.var,
        "none": Excel.TotalsCalculation.none
    };
}

async function toggleTableTotals(ctx, sheet, action) {
    logDebug(`Starting toggleTableTotals for target "${action.target}"`);
    
    let options = { tableName: action.target, show: true, totals: null };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for toggleTableTotals, using defaults`);
        }
    }
    
    const tableName = options.tableName || action.target;
    const table = sheet.tables.getItemOrNullObject(tableName);
    table.load(["name", "isNullObject"]);
    await ctx.sync();
    
    if (table.isNullObject) {
        const errorMsg = `Table "${tableName}" not found on the active sheet.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    // Toggle totals row visibility
    table.showTotals = options.show;
    await ctx.sync();
    
    logDebug(`Set showTotals=${options.show} for table "${tableName}"`);
    
    // If enabling totals and specific functions are requested, apply them
    const appliedFunctions = [];
    if (options.show && options.totals && Array.isArray(options.totals) && options.totals.length > 0) {
        table.columns.load("count");
        await ctx.sync();
        
        const columnCount = table.columns.count;
        
        for (const totalConfig of options.totals) {
            // Validate columnIndex
            if (totalConfig.columnIndex === undefined || totalConfig.columnIndex === null) {
                logWarn(`Skipping totals config - missing columnIndex`);
                continue;
            }
            
            if (typeof totalConfig.columnIndex !== "number" || totalConfig.columnIndex < 0) {
                logWarn(`Skipping totals config - invalid columnIndex "${totalConfig.columnIndex}"`);
                continue;
            }
            
            if (totalConfig.columnIndex >= columnCount) {
                logWarn(`Skipping totals config - columnIndex ${totalConfig.columnIndex} exceeds table column count ${columnCount}`);
                continue;
            }
            
            // Validate function name
            if (!totalConfig.function) {
                logWarn(`Skipping totals config for column ${totalConfig.columnIndex} - missing function`);
                continue;
            }
            
            const funcName = String(totalConfig.function).toLowerCase().replace(/\s/g, "");
            const validFunctions = getValidTotalsFunctions();
            const calcFunc = validFunctions[funcName];
            
            if (!calcFunc) {
                logWarn(`Invalid totals function "${totalConfig.function}" for column ${totalConfig.columnIndex}. Valid functions: Sum, Average, Count, Max, Min, StdDev, Var, None`);
                continue;
            }
            
            try {
                const column = table.columns.getItemAt(totalConfig.columnIndex);
                column.totalsCalculation = calcFunc;
                appliedFunctions.push(`column ${totalConfig.columnIndex}: ${totalConfig.function}`);
            } catch (e) {
                logWarn(`Failed to apply ${totalConfig.function} to column ${totalConfig.columnIndex}: ${e.message}`);
            }
        }
        
        await ctx.sync();
    }
    
    if (appliedFunctions.length > 0) {
        logInfo(`Applied totals functions for table "${tableName}": ${appliedFunctions.join(", ")}`);
    }
    
    logInfo(`Completed toggleTableTotals for table "${tableName}": show=${options.show}`);
}

// ============================================================================
// Data Manipulation Operations
// ============================================================================

/**
 * Inserts rows at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with row options
 */
async function insertRows(ctx, sheet, action) {
    logDebug(`Starting insertRows for target "${action.target}"`);
    
    let options = { count: 1 };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for insertRows, using defaults`);
        }
    }
    
    // Validate target is a row range (e.g., "5" or "5:7")
    const rowPattern = /^(\d+)(:\d+)?$/;
    if (!rowPattern.test(action.target)) {
        const errorMsg = `Invalid row range "${action.target}". Use format "5" or "5:7".`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    // Validate count
    if (typeof options.count !== "number" || options.count < 1) {
        logWarn(`Invalid count "${options.count}", using 1`);
        options.count = 1;
    }
    
    try {
        const range = sheet.getRange(`${action.target}:${action.target}`);
        const entireRow = range.getEntireRow();
        
        // Insert rows multiple times if count > 1
        for (let i = 0; i < options.count; i++) {
            entireRow.insert(Excel.InsertShiftDirection.down);
        }
        
        await ctx.sync();
        logInfo(`Successfully inserted ${options.count} row(s) at row ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to insert rows at "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Inserts columns at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with column options
 */
async function insertColumns(ctx, sheet, action) {
    logDebug(`Starting insertColumns for target "${action.target}"`);
    
    let options = { count: 1 };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for insertColumns, using defaults`);
        }
    }
    
    // Validate target is a column range (e.g., "C" or "C:E")
    const colPattern = /^([A-Z]+)(:[A-Z]+)?$/i;
    if (!colPattern.test(action.target)) {
        const errorMsg = `Invalid column range "${action.target}". Use format "C" or "C:E".`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    // Validate count
    if (typeof options.count !== "number" || options.count < 1) {
        logWarn(`Invalid count "${options.count}", using 1`);
        options.count = 1;
    }
    
    try {
        const range = sheet.getRange(`${action.target}:${action.target}`);
        const entireColumn = range.getEntireColumn();
        
        // Insert columns multiple times if count > 1
        for (let i = 0; i < options.count; i++) {
            entireColumn.insert(Excel.InsertShiftDirection.right);
        }
        
        await ctx.sync();
        logInfo(`Successfully inserted ${options.count} column(s) at column ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to insert columns at "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes rows at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with row options
 */
async function deleteRows(ctx, sheet, action) {
    logDebug(`Starting deleteRows for target "${action.target}"`);
    
    // Validate target is a row range (e.g., "10" or "10:15")
    const rowPattern = /^(\d+)(:\d+)?$/;
    if (!rowPattern.test(action.target)) {
        const errorMsg = `Invalid row range "${action.target}". Use format "10" or "10:15".`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        const range = sheet.getRange(`${action.target}:${action.target}`);
        const entireRow = range.getEntireRow();
        entireRow.delete(Excel.DeleteShiftDirection.up);
        
        await ctx.sync();
        logInfo(`Successfully deleted row(s) at ${action.target}. Warning: This may affect formula references.`);
    } catch (e) {
        const errorMsg = `Failed to delete rows at "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes columns at the specified position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with column options
 */
async function deleteColumns(ctx, sheet, action) {
    logDebug(`Starting deleteColumns for target "${action.target}"`);
    
    // Validate target is a column range (e.g., "D" or "D:F")
    const colPattern = /^([A-Z]+)(:[A-Z]+)?$/i;
    if (!colPattern.test(action.target)) {
        const errorMsg = `Invalid column range "${action.target}". Use format "D" or "D:F".`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        const range = sheet.getRange(`${action.target}:${action.target}`);
        const entireColumn = range.getEntireColumn();
        entireColumn.delete(Excel.DeleteShiftDirection.left);
        
        await ctx.sync();
        logInfo(`Successfully deleted column(s) at ${action.target}. Warning: This may affect formula references.`);
    } catch (e) {
        const errorMsg = `Failed to delete columns at "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Merges cells in the specified range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with merge options
 */
async function mergeCells(ctx, sheet, action) {
    logDebug(`Starting mergeCells for target "${action.target}"`);
    
    try {
        const range = sheet.getRange(action.target);
        range.load(["address", "rowCount", "columnCount"]);
        await ctx.sync();
        
        // Validate range is at least 2 cells
        if (range.rowCount === 1 && range.columnCount === 1) {
            const errorMsg = `Cannot merge a single cell. Range must contain at least 2 cells.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        range.merge(false);
        await ctx.sync();
        
        logInfo(`Successfully merged cells at ${action.target}. Note: Only the top-left cell value is retained.`);
    } catch (e) {
        if (e.message && e.message.includes("merge")) {
            const errorMsg = `Failed to merge cells at "${action.target}": Range may already contain merged cells.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        const errorMsg = `Failed to merge cells at "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Unmerges cells in the specified range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with unmerge options
 */
async function unmergeCells(ctx, sheet, action) {
    logDebug(`Starting unmergeCells for target "${action.target}"`);
    
    try {
        const range = sheet.getRange(action.target);
        range.unmerge();
        await ctx.sync();
        
        logInfo(`Successfully unmerged cells at ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to unmerge cells at "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Finds and replaces text in the specified range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with find/replace options
 */
async function findReplace(ctx, sheet, action) {
    logDebug(`Starting findReplace for target "${action.target}"`);
    
    let options = { find: "", replace: "", matchCase: false, matchEntireCell: false };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for findReplace`);
        }
    }
    
    // Validate find string
    if (!options.find || options.find.length === 0) {
        const errorMsg = `Find string cannot be empty.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        const range = sheet.getRange(action.target);
        
        const searchCriteria = {
            completeMatch: options.matchEntireCell,
            matchCase: options.matchCase
        };
        
        range.replaceAll(options.find, options.replace || "", searchCriteria);
        await ctx.sync();
        
        logInfo(`Successfully replaced "${options.find}" with "${options.replace}" in ${action.target}`);
    } catch (e) {
        const errorMsg = `Failed to find/replace in "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Splits text in a column into multiple columns based on delimiter
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with text-to-columns options
 */
async function textToColumns(ctx, sheet, action) {
    logDebug(`Starting textToColumns for target "${action.target}"`);
    
    let options = { delimiter: ",", destination: null };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for textToColumns, using defaults`);
        }
    }
    
    try {
        const sourceRange = sheet.getRange(action.target);
        sourceRange.load(["values", "rowCount", "columnCount", "columnIndex", "rowIndex"]);
        await ctx.sync();
        
        // Validate source is single column
        if (sourceRange.columnCount !== 1) {
            const errorMsg = `Text to columns requires a single-column range. Got ${sourceRange.columnCount} columns.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        const values = sourceRange.values;
        const delimiter = options.delimiter || ",";
        
        // Split each cell value
        const splitData = [];
        let maxColumns = 0;
        
        for (const row of values) {
            const cellValue = row[0];
            const parts = cellValue !== null && cellValue !== undefined 
                ? String(cellValue).split(delimiter) 
                : [""];
            splitData.push(parts);
            maxColumns = Math.max(maxColumns, parts.length);
        }
        
        // Pad arrays to same length
        for (const row of splitData) {
            while (row.length < maxColumns) {
                row.push("");
            }
        }
        
        // Determine destination
        let destRange;
        if (options.destination) {
            destRange = sheet.getRange(options.destination);
            destRange = destRange.getResizedRange(splitData.length - 1, maxColumns - 1);
        } else {
            // Use columns immediately to the right of source
            const destStartCol = sourceRange.columnIndex + 1;
            destRange = sheet.getRangeByIndexes(
                sourceRange.rowIndex,
                destStartCol,
                splitData.length,
                maxColumns
            );
        }
        
        destRange.values = splitData;
        await ctx.sync();
        
        logInfo(`Successfully split ${values.length} cells into ${maxColumns} columns. Warning: Adjacent data may have been overwritten.`);
    } catch (e) {
        const errorMsg = `Failed to split text to columns for "${action.target}": ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// Hyperlink Operations
// ============================================================================

// Cache for hyperlink API support check
let hyperlinkSupportChecked = false;
let hyperlinkSupported = false;

/**
 * Checks if the Range.hyperlink API is supported (ExcelApi 1.7+)
 * @param {Excel.RequestContext} ctx - Excel context
 * @returns {Promise<boolean>} True if hyperlinks are supported
 */
async function isHyperlinkSupported(ctx) {
    if (hyperlinkSupportChecked) {
        return hyperlinkSupported;
    }
    
    try {
        // Check using Office.context.requirements if available
        if (typeof Office !== 'undefined' && Office.context && Office.context.requirements) {
            hyperlinkSupported = Office.context.requirements.isSetSupported('ExcelApi', '1.7');
            hyperlinkSupportChecked = true;
            return hyperlinkSupported;
        }
        
        // Fallback: try a lightweight operation to test support
        const testRange = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1");
        testRange.load("hyperlink");
        await ctx.sync();
        hyperlinkSupported = true;
        hyperlinkSupportChecked = true;
        return true;
    } catch (e) {
        hyperlinkSupported = false;
        hyperlinkSupportChecked = true;
        return false;
    }
}

/**
 * Adds a hyperlink to a cell or range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string with hyperlink options
 * 
 * Supported options:
 * - url: Web URL (e.g., "https://example.com")
 * - email: Email address (automatically prefixed with "mailto:")
 * - documentReference: Internal link (e.g., "'Sheet2'!A1")
 * - displayText: Text to display in cell (defaults to URL/email/reference)
 * - tooltip: Hover tooltip text (screenTip)
 * 
 * Note: Only one of url, email, or documentReference should be provided.
 * Requires ExcelApi 1.7+ (Excel 2016+, Excel Online, Excel 365)
 */
async function addHyperlink(ctx, range, data) {
    console.log(`[addHyperlink] Starting hyperlink addition`);
    
    // Check API support
    const supported = await isHyperlinkSupported(ctx);
    if (!supported) {
        throw new Error("Hyperlinks require ExcelApi 1.7+; your version does not support this feature.");
    }
    
    let options = { url: null, email: null, documentReference: null, displayText: null, tooltip: "" };
    if (data) {
        try {
            options = { ...options, ...JSON.parse(data) };
        } catch (e) {
            console.warn(`[addHyperlink] Warning: Failed to parse data: ${e.message}`);
        }
    }
    
    // Validate: must have exactly one of url, email, or documentReference
    const linkTypes = [options.url, options.email, options.documentReference].filter(v => v);
    if (linkTypes.length === 0) {
        throw new Error("Invalid hyperlink data: must provide url, email, or documentReference");
    }
    if (linkTypes.length > 1) {
        throw new Error("Invalid hyperlink data: provide only one of url, email, or documentReference");
    }
    
    try {
        let hyperlinkObj = { screenTip: options.tooltip || "" };
        
        if (options.url) {
            // Validate URL format
            if (!options.url.match(/^https?:\/\//i) && !options.url.startsWith("//")) {
                options.url = "https://" + options.url;
            }
            hyperlinkObj.address = options.url;
            hyperlinkObj.textToDisplay = options.displayText || options.url;
            console.log(`[addHyperlink] Adding web URL: ${options.url}`);
        } else if (options.email) {
            // Automatically add mailto: prefix
            const emailAddress = options.email.startsWith("mailto:") ? options.email : "mailto:" + options.email;
            hyperlinkObj.address = emailAddress;
            hyperlinkObj.textToDisplay = options.displayText || options.email;
            console.log(`[addHyperlink] Adding email link: ${options.email}`);
        } else if (options.documentReference) {
            hyperlinkObj.documentReference = options.documentReference;
            hyperlinkObj.textToDisplay = options.displayText || options.documentReference;
            console.log(`[addHyperlink] Adding internal link: ${options.documentReference}`);
        }
        
        range.hyperlink = hyperlinkObj;
        await ctx.sync();
        
        logInfo(`Successfully added hyperlink`);
    } catch (e) {
        throw new Error(`Failed to add hyperlink: ${e.message}`);
    }
}

/**
 * Removes hyperlink(s) from a cell or range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * 
 * Note: This clears only the hyperlink, preserving cell values and formatting.
 * Always clears the entire range even if only some cells have hyperlinks.
 * Requires ExcelApi 1.7+
 */
async function removeHyperlink(ctx, range) {
    console.log(`[removeHyperlink] Starting hyperlink removal`);
    
    // Check API support
    const supported = await isHyperlinkSupported(ctx);
    if (!supported) {
        throw new Error("Hyperlinks require ExcelApi 1.7+; your version does not support this feature.");
    }
    
    try {
        // Clear hyperlinks from entire range using clear method
        // This works even if only some cells in the range have hyperlinks
        range.clear(Excel.ClearApplyTo.hyperlinks);
        await ctx.sync();
        
        logInfo(`Successfully removed hyperlinks from range`);
    } catch (e) {
        throw new Error(`Failed to remove hyperlink: ${e.message}`);
    }
}

/**
 * Edits an existing hyperlink or adds a new one if none exists
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Range} range - Target range
 * @param {string} data - JSON string with hyperlink options to update
 * 
 * Supported options (all optional - only provided fields are updated):
 * - url: New web URL
 * - email: New email address
 * - documentReference: New internal link
 * - displayText: New display text
 * - tooltip: New tooltip text
 * 
 * Note: If changing link type (e.g., url to documentReference), the old type is cleared.
 * Requires ExcelApi 1.7+
 */
async function editHyperlink(ctx, range, data) {
    console.log(`[editHyperlink] Starting hyperlink edit`);
    
    // Check API support
    const supported = await isHyperlinkSupported(ctx);
    if (!supported) {
        throw new Error("Hyperlinks require ExcelApi 1.7+; your version does not support this feature.");
    }
    
    let options = {};
    if (data) {
        try {
            options = JSON.parse(data);
        } catch (e) {
            console.warn(`[editHyperlink] Warning: Failed to parse data: ${e.message}`);
        }
    }
    
    try {
        // Load existing hyperlink
        range.load("hyperlink");
        await ctx.sync();
        
        const existingHyperlink = range.hyperlink || {};
        let hyperlinkObj = {
            screenTip: options.tooltip !== undefined ? options.tooltip : (existingHyperlink.screenTip || ""),
            textToDisplay: options.displayText !== undefined ? options.displayText : existingHyperlink.textToDisplay
        };
        
        // Determine link type - new value takes precedence
        if (options.url) {
            if (!options.url.match(/^https?:\/\//i) && !options.url.startsWith("//")) {
                options.url = "https://" + options.url;
            }
            hyperlinkObj.address = options.url;
            if (!options.displayText && !existingHyperlink.textToDisplay) {
                hyperlinkObj.textToDisplay = options.url;
            }
            console.log(`[editHyperlink] Updating to web URL: ${options.url}`);
        } else if (options.email) {
            const emailAddress = options.email.startsWith("mailto:") ? options.email : "mailto:" + options.email;
            hyperlinkObj.address = emailAddress;
            if (!options.displayText && !existingHyperlink.textToDisplay) {
                hyperlinkObj.textToDisplay = options.email;
            }
            console.log(`[editHyperlink] Updating to email link: ${options.email}`);
        } else if (options.documentReference) {
            hyperlinkObj.documentReference = options.documentReference;
            if (!options.displayText && !existingHyperlink.textToDisplay) {
                hyperlinkObj.textToDisplay = options.documentReference;
            }
            console.log(`[editHyperlink] Updating to internal link: ${options.documentReference}`);
        } else {
            // Keep existing link type
            if (existingHyperlink.address) {
                hyperlinkObj.address = existingHyperlink.address;
            } else if (existingHyperlink.documentReference) {
                hyperlinkObj.documentReference = existingHyperlink.documentReference;
            } else {
                throw new Error("No existing hyperlink to edit and no new link provided");
            }
        }
        
        range.hyperlink = hyperlinkObj;
        await ctx.sync();
        
        logInfo(`Successfully edited hyperlink`);
    } catch (e) {
        throw new Error(`Failed to edit hyperlink: ${e.message}`);
    }
}

// ============================================================================
// PivotTable Operations
// ============================================================================

/**
 * Creates a PivotTable from a data range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with PivotTable options
 */
async function createPivotTable(ctx, sheet, action) {
    logDebug(`Starting createPivotTable for target "${action.target}"`);
    
    let options = { name: null, destination: null, layout: null };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for createPivotTable`);
        }
    }
    
    // Validate required fields
    if (!options.name) {
        const errorMsg = `PivotTable name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.destination) {
        const errorMsg = `Destination is required (e.g., "PivotSheet!A1").`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Parse destination into sheet name and cell
        let destSheetName, destCell;
        if (options.destination.includes("!")) {
            const parts = options.destination.split("!");
            destSheetName = parts[0];
            destCell = parts[1];
        } else {
            destSheetName = sheet.name;
            destCell = options.destination;
        }
        
        // Get or create destination sheet
        let destSheet = ctx.workbook.worksheets.getItemOrNullObject(destSheetName);
        await ctx.sync();
        
        if (destSheet.isNullObject) {
            logDebug(`Creating new sheet "${destSheetName}" for PivotTable`);
            destSheet = ctx.workbook.worksheets.add(destSheetName);
            await ctx.sync();
        }
        
        // Get source range - could be a range address or table name
        let sourceRange;
        const source = action.target;
        
        // Check if source is a table name (no colon in address)
        if (source && !source.includes(":") && !source.match(/^[A-Z]+\d+$/i)) {
            // Try to get as table
            const table = sheet.tables.getItemOrNullObject(source);
            table.load("isNullObject");
            await ctx.sync();
            
            if (!table.isNullObject) {
                sourceRange = table.getRange();
                logDebug(`Using table "${source}" as PivotTable source`);
            } else {
                sourceRange = sheet.getRange(source);
            }
        } else {
            sourceRange = sheet.getRange(source);
        }
        
        // Get destination range
        const destRange = destSheet.getRange(destCell);
        
        // Create PivotTable
        const pivotTable = destSheet.pivotTables.add(options.name, sourceRange, destRange);
        
        // Set layout if specified
        if (options.layout) {
            const layoutType = options.layout.toLowerCase();
            if (layoutType === "compact") {
                pivotTable.layout.layoutType = Excel.PivotLayoutType.compact;
            } else if (layoutType === "outline") {
                pivotTable.layout.layoutType = Excel.PivotLayoutType.outline;
            } else if (layoutType === "tabular") {
                pivotTable.layout.layoutType = Excel.PivotLayoutType.tabular;
            }
        }
        
        await ctx.sync();
        logInfo(`Successfully created PivotTable "${options.name}" from ${source} to ${options.destination}`);
    } catch (e) {
        const errorMsg = `Failed to create PivotTable: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Adds a field to a PivotTable
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with field options
 */
async function addPivotField(ctx, sheet, action) {
    logDebug(`Starting addPivotField for target "${action.target}"`);
    
    let options = { pivotName: action.target, field: null, area: null, function: "Sum" };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for addPivotField`);
        }
    }
    
    const pivotName = options.pivotName || action.target;
    
    // Validate required fields
    if (!options.field) {
        const errorMsg = `Field name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.area) {
        const errorMsg = `Area is required (row, column, data, or filter).`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for PivotTable in all sheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let pivotTable = null;
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${pivotName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Get the hierarchy for the field
        const hierarchy = pivotTable.hierarchies.getItem(options.field);
        
        // Add to appropriate area
        const area = options.area.toLowerCase();
        if (area === "row") {
            pivotTable.rowHierarchies.add(hierarchy);
            logDebug(`Added field "${options.field}" to row area of PivotTable "${pivotName}"`);
        } else if (area === "column") {
            pivotTable.columnHierarchies.add(hierarchy);
            logDebug(`Added field "${options.field}" to column area of PivotTable "${pivotName}"`);
        } else if (area === "data" || area === "value" || area === "values") {
            const dataHierarchy = pivotTable.dataHierarchies.add(hierarchy);
            
            // Set aggregation function with validation
            const rawFuncName = options.function || "Sum";
            const funcName = rawFuncName.toLowerCase().replace(/_/g, ""); // Normalize aliases like "count_numbers"
            
            const funcMap = {
                "sum": Excel.AggregationFunction.sum,
                "count": Excel.AggregationFunction.count,
                "average": Excel.AggregationFunction.average,
                "avg": Excel.AggregationFunction.average,  // Common alias
                "max": Excel.AggregationFunction.max,
                "min": Excel.AggregationFunction.min,
                "countnumbers": Excel.AggregationFunction.countNumbers,
                "stddev": Excel.AggregationFunction.standardDeviation,
                "stdev": Excel.AggregationFunction.standardDeviation,  // Common alias
                "standarddeviation": Excel.AggregationFunction.standardDeviation,
                "var": Excel.AggregationFunction.variance,
                "variance": Excel.AggregationFunction.variance
            };
            
            const supportedFunctions = "Sum, Count, Average, Max, Min, CountNumbers, StdDev, Var";
            
            if (funcMap[funcName]) {
                dataHierarchy.summarizeBy = funcMap[funcName];
                logDebug(`Added field "${options.field}" to data area of PivotTable "${pivotName}" with ${funcName} aggregation`);
            } else {
                // Invalid function - warn and fall back to Sum
                logWarn(`Unknown aggregation function "${rawFuncName}". Supported: ${supportedFunctions}. Falling back to Sum.`);
                dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
                logDebug(`Added field "${options.field}" to data area of PivotTable "${pivotName}" with Sum aggregation (fallback)`);
            }
        } else if (area === "filter") {
            pivotTable.filterHierarchies.add(hierarchy);
            logDebug(`Added field "${options.field}" to filter area of PivotTable "${pivotName}"`);
        } else {
            const errorMsg = `Invalid area "${options.area}". Use row, column, data, or filter.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        await ctx.sync();
        logInfo(`Successfully added field "${options.field}" to ${area} area of PivotTable "${pivotName}"`);
    } catch (e) {
        const errorMsg = `Failed to add field "${options.field}" to PivotTable: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Configures PivotTable layout
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with layout options
 */
async function configurePivotLayout(ctx, sheet, action) {
    logDebug(`Starting configurePivotLayout for target "${action.target}"`);
    
    let options = { pivotName: action.target, layout: null, showRowHeaders: null, showColumnHeaders: null };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for configurePivotLayout`);
        }
    }
    
    const pivotName = options.pivotName || action.target;
    
    // Validate required fields
    if (!options.layout) {
        const errorMsg = `Layout type is required (Compact, Outline, or Tabular).`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for PivotTable in all sheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let pivotTable = null;
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${pivotName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Set layout type
        const layoutType = options.layout.toLowerCase();
        if (layoutType === "compact") {
            pivotTable.layout.layoutType = Excel.PivotLayoutType.compact;
        } else if (layoutType === "outline") {
            pivotTable.layout.layoutType = Excel.PivotLayoutType.outline;
        } else if (layoutType === "tabular") {
            pivotTable.layout.layoutType = Excel.PivotLayoutType.tabular;
        } else {
            const errorMsg = `Invalid layout type "${options.layout}". Use Compact, Outline, or Tabular.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Set optional header visibility
        if (options.showRowHeaders !== null && options.showRowHeaders !== undefined) {
            pivotTable.layout.showRowHeaders = options.showRowHeaders;
        }
        if (options.showColumnHeaders !== null && options.showColumnHeaders !== undefined) {
            pivotTable.layout.showColumnHeaders = options.showColumnHeaders;
        }
        
        await ctx.sync();
        logInfo(`Successfully configured layout for PivotTable "${pivotName}" to ${options.layout}`);
    } catch (e) {
        const errorMsg = `Failed to configure PivotTable layout: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Refreshes a PivotTable or all PivotTables
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with refresh options
 */
async function refreshPivotTable(ctx, sheet, action) {
    logDebug(`Starting refreshPivotTable for target "${action.target}"`);
    
    let options = { pivotName: action.target, refreshAll: false };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for refreshPivotTable`);
        }
    }
    
    try {
        if (options.refreshAll) {
            // Refresh all PivotTables in workbook
            ctx.workbook.pivotTables.refreshAll();
            await ctx.sync();
            logInfo(`Successfully refreshed all PivotTables in workbook`);
        } else {
            const pivotName = options.pivotName || action.target;
            
            // Search for PivotTable in all sheets
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            let pivotTable = null;
            for (const ws of sheets.items) {
                const pt = ws.pivotTables.getItemOrNullObject(pivotName);
                pt.load("isNullObject");
                await ctx.sync();
                
                if (!pt.isNullObject) {
                    pivotTable = pt;
                    break;
                }
            }
            
            if (!pivotTable) {
                const errorMsg = `PivotTable "${pivotName}" not found.`;
                logError(errorMsg);
                throw new Error(errorMsg);
            }
            
            pivotTable.refresh();
            await ctx.sync();
            logInfo(`Successfully refreshed PivotTable "${pivotName}"`);
        }
    } catch (e) {
        const errorMsg = `Failed to refresh PivotTable: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes a PivotTable
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with PivotTable name
 */
async function deletePivotTable(ctx, sheet, action) {
    logDebug(`Starting deletePivotTable for target "${action.target}"`);
    
    let options = { pivotName: action.target };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for deletePivotTable`);
        }
    }
    
    const pivotName = options.pivotName || action.target;
    
    try {
        // Search for PivotTable in all sheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let pivotTable = null;
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${pivotName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        pivotTable.delete();
        await ctx.sync();
        logInfo(`Successfully deleted PivotTable "${pivotName}"`);
    } catch (e) {
        const errorMsg = `Failed to delete PivotTable: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// Slicer Operations
// ============================================================================

/**
 * Valid slicer styles for validation
 */
const VALID_SLICER_STYLES = [
    ...Array.from({ length: 6 }, (_, i) => `SlicerStyleLight${i + 1}`),
    ...Array.from({ length: 6 }, (_, i) => `SlicerStyleDark${i + 1}`)
];

/**
 * Creates a slicer for a Table or PivotTable
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with slicer options
 */
async function createSlicer(ctx, sheet, action) {
    logDebug(`Starting createSlicer for target "${action.target}"`);
    
    let options = {
        slicerName: null,
        sourceType: null,
        sourceName: action.target,
        field: null,
        position: { left: 100, top: 100, width: 200, height: 200 },
        style: "SlicerStyleLight1",
        selectedItems: null,  // Array of items to select
        multiSelect: true     // Whether multiple items can be selected
    };
    
    if (action.data) {
        try {
            const parsed = JSON.parse(action.data);
            options = { ...options, ...parsed };
            if (parsed.position) {
                options.position = { ...options.position, ...parsed.position };
            }
        } catch (e) {
            logWarn(`Failed to parse action.data for createSlicer`);
        }
    }
    
    // Validate required fields
    if (!options.sourceName) {
        const errorMsg = `Source name (table or pivot name) is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.field) {
        const errorMsg = `Field name is required for slicer.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.sourceType || !["table", "pivot"].includes(options.sourceType.toLowerCase())) {
        const errorMsg = `Source type must be "table" or "pivot".`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        let slicerSource = null;
        let targetWorksheet = sheet;
        const sourceType = options.sourceType.toLowerCase();
        
        if (sourceType === "table") {
            // Search for table in all worksheets (Comment 4: align with pivot search behavior)
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            let table = null;
            for (const ws of sheets.items) {
                const tbl = ws.tables.getItemOrNullObject(options.sourceName);
                tbl.load("isNullObject");
                await ctx.sync();
                
                if (!tbl.isNullObject) {
                    table = tbl;
                    targetWorksheet = ws;
                    break;
                }
            }
            
            if (!table) {
                const errorMsg = `Table "${options.sourceName}" not found in any worksheet.`;
                logError(errorMsg);
                throw new Error(errorMsg);
            }
            
            // Comment 2: Validate that the field exists in the table
            table.columns.load("items");
            await ctx.sync();
            
            const columnNames = table.columns.items.map(col => {
                col.load("name");
                return col;
            });
            await ctx.sync();
            
            const fieldExists = columnNames.some(col => col.name === options.field);
            if (!fieldExists) {
                const availableColumns = columnNames.map(col => col.name).join(", ");
                const errorMsg = `Field "${options.field}" not found in table "${options.sourceName}". Available columns: ${availableColumns}`;
                logError(errorMsg);
                throw new Error(errorMsg);
            }
            
            slicerSource = table;
            logDebug(`Found table "${options.sourceName}" for slicer with valid field "${options.field}"`);
        } else if (sourceType === "pivot") {
            // Search for PivotTable in all sheets
            const sheets = ctx.workbook.worksheets;
            sheets.load("items");
            await ctx.sync();
            
            let pivotTable = null;
            for (const ws of sheets.items) {
                const pt = ws.pivotTables.getItemOrNullObject(options.sourceName);
                pt.load("isNullObject");
                await ctx.sync();
                
                if (!pt.isNullObject) {
                    pivotTable = pt;
                    ws.load("name");
                    await ctx.sync();
                    targetWorksheet = ws;
                    break;
                }
            }
            
            if (!pivotTable) {
                const errorMsg = `PivotTable "${options.sourceName}" not found.`;
                logError(errorMsg);
                throw new Error(errorMsg);
            }
            
            // Comment 2: Validate that the field exists in the pivot table hierarchies
            pivotTable.hierarchies.load("items");
            await ctx.sync();
            
            const hierarchyNames = pivotTable.hierarchies.items.map(h => {
                h.load("name");
                return h;
            });
            await ctx.sync();
            
            const fieldExists = hierarchyNames.some(h => h.name === options.field);
            if (!fieldExists) {
                const availableFields = hierarchyNames.map(h => h.name).join(", ");
                const errorMsg = `Field "${options.field}" not found in PivotTable "${options.sourceName}". Available fields: ${availableFields}`;
                logError(errorMsg);
                throw new Error(errorMsg);
            }
            
            slicerSource = pivotTable;
            logDebug(`Found PivotTable "${options.sourceName}" for slicer with valid field "${options.field}"`);
        }
        
        // Create the slicer
        const slicer = targetWorksheet.slicers.add(slicerSource, options.field, targetWorksheet);
        
        // Set slicer name if provided
        if (options.slicerName) {
            slicer.name = options.slicerName;
        }
        
        // Set position and size
        slicer.left = options.position.left || 100;
        slicer.top = options.position.top || 100;
        slicer.width = options.position.width || 200;
        slicer.height = options.position.height || 200;
        
        // Set style with validation
        if (options.style) {
            if (VALID_SLICER_STYLES.includes(options.style)) {
                slicer.style = options.style;
            } else {
                logWarn(`Invalid slicer style "${options.style}". Using default SlicerStyleLight1.`);
                slicer.style = "SlicerStyleLight1";
            }
        }
        
        await ctx.sync();
        
        // Comment 1: Configure slicer item selections if specified
        if (options.selectedItems && Array.isArray(options.selectedItems) && options.selectedItems.length > 0) {
            slicer.slicerItems.load("items");
            await ctx.sync();
            
            const slicerItems = slicer.slicerItems.items;
            for (const item of slicerItems) {
                item.load("name");
            }
            await ctx.sync();
            
            // If multiSelect is false, only select the first item from selectedItems
            const itemsToSelect = options.multiSelect === false 
                ? [options.selectedItems[0]] 
                : options.selectedItems;
            
            for (const item of slicerItems) {
                const shouldBeSelected = itemsToSelect.includes(item.name);
                item.isSelected = shouldBeSelected;
            }
            
            await ctx.sync();
            logDebug(`Configured slicer selections: ${itemsToSelect.join(", ")}`);
        }
        
        const slicerDisplayName = options.slicerName || options.field;
        logInfo(`Successfully created slicer "${slicerDisplayName}" for ${sourceType} "${options.sourceName}" on field "${options.field}"`);
    } catch (e) {
        const errorMsg = `Failed to create slicer: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Configures an existing slicer's properties
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with configuration options
 */
async function configureSlicer(ctx, sheet, action) {
    logDebug(`Starting configureSlicer for target "${action.target}"`);
    
    let options = {
        slicerName: action.target,
        caption: null,
        style: null,
        sortBy: null,
        width: null,
        height: null,
        left: null,
        top: null,
        selectedItems: null,  // Comment 1: Array of items to select
        multiSelect: true     // Comment 1: Whether multiple items can be selected
    };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for configureSlicer`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for slicer in all worksheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let slicer = null;
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load("isNullObject");
            await ctx.sync();
            
            if (!sl.isNullObject) {
                slicer = sl;
                break;
            }
        }
        
        if (!slicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        const updatedProps = [];
        
        // Apply properties conditionally
        if (options.caption !== null && options.caption !== undefined) {
            slicer.caption = options.caption;
            updatedProps.push("caption");
        }
        
        if (options.style) {
            if (VALID_SLICER_STYLES.includes(options.style)) {
                slicer.style = options.style;
                updatedProps.push("style");
            } else {
                logWarn(`Invalid slicer style "${options.style}". Skipping style update.`);
            }
        }
        
        if (options.sortBy) {
            const sortMap = {
                "datasourceorder": Excel.SlicerSortType.dataSourceOrder,
                "ascending": Excel.SlicerSortType.ascending,
                "descending": Excel.SlicerSortType.descending
            };
            const sortKey = options.sortBy.toLowerCase().replace(/\s/g, "");
            if (sortMap[sortKey]) {
                slicer.sortBy = sortMap[sortKey];
                updatedProps.push("sortBy");
            } else {
                logWarn(`Invalid sortBy value "${options.sortBy}". Use DataSourceOrder, Ascending, or Descending.`);
            }
        }
        
        if (options.width !== null && options.width !== undefined) {
            slicer.width = options.width;
            updatedProps.push("width");
        }
        
        if (options.height !== null && options.height !== undefined) {
            slicer.height = options.height;
            updatedProps.push("height");
        }
        
        if (options.left !== null && options.left !== undefined) {
            slicer.left = options.left;
            updatedProps.push("left");
        }
        
        if (options.top !== null && options.top !== undefined) {
            slicer.top = options.top;
            updatedProps.push("top");
        }
        
        await ctx.sync();
        
        // Comment 1: Configure slicer item selections if specified
        if (options.selectedItems && Array.isArray(options.selectedItems) && options.selectedItems.length > 0) {
            slicer.slicerItems.load("items");
            await ctx.sync();
            
            const slicerItems = slicer.slicerItems.items;
            for (const item of slicerItems) {
                item.load("name");
            }
            await ctx.sync();
            
            // If multiSelect is false, only select the first item from selectedItems
            const itemsToSelect = options.multiSelect === false 
                ? [options.selectedItems[0]] 
                : options.selectedItems;
            
            for (const item of slicerItems) {
                const shouldBeSelected = itemsToSelect.includes(item.name);
                item.isSelected = shouldBeSelected;
            }
            
            await ctx.sync();
            updatedProps.push(`selectedItems(${itemsToSelect.length})`);
            logDebug(`Configured slicer selections: ${itemsToSelect.join(", ")}`);
        }
        
        logInfo(`Successfully configured slicer "${slicerName}". Updated: ${updatedProps.join(", ") || "none"}`);
    } catch (e) {
        const errorMsg = `Failed to configure slicer: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Connects a slicer to a different table (recreates slicer)
 * Note: Office.js doesn't support rebinding slicers; this deletes and recreates
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with connection options
 */
async function connectSlicerToTable(ctx, sheet, action) {
    logDebug(`Starting connectSlicerToTable for target "${action.target}"`);
    
    let options = {
        slicerName: action.target,
        tableName: null,
        field: null
    };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for connectSlicerToTable`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.tableName) {
        const errorMsg = `Table name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.field) {
        const errorMsg = `Field name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Find existing slicer to get its properties
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let existingSlicer = null;
        let slicerWorksheet = null;
        
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load(["isNullObject", "left", "top", "width", "height", "style", "caption"]);
            await ctx.sync();
            
            if (!sl.isNullObject) {
                existingSlicer = sl;
                slicerWorksheet = ws;
                break;
            }
        }
        
        if (!existingSlicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Store slicer properties before deletion
        const slicerProps = {
            left: existingSlicer.left,
            top: existingSlicer.top,
            width: existingSlicer.width,
            height: existingSlicer.height,
            style: existingSlicer.style,
            caption: existingSlicer.caption
        };
        
        // Delete existing slicer
        existingSlicer.delete();
        await ctx.sync();
        logDebug(`Deleted existing slicer "${slicerName}" for reconnection`);
        
        // Comment 3: Search for table in all worksheets instead of just active sheet
        let table = null;
        let tableWorksheet = slicerWorksheet; // Default to original slicer's worksheet
        
        for (const ws of sheets.items) {
            const tbl = ws.tables.getItemOrNullObject(options.tableName);
            tbl.load("isNullObject");
            await ctx.sync();
            
            if (!tbl.isNullObject) {
                table = tbl;
                tableWorksheet = ws;
                break;
            }
        }
        
        if (!table) {
            const errorMsg = `Table "${options.tableName}" not found in any worksheet.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Create new slicer on the table's worksheet (Comment 3: use correct worksheet)
        const newSlicer = tableWorksheet.slicers.add(table, options.field, tableWorksheet);
        newSlicer.name = slicerName;
        newSlicer.left = slicerProps.left;
        newSlicer.top = slicerProps.top;
        newSlicer.width = slicerProps.width;
        newSlicer.height = slicerProps.height;
        if (slicerProps.style) newSlicer.style = slicerProps.style;
        if (slicerProps.caption) newSlicer.caption = slicerProps.caption;
        
        await ctx.sync();
        logInfo(`Successfully reconnected slicer "${slicerName}" to table "${options.tableName}" on field "${options.field}"`);
    } catch (e) {
        const errorMsg = `Failed to connect slicer to table: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Connects a slicer to a different PivotTable (recreates slicer)
 * Note: Office.js doesn't support rebinding slicers; this deletes and recreates
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with connection options
 */
async function connectSlicerToPivot(ctx, sheet, action) {
    logDebug(`Starting connectSlicerToPivot for target "${action.target}"`);
    
    let options = {
        slicerName: action.target,
        pivotName: null,
        field: null
    };
    
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for connectSlicerToPivot`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.pivotName) {
        const errorMsg = `PivotTable name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (!options.field) {
        const errorMsg = `Field name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Find existing slicer to get its properties
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let existingSlicer = null;
        
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load(["isNullObject", "left", "top", "width", "height", "style", "caption"]);
            await ctx.sync();
            
            if (!sl.isNullObject) {
                existingSlicer = sl;
                break;
            }
        }
        
        if (!existingSlicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Store slicer properties before deletion
        const slicerProps = {
            left: existingSlicer.left,
            top: existingSlicer.top,
            width: existingSlicer.width,
            height: existingSlicer.height,
            style: existingSlicer.style,
            caption: existingSlicer.caption
        };
        
        // Delete existing slicer
        existingSlicer.delete();
        await ctx.sync();
        logDebug(`Deleted existing slicer "${slicerName}" for reconnection`);
        
        // Find the PivotTable
        let pivotTable = null;
        let pivotWorksheet = null;
        
        for (const ws of sheets.items) {
            const pt = ws.pivotTables.getItemOrNullObject(options.pivotName);
            pt.load("isNullObject");
            await ctx.sync();
            
            if (!pt.isNullObject) {
                pivotTable = pt;
                pivotWorksheet = ws;
                break;
            }
        }
        
        if (!pivotTable) {
            const errorMsg = `PivotTable "${options.pivotName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Create new slicer with same properties
        const newSlicer = pivotWorksheet.slicers.add(pivotTable, options.field, pivotWorksheet);
        newSlicer.name = slicerName;
        newSlicer.left = slicerProps.left;
        newSlicer.top = slicerProps.top;
        newSlicer.width = slicerProps.width;
        newSlicer.height = slicerProps.height;
        if (slicerProps.style) newSlicer.style = slicerProps.style;
        if (slicerProps.caption) newSlicer.caption = slicerProps.caption;
        
        await ctx.sync();
        logInfo(`Successfully reconnected slicer "${slicerName}" to PivotTable "${options.pivotName}" on field "${options.field}"`);
    } catch (e) {
        const errorMsg = `Failed to connect slicer to PivotTable: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes a slicer
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with slicer name
 */
async function deleteSlicer(ctx, sheet, action) {
    logDebug(`Starting deleteSlicer for target "${action.target}"`);
    
    let options = { slicerName: action.target };
    if (action.data) {
        try {
            options = { ...options, ...JSON.parse(action.data) };
        } catch (e) {
            logWarn(`Failed to parse action.data for deleteSlicer`);
        }
    }
    
    const slicerName = options.slicerName || action.target;
    
    if (!slicerName) {
        const errorMsg = `Slicer name is required.`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        // Search for slicer in all worksheets
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let slicer = null;
        for (const ws of sheets.items) {
            ws.slicers.load("items");
            await ctx.sync();
            
            const sl = ws.slicers.getItemOrNullObject(slicerName);
            sl.load("isNullObject");
            await ctx.sync();
            
            if (!sl.isNullObject) {
                slicer = sl;
                break;
            }
        }
        
        if (!slicer) {
            const errorMsg = `Slicer "${slicerName}" not found.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        slicer.delete();
        await ctx.sync();
        logInfo(`Successfully deleted slicer "${slicerName}"`);
    } catch (e) {
        const errorMsg = `Failed to delete slicer: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// Named Range Operations
// ============================================================================

/**
 * Validates a named range name
 * @param {string} name - Name to validate
 * @returns {Object} Validation result with isValid and error message
 */
function validateNamedRangeName(name) {
    if (!name || typeof name !== "string") {
        return { isValid: false, error: "Named range name is required." };
    }
    
    // Must start with a letter or underscore
    if (!/^[A-Za-z_]/.test(name)) {
        return { isValid: false, error: "Named range name must start with a letter or underscore." };
    }
    
    // Can only contain letters, numbers, underscores, and periods
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name)) {
        return { isValid: false, error: "Named range name can only contain letters, numbers, underscores, and periods. Spaces are not allowed." };
    }
    
    // Cannot be a cell reference (e.g., A1, XFD1048576)
    if (/^[A-Za-z]{1,3}\d+$/.test(name)) {
        return { isValid: false, error: "Named range name cannot look like a cell reference (e.g., A1, B2)." };
    }
    
    // Max 255 characters
    if (name.length > 255) {
        return { isValid: false, error: "Named range name cannot exceed 255 characters." };
    }
    
    return { isValid: true };
}

/**
 * Creates a named range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with name, scope, formula, comment
 * 
 * For workbook-scoped named ranges referencing other sheets, use one of:
 * 1. Sheet-qualified target: "Sheet2!A1:B5" - will resolve to the correct sheet
 * 2. Formula option: {"formula":"=Sheet2!A1:B5"} - for explicit formula-based references
 * 
 * For worksheet-scoped names, target is always relative to the active sheet.
 */
async function createNamedRange(ctx, sheet, action) {
    logDebug(`Starting createNamedRange for target "${action.target}"`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logWarn(`Failed to parse action.data for createNamedRange`);
        }
    }
    
    const name = options.name;
    const scope = options.scope || "workbook";
    const formula = options.formula;
    const comment = options.comment || "";
    
    // Validate name
    const validation = validateNamedRangeName(name);
    if (!validation.isValid) {
        logError(validation.error);
        throw new Error(validation.error);
    }
    
    try {
        // Check for existing name
        let existingName;
        if (scope === "worksheet") {
            existingName = sheet.names.getItemOrNullObject(name);
        } else {
            existingName = ctx.workbook.names.getItemOrNullObject(name);
        }
        existingName.load("isNullObject");
        await ctx.sync();
        
        if (!existingName.isNullObject) {
            const errorMsg = `A named range called '${name}' already exists in ${scope} scope. Choose a different name or delete the existing one first.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        // Determine what to add - formula or range reference
        let namedItem;
        if (formula) {
            // Named formula or constant
            const formulaValue = formula.startsWith("=") ? formula : `=${formula}`;
            if (scope === "worksheet") {
                namedItem = sheet.names.add(name, formulaValue, comment);
            } else {
                namedItem = ctx.workbook.names.add(name, formulaValue, comment);
            }
            logDebug(`Creating named formula '${name}' with formula '${formulaValue}'`);
        } else {
            // Named range reference
            if (!action.target) {
                const errorMsg = "Target range is required for named range (e.g., 'A1:E100' or 'Sheet2!A1:B5').";
                logError(errorMsg);
                throw new Error(errorMsg);
            }
            
            // Check if target contains sheet reference (e.g., "Sheet2!A1:B5")
            let targetRange;
            if (action.target.includes("!")) {
                // Sheet-qualified reference - parse and resolve
                const parts = action.target.split("!");
                const sheetName = parts[0].replace(/^'|'$/g, ""); // Remove quotes if present
                const rangeAddress = parts.slice(1).join("!"); // Handle edge case of ! in range
                
                const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
                targetSheet.load("isNullObject");
                await ctx.sync();
                
                if (targetSheet.isNullObject) {
                    const errorMsg = `Sheet "${sheetName}" not found. Check the sheet name in target "${action.target}".`;
                    logError(errorMsg);
                    throw new Error(errorMsg);
                }
                
                targetRange = targetSheet.getRange(rangeAddress);
                logDebug(`Resolved cross-sheet reference: ${sheetName}!${rangeAddress}`);
            } else {
                // Local range on active sheet
                targetRange = sheet.getRange(action.target);
            }
            
            if (scope === "worksheet") {
                namedItem = sheet.names.add(name, targetRange, comment);
            } else {
                namedItem = ctx.workbook.names.add(name, targetRange, comment);
            }
            logDebug(`Creating named range '${name}' for range '${action.target}'`);
        }
        
        await ctx.sync();
        logInfo(`Successfully created named range '${name}' with ${scope} scope`);
    } catch (e) {
        const errorMsg = `Failed to create named range: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Deletes a named range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with named range name as target
 */
async function deleteNamedRange(ctx, sheet, action) {
    logDebug(`Starting deleteNamedRange for target "${action.target}"`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logWarn(`Failed to parse action.data for deleteNamedRange`);
        }
    }
    
    const name = action.target;
    const scope = options.scope || "workbook";
    
    if (!name) {
        const errorMsg = "Named range name is required.";
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        let namedItem;
        if (scope === "worksheet") {
            namedItem = sheet.names.getItemOrNullObject(name);
        } else {
            namedItem = ctx.workbook.names.getItemOrNullObject(name);
        }
        namedItem.load("isNullObject");
        await ctx.sync();
        
        if (namedItem.isNullObject) {
            logWarn(`Named range '${name}' not found in ${scope} scope. Nothing to delete.`);
            return;
        }
        
        namedItem.delete();
        await ctx.sync();
        logInfo(`Successfully deleted named range '${name}' from ${scope} scope`);
    } catch (e) {
        const errorMsg = `Failed to delete named range: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Updates a named range
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with named range name as target, newFormula, newComment
 */
async function updateNamedRange(ctx, sheet, action) {
    logDebug(`Starting updateNamedRange for target "${action.target}"`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logWarn(`Failed to parse action.data for updateNamedRange`);
        }
    }
    
    const name = action.target;
    const scope = options.scope || "workbook";
    const newFormula = options.newFormula;
    const newComment = options.newComment;
    
    if (!name) {
        const errorMsg = "Named range name is required.";
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    if (newFormula === undefined && newComment === undefined) {
        const errorMsg = "At least one of newFormula or newComment must be provided.";
        logError(errorMsg);
        throw new Error(errorMsg);
    }
    
    try {
        let namedItem;
        if (scope === "worksheet") {
            namedItem = sheet.names.getItemOrNullObject(name);
        } else {
            namedItem = ctx.workbook.names.getItemOrNullObject(name);
        }
        namedItem.load(["isNullObject", "formula", "comment"]);
        await ctx.sync();
        
        if (namedItem.isNullObject) {
            const errorMsg = `Named range '${name}' not found in ${scope} scope. Use listNamedRanges to see available names.`;
            logError(errorMsg);
            throw new Error(errorMsg);
        }
        
        const updates = [];
        if (newFormula !== undefined) {
            const formulaValue = newFormula.startsWith("=") ? newFormula : `=${newFormula}`;
            namedItem.formula = formulaValue;
            updates.push(`formula=${formulaValue}`);
        }
        if (newComment !== undefined) {
            namedItem.comment = newComment;
            updates.push(`comment=${newComment}`);
        }
        
        await ctx.sync();
        logInfo(`Successfully updated named range '${name}': ${updates.join(", ")}`);
    } catch (e) {
        const errorMsg = `Failed to update named range: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

/**
 * Lists named ranges (diagnostics-only)
 * 
 * NOTE: This action is primarily for diagnostics and debugging purposes.
 * Results are logged to the diagnostics panel but are NOT returned to the AI
 * or surfaced in the UI, as the executeAction architecture does not currently
 * support action return values. The AI can reference named ranges through the
 * data context built by excel-data.js which includes existing named ranges.
 * 
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Active worksheet
 * @param {Object} action - Action with scope option
 * @returns {Promise<Array>} Array of named range objects (for internal use only)
 */
async function listNamedRanges(ctx, sheet, action) {
    logDebug(`Starting listNamedRanges (diagnostics-only)`);
    
    let options = {};
    if (action.data) {
        try {
            options = JSON.parse(action.data);
        } catch (e) {
            logWarn(`Failed to parse action.data for listNamedRanges`);
        }
    }
    
    const scope = options.scope || "all";
    
    try {
        const results = [];
        
        // Load workbook-scoped names
        if (scope === "all" || scope === "workbook") {
            ctx.workbook.names.load("items");
            await ctx.sync();
            
            for (const item of ctx.workbook.names.items) {
                item.load(["name", "formula", "comment", "type", "visible"]);
            }
            await ctx.sync();
            
            for (const item of ctx.workbook.names.items) {
                results.push({
                    name: item.name,
                    scope: "workbook",
                    formula: item.formula,
                    comment: item.comment || "",
                    type: item.type,
                    visible: item.visible
                });
            }
            logDebug(`Found ${ctx.workbook.names.items.length} workbook-scoped named ranges`);
        }
        
        // Load worksheet-scoped names
        if (scope === "all" || scope === "worksheet") {
            sheet.names.load("items");
            await ctx.sync();
            
            for (const item of sheet.names.items) {
                item.load(["name", "formula", "comment", "type", "visible"]);
            }
            await ctx.sync();
            
            for (const item of sheet.names.items) {
                results.push({
                    name: item.name,
                    scope: "worksheet",
                    sheetName: sheet.name,
                    formula: item.formula,
                    comment: item.comment || "",
                    type: item.type,
                    visible: item.visible
                });
            }
            logDebug(`Found ${sheet.names.items.length} worksheet-scoped named ranges`);
        }
        
        // Log results
        if (results.length === 0) {
            logInfo("No named ranges found.");
        } else {
            logInfo(`=== Named Ranges (${results.length} total) ===`);
            for (const nr of results) {
                const scopeInfo = nr.scope === "worksheet" ? `worksheet:${nr.sheetName}` : "workbook";
                logDebug(`  ${nr.name} [${scopeInfo}]: ${nr.formula}${nr.comment ? ` (${nr.comment})` : ""}`);
            }
        }
        
        return results;
    } catch (e) {
        const errorMsg = `Failed to list named ranges: ${e.message}`;
        logError(errorMsg);
        throw new Error(errorMsg);
    }
}

// ============================================================================
// Protection Operations
// ============================================================================

/**
 * Protects a worksheet with optional password and permissions
 */
async function protectWorksheet(ctx, sheet, action) {
    console.log(`[protectWorksheet] Protecting worksheet: ${action.target || sheet.name}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        const targetSheetName = action.target || sheet.name;
        
        // Get target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(targetSheetName);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${targetSheetName}" not found`);
        }
        
        // Check if already protected
        const protection = targetSheet.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (protection.protected) {
            throw new Error(`Sheet "${targetSheetName}" is already protected. Unprotect it first.`);
        }
        
        // Build protection options
        const options = {
            allowAutoFilter: data.allowAutoFilter !== false,
            allowDeleteColumns: data.allowDeleteColumns === true,
            allowDeleteRows: data.allowDeleteRows === true,
            allowFormatCells: data.allowFormatCells === true,
            allowFormatColumns: data.allowFormatColumns === true,
            allowFormatRows: data.allowFormatRows === true,
            allowInsertColumns: data.allowInsertColumns === true,
            allowInsertRows: data.allowInsertRows === true,
            allowInsertHyperlinks: data.allowInsertHyperlinks === true,
            allowPivotTables: data.allowPivotTables === true,
            allowSort: data.allowSort === true,
            selectionMode: data.selectionMode || "Normal"
        };
        
        // Apply protection
        const password = data.password || undefined;
        protection.protect(options, password);
        await ctx.sync();
        
        console.log(`[protectWorksheet] Successfully protected "${targetSheetName}"`);
    } catch (error) {
        console.log(`[protectWorksheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unprotects a worksheet
 */
async function unprotectWorksheet(ctx, sheet, action) {
    console.log(`[unprotectWorksheet] Unprotecting worksheet: ${action.target || sheet.name}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        const targetSheetName = action.target || sheet.name;
        
        // Get target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(targetSheetName);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${targetSheetName}" not found`);
        }
        
        // Check if protected
        const protection = targetSheet.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (!protection.protected) {
            console.log(`[unprotectWorksheet] Sheet "${targetSheetName}" is not protected, skipping`);
            return;
        }
        
        // Unprotect with password if provided
        const password = data.password || undefined;
        protection.unprotect(password);
        await ctx.sync();
        
        console.log(`[unprotectWorksheet] Successfully unprotected "${targetSheetName}"`);
    } catch (error) {
        console.log(`[unprotectWorksheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Protects a range by locking cells (requires worksheet protection to take effect)
 */
async function protectRange(ctx, sheet, action) {
    console.log(`[protectRange] Protecting range: ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        const range = sheet.getRange(action.target);
        
        // Set protection properties
        range.format.protection.locked = data.locked !== false;
        range.format.protection.formulaHidden = data.formulaHidden === true;
        await ctx.sync();
        
        console.log(`[protectRange] Successfully set protection for "${action.target}" (locked: ${data.locked !== false}, formulaHidden: ${data.formulaHidden === true})`);
        console.log(`[protectRange] Note: Protection takes effect only when worksheet is protected`);
    } catch (error) {
        console.log(`[protectRange] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unprotects a range by unlocking cells
 */
async function unprotectRange(ctx, sheet, action) {
    console.log(`[unprotectRange] Unprotecting range: ${action.target}`);
    
    try {
        const range = sheet.getRange(action.target);
        
        // Unlock cells and unhide formulas
        range.format.protection.locked = false;
        range.format.protection.formulaHidden = false;
        await ctx.sync();
        
        console.log(`[unprotectRange] Successfully unlocked "${action.target}"`);
    } catch (error) {
        console.log(`[unprotectRange] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Protects workbook structure (prevents sheet add/delete/rename/move)
 */
async function protectWorkbook(ctx, sheet, action) {
    console.log(`[protectWorkbook] Protecting workbook structure`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Check if already protected
        const protection = ctx.workbook.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (protection.protected) {
            throw new Error("Workbook is already protected. Unprotect it first.");
        }
        
        // Apply protection
        const password = data.password || undefined;
        protection.protect(password);
        await ctx.sync();
        
        console.log(`[protectWorkbook] Successfully protected workbook structure`);
    } catch (error) {
        console.log(`[protectWorkbook] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unprotects workbook structure
 */
async function unprotectWorkbook(ctx, sheet, action) {
    console.log(`[unprotectWorkbook] Unprotecting workbook structure`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Check if protected
        const protection = ctx.workbook.protection;
        protection.load("protected");
        await ctx.sync();
        
        if (!protection.protected) {
            console.log(`[unprotectWorkbook] Workbook is not protected, skipping`);
            return;
        }
        
        // Unprotect with password if provided
        const password = data.password || undefined;
        protection.unprotect(password);
        await ctx.sync();
        
        console.log(`[unprotectWorkbook] Successfully unprotected workbook structure`);
    } catch (error) {
        console.log(`[unprotectWorkbook] Error: ${error.message}`);
        throw error;
    }
}

// ============================================================================
// Shape and Image Operations
// ============================================================================

// Valid shape types for Office.js (all lowercase for validation)
// Original casing is preserved in SHAPE_TYPE_MAP for Office.js API calls
const VALID_SHAPE_TYPES = [
    "rectangle", "oval", "triangle", "righttriangle", "parallelogram", "trapezoid",
    "hexagon", "octagon", "pentagon", "plus", "star4", "star5", "star6",
    "arrow", "chevron", "homeplate", "cube", "bevel", "foldedcorner",
    "smileyface", "donut", "nosmoking", "blockarc", "heart", "lightningbolt",
    "sun", "moon", "cloud", "arc", "bracepair", "bracketpair", "can",
    "flowchartprocess", "flowchartdecision", "flowchartdata", "flowchartterminator",
    "line", "lineinverse", "straightconnector1", "bentconnector2", "bentconnector3"
];

// Maps lowercase shape types to proper Office.js enum casing
const SHAPE_TYPE_MAP = {
    "rectangle": "Rectangle", "oval": "Oval", "triangle": "Triangle",
    "righttriangle": "RightTriangle", "parallelogram": "Parallelogram", "trapezoid": "Trapezoid",
    "hexagon": "Hexagon", "octagon": "Octagon", "pentagon": "Pentagon",
    "plus": "Plus", "star4": "Star4", "star5": "Star5", "star6": "Star6",
    "arrow": "Arrow", "chevron": "Chevron", "homeplate": "HomePlate",
    "cube": "Cube", "bevel": "Bevel", "foldedcorner": "FoldedCorner",
    "smileyface": "SmileyFace", "donut": "Donut", "nosmoking": "NoSmoking",
    "blockarc": "BlockArc", "heart": "Heart", "lightningbolt": "LightningBolt",
    "sun": "Sun", "moon": "Moon", "cloud": "Cloud", "arc": "Arc",
    "bracepair": "BracePair", "bracketpair": "BracketPair", "can": "Can",
    "flowchartprocess": "FlowchartProcess", "flowchartdecision": "FlowchartDecision",
    "flowchartdata": "FlowchartData", "flowchartterminator": "FlowchartTerminator",
    "line": "Line", "lineinverse": "LineInverse",
    "straightconnector1": "StraightConnector1", "bentconnector2": "BentConnector2", "bentconnector3": "BentConnector3"
};

/**
 * Inserts a geometric shape at a specified cell position
 */
async function insertShape(ctx, sheet, action) {
    console.log(`[insertShape] Starting shape insertion at ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        const shapeType = data.shapeType || "rectangle";
        const normalizedType = shapeType.toLowerCase();
        
        // Validate shape type using normalized lowercase comparison
        if (!VALID_SHAPE_TYPES.includes(normalizedType)) {
            console.log(`[insertShape] Error: Invalid shape type "${shapeType}"`);
            throw new Error(`Invalid shape type "${shapeType}". Valid types: rectangle, oval, triangle, rightTriangle, arrow, star5, hexagon, line, etc.`);
        }
        
        // Get position from target cell
        let left = 100, top = 100;
        if (action.target) {
            try {
                const posRange = sheet.getRange(action.target);
                posRange.load(["left", "top"]);
                await ctx.sync();
                left = posRange.left;
                top = posRange.top;
            } catch (posError) {
                console.log(`[insertShape] Warning: Could not parse position "${action.target}", using default`);
            }
        }
        
        // Map normalized shape type to proper Office.js enum casing
        const excelShapeType = SHAPE_TYPE_MAP[normalizedType] || (shapeType.charAt(0).toUpperCase() + shapeType.slice(1));
        
        // Create shape
        const shape = sheet.shapes.addGeometricShape(excelShapeType);
        
        // Set position
        shape.left = left;
        shape.top = top;
        
        // Set dimensions
        const width = data.width || 150;
        const height = data.height || 100;
        if (width <= 0 || height <= 0) {
            throw new Error("Shape dimensions must be positive numbers");
        }
        shape.width = width;
        shape.height = height;
        
        // Set rotation
        if (data.rotation !== undefined) {
            shape.rotation = data.rotation;
        }
        
        // Apply fill color
        if (data.fill && data.fill !== "none") {
            shape.fill.setSolidColor(data.fill);
        } else if (data.fill === "none") {
            shape.fill.clear();
        }
        
        // Apply line/border formatting
        if (data.lineColor && data.lineColor !== "none") {
            shape.lineFormat.color = data.lineColor;
        }
        if (data.lineWeight) {
            shape.lineFormat.weight = data.lineWeight;
        }
        
        // Add text if provided
        if (data.text) {
            shape.textFrame.textRange.text = data.text;
        }
        
        // Set custom name if provided
        if (data.name) {
            shape.name = data.name;
        }
        
        await ctx.sync();
        
        console.log(`[insertShape] Successfully created ${shapeType} shape at position (${left}, ${top})`);
    } catch (error) {
        console.log(`[insertShape] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Inserts an image from Base64-encoded data
 */
async function insertImage(ctx, sheet, action) {
    console.log(`[insertImage] Starting image insertion at ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Validate source
        if (!data.source) {
            throw new Error("insertImage requires a Base64-encoded image string in data.source");
        }
        
        // Extract Base64 data (handle data URI format)
        let base64Data = data.source;
        let isSvg = false;
        
        if (base64Data.startsWith("data:image/svg")) {
            isSvg = true;
            // For SVG, we need the XML content, not Base64
            if (base64Data.includes(";base64,")) {
                base64Data = atob(base64Data.split(";base64,")[1]);
            }
        } else if (base64Data.startsWith("data:image/")) {
            // Extract just the Base64 part
            base64Data = base64Data.split(",")[1] || base64Data;
        }
        
        // Get position from target cell
        let left = 100, top = 100;
        if (action.target) {
            try {
                const posRange = sheet.getRange(action.target);
                posRange.load(["left", "top"]);
                await ctx.sync();
                left = posRange.left;
                top = posRange.top;
            } catch (posError) {
                console.log(`[insertImage] Warning: Could not parse position "${action.target}", using default`);
            }
        }
        
        // Insert image
        let image;
        if (isSvg) {
            image = sheet.shapes.addSvg(base64Data);
        } else {
            image = sheet.shapes.addImage(base64Data);
        }
        
        // Set position
        image.left = left;
        image.top = top;
        
        // Set dimensions
        if (data.width) image.width = data.width;
        if (data.height) image.height = data.height;
        
        // Lock aspect ratio by default
        image.lockAspectRatio = data.lockAspectRatio !== false;
        
        // Set name and alt text
        if (data.name) image.name = data.name;
        if (data.altText) image.altTextDescription = data.altText;
        
        await ctx.sync();
        
        console.log(`[insertImage] Successfully inserted image at position (${left}, ${top})`);
    } catch (error) {
        console.log(`[insertImage] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Inserts a text box at a specified cell position
 */
async function insertTextBox(ctx, sheet, action) {
    console.log(`[insertTextBox] Starting text box insertion at ${action.target}`);
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Validate text
        if (!data.text) {
            throw new Error("insertTextBox requires text content in data.text");
        }
        
        // Get position from target cell
        let left = 100, top = 100;
        if (action.target) {
            try {
                const posRange = sheet.getRange(action.target);
                posRange.load(["left", "top"]);
                await ctx.sync();
                left = posRange.left;
                top = posRange.top;
            } catch (posError) {
                console.log(`[insertTextBox] Warning: Could not parse position "${action.target}", using default`);
            }
        }
        
        // Create text box (rectangle shape with text)
        const textBox = sheet.shapes.addTextBox(data.text);
        
        // Set position
        textBox.left = left;
        textBox.top = top;
        
        // Set dimensions
        textBox.width = data.width || 150;
        textBox.height = data.height || 50;
        
        // Apply fill
        if (data.fill && data.fill !== "none") {
            textBox.fill.setSolidColor(data.fill);
        } else if (data.fill === "none") {
            textBox.fill.clear();
        }
        
        // Apply border
        if (data.lineColor === "none") {
            textBox.lineFormat.visible = false;
        } else if (data.lineColor) {
            textBox.lineFormat.color = data.lineColor;
        }
        
        // Set name
        if (data.name) textBox.name = data.name;
        
        await ctx.sync();
        
        // Apply text formatting (requires separate sync)
        if (data.fontSize || data.fontColor || data.horizontalAlignment || data.verticalAlignment) {
            const textRange = textBox.textFrame.textRange;
            if (data.fontSize) textRange.font.size = data.fontSize;
            if (data.fontColor) textRange.font.color = data.fontColor;
            if (data.horizontalAlignment) textBox.textFrame.horizontalAlignment = data.horizontalAlignment;
            if (data.verticalAlignment) textBox.textFrame.verticalAlignment = data.verticalAlignment;
            await ctx.sync();
        }
        
        console.log(`[insertTextBox] Successfully created text box at position (${left}, ${top})`);
    } catch (error) {
        console.log(`[insertTextBox] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Formats an existing shape
 */
async function formatShape(ctx, sheet, target, data) {
    console.log(`[formatShape] Formatting shape "${target}"`);
    
    if (!target) {
        throw new Error("formatShape requires a shape name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        
        // Get shape
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load("isNullObject");
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        // Apply fill
        if (options.fill !== undefined) {
            if (options.fill === "none") {
                shape.fill.clear();
            } else {
                shape.fill.setSolidColor(options.fill);
            }
        }
        
        // Apply transparency (clamped to 0-1 range)
        if (options.transparency !== undefined) {
            shape.fill.transparency = Math.max(0, Math.min(1, options.transparency));
        }
        
        // Apply line format
        if (options.lineColor !== undefined) {
            if (options.lineColor === "none") {
                shape.lineFormat.visible = false;
            } else {
                shape.lineFormat.visible = true;
                shape.lineFormat.color = options.lineColor;
            }
        }
        if (options.lineWeight !== undefined) {
            shape.lineFormat.weight = options.lineWeight;
        }
        if (options.lineStyle !== undefined) {
            shape.lineFormat.dashStyle = options.lineStyle;
        }
        
        // Apply dimensions
        if (options.width !== undefined) shape.width = options.width;
        if (options.height !== undefined) shape.height = options.height;
        if (options.rotation !== undefined) shape.rotation = options.rotation;
        
        await ctx.sync();
        
        console.log(`[formatShape] Successfully formatted shape "${target}"`);
    } catch (error) {
        console.log(`[formatShape] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Deletes a shape by name
 */
async function deleteShape(ctx, sheet, target) {
    console.log(`[deleteShape] Deleting shape "${target}"`);
    
    if (!target) {
        throw new Error("deleteShape requires a shape name in target");
    }
    
    try {
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load(["isNullObject", "name"]);
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        shape.delete();
        await ctx.sync();
        
        console.log(`[deleteShape] Successfully deleted shape "${target}"`);
    } catch (error) {
        console.log(`[deleteShape] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Groups multiple shapes together
 */
async function groupShapes(ctx, sheet, action) {
    console.log(`[groupShapes] Grouping shapes: ${action.target}`);
    
    if (!action.target) {
        throw new Error("groupShapes requires shape names in target (comma-separated)");
    }
    
    try {
        const data = action.data ? JSON.parse(action.data) : {};
        
        // Parse shape names
        const shapeNames = action.target.split(",").map(s => s.trim()).filter(s => s);
        
        if (shapeNames.length < 2) {
            throw new Error("groupShapes requires at least 2 shapes to group");
        }
        
        // Get all shapes and collect their IDs
        const shapes = [];
        for (const name of shapeNames) {
            const shape = sheet.shapes.getItemOrNullObject(name);
            shape.load(["isNullObject", "id"]);
            shapes.push({ name, shape });
        }
        await ctx.sync();
        
        // Validate all shapes exist
        const shapeIds = [];
        for (const { name, shape } of shapes) {
            if (shape.isNullObject) {
                throw new Error(`Shape "${name}" not found`);
            }
            shapeIds.push(shape.id);
        }
        
        // Create group
        const group = sheet.shapes.addGroup(shapeIds);
        
        // Set group name if provided
        if (data.groupName) {
            group.name = data.groupName;
        }
        
        await ctx.sync();
        
        console.log(`[groupShapes] Successfully grouped ${shapeNames.length} shapes`);
    } catch (error) {
        console.log(`[groupShapes] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Arranges shape z-order (layering)
 */
async function arrangeShapes(ctx, sheet, target, data) {
    console.log(`[arrangeShapes] Arranging shape "${target}"`);
    
    if (!target) {
        throw new Error("arrangeShapes requires a shape name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        
        if (!options.order) {
            throw new Error("arrangeShapes requires an order option: bringToFront, sendToBack, bringForward, sendBackward");
        }
        
        // Get shape
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load("isNullObject");
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        // Map order to Excel enum
        const orderMap = {
            "bringToFront": "BringToFront",
            "sendToBack": "SendToBack",
            "bringForward": "BringForward",
            "sendBackward": "SendBackward"
        };
        
        const excelOrder = orderMap[options.order];
        if (!excelOrder) {
            throw new Error(`Invalid order "${options.order}". Valid options: bringToFront, sendToBack, bringForward, sendBackward`);
        }
        
        shape.incrementZOrder(excelOrder);
        await ctx.sync();
        
        console.log(`[arrangeShapes] Successfully applied ${options.order} to shape "${target}"`);
    } catch (error) {
        console.log(`[arrangeShapes] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Ungroups a shape group back into individual shapes
 */
async function ungroupShapes(ctx, sheet, target) {
    console.log(`[ungroupShapes] Ungrouping shape group "${target}"`);
    
    if (!target) {
        throw new Error("ungroupShapes requires a group name in target");
    }
    
    try {
        // Get the shape group
        const shape = sheet.shapes.getItemOrNullObject(target);
        shape.load(["isNullObject", "type"]);
        await ctx.sync();
        
        if (shape.isNullObject) {
            throw new Error(`Shape "${target}" not found`);
        }
        
        // Verify it's a group
        if (shape.type !== "Group") {
            throw new Error(`Shape "${target}" is not a group. Only grouped shapes can be ungrouped.`);
        }
        
        // Get the group and ungroup it
        const group = shape.group;
        group.ungroup();
        await ctx.sync();
        
        console.log(`[ungroupShapes] Successfully ungrouped "${target}"`);
    } catch (error) {
        console.log(`[ungroupShapes] Error: ${error.message}`);
        throw error;
    }
}

// ============================================================================
// Comments and Notes Operations
// ============================================================================

/**
 * Adds a threaded comment to a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function addComment(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[addComment] Adding comment to cell "${target}"`);
    
    if (!target) {
        throw new Error("addComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            logWarn(`[addComment] Warning: Sheet is protected, comment may not be added`);
        }
        
        const options = data ? JSON.parse(data) : {};
        const content = options.content || options.text || "";
        const contentType = options.contentType === "Mention" ? Excel.ContentType.mention : Excel.ContentType.plain;
        
        if (!content) {
            throw new Error("addComment requires content in data");
        }
        
        // Add comment to the cell
        const comment = sheet.comments.add(target, content, contentType);
        comment.load(["id", "authorName", "creationDate"]);
        await ctx.sync();
        
        logInfo(`[addComment] Successfully added comment (ID: ${comment.id}) to "${target}" by ${comment.authorName}`);
    } catch (error) {
        logError(`[addComment] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Adds a legacy note to a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function addNote(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[addNote] Adding note to cell "${target}"`);
    
    if (!target) {
        throw new Error("addNote requires a cell address in target");
    }
    
    try {
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            logWarn(`[addNote] Warning: Sheet is protected, note may not be added`);
        }
        
        const options = data ? JSON.parse(data) : {};
        const text = options.text || options.content || "";
        
        if (!text) {
            throw new Error("addNote requires text in data");
        }
        
        // Get the range and check if note API is available
        const range = sheet.getRange(target);
        
        // Check if note property exists
        if (range.note === undefined) {
            throw new Error("Notes API is not available in this Excel version");
        }
        
        range.note = text;
        await ctx.sync();
        
        logInfo(`[addNote] Successfully added note to "${target}"`);
    } catch (error) {
        logError(`[addNote] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Edits an existing comment on a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function editComment(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[editComment] Editing comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("editComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot modify comments on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const content = options.content || options.text || "";
        
        if (!content) {
            throw new Error("editComment requires content in data");
        }
        
        // Get comment by cell address
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            throw new Error(`No comment found at cell "${target}"`);
        }
        
        // Update the comment content
        comment.content = content;
        await ctx.sync();
        
        logInfo(`[editComment] Successfully edited comment at "${target}"`);
    } catch (error) {
        logError(`[editComment] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Edits an existing note on a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function editNote(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[editNote] Editing note at cell "${target}"`);
    
    if (!target) {
        throw new Error("editNote requires a cell address in target");
    }
    
    try {
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot modify notes on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const text = options.text || options.content || "";
        
        if (!text) {
            throw new Error("editNote requires text in data");
        }
        
        // Get the range and update the note
        const range = sheet.getRange(target);
        
        // Check if note property exists
        if (range.note === undefined) {
            throw new Error("Notes API is not available in this Excel version");
        }
        
        range.note = text;
        await ctx.sync();
        
        logInfo(`[editNote] Successfully edited note at "${target}"`);
    } catch (error) {
        logError(`[editNote] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Deletes a comment from a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target
 */
async function deleteComment(ctx, sheet, action) {
    const { target } = action;
    logDebug(`[deleteComment] Deleting comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("deleteComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot delete comments on a protected sheet");
        }
        
        // Get comment by cell address using getItemByCell
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            logDebug(`[deleteComment] No comment found at "${target}" - nothing to delete`);
            return;
        }
        
        // Delete the comment
        comment.delete();
        await ctx.sync();
        
        logInfo(`[deleteComment] Successfully deleted comment at "${target}"`);
    } catch (error) {
        logError(`[deleteComment] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Deletes a note from a cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target
 */
async function deleteNote(ctx, sheet, action) {
    const { target } = action;
    logDebug(`[deleteNote] Deleting note at cell "${target}"`);
    
    if (!target) {
        throw new Error("deleteNote requires a cell address in target");
    }
    
    try {
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot delete notes on a protected sheet");
        }
        
        // Get the range and clear the note
        const range = sheet.getRange(target);
        
        // Check if note property exists
        if (range.note === undefined) {
            throw new Error("Notes API is not available in this Excel version");
        }
        
        range.note = "";
        await ctx.sync();
        
        logInfo(`[deleteNote] Successfully deleted note at "${target}"`);
    } catch (error) {
        logError(`[deleteNote] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Adds a reply to an existing comment thread
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function replyToComment(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[replyToComment] Adding reply to comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("replyToComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot reply to comments on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const content = options.content || options.text || "";
        const contentType = options.contentType === "Mention" ? Excel.ContentType.mention : Excel.ContentType.plain;
        
        if (!content) {
            throw new Error("replyToComment requires content in data");
        }
        
        // Get the parent comment by cell address
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            throw new Error(`No comment found at cell "${target}" to reply to`);
        }
        
        // Add reply to the comment
        const reply = comment.replies.add(content, contentType);
        reply.load(["id", "authorName", "creationDate"]);
        await ctx.sync();
        
        logInfo(`[replyToComment] Successfully added reply (ID: ${reply.id}) to comment at "${target}"`);
    } catch (error) {
        logError(`[replyToComment] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Resolves or reopens a comment thread
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function resolveComment(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[resolveComment] Resolving/reopening comment at cell "${target}"`);
    
    if (!target) {
        throw new Error("resolveComment requires a cell address in target");
    }
    
    try {
        // Check if comments API is available
        if (!sheet.comments) {
            throw new Error("Comments API is not available in this Excel version");
        }
        
        // Check worksheet protection
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            throw new Error("Cannot modify comments on a protected sheet");
        }
        
        const options = data ? JSON.parse(data) : {};
        const resolved = options.resolved !== false; // Default to true
        
        // Get comment by cell address
        const comment = sheet.comments.getItemByCell(target);
        comment.load("isNullObject");
        await ctx.sync();
        
        if (comment.isNullObject) {
            throw new Error(`No comment found at cell "${target}"`);
        }
        
        // Set resolution status
        comment.resolved = resolved;
        await ctx.sync();
        
        logInfo(`[resolveComment] Successfully ${resolved ? "resolved" : "reopened"} comment at "${target}"`);
    } catch (error) {
        logError(`[resolveComment] Error: ${error.message}`);
        throw error;
    }
}

// ============================================================================
// Sparkline Operations
// ============================================================================

/**
 * Checks if sparkline API is supported in the current Excel version
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @returns {Promise<boolean>} True if sparklines are supported
 */
async function isSparklineSupported(ctx, sheet) {
    try {
        if (!sheet.sparklineGroups) {
            return false;
        }
        sheet.sparklineGroups.load("count");
        await ctx.sync();
        return true;
    } catch (error) {
        return false;
    }
}

/**
 * Creates sparkline(s) at the specified location
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function createSparkline(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[createSparkline] Creating sparkline at "${target}"`);
    
    if (!target) {
        throw new Error("createSparkline requires a cell address in target");
    }
    
    try {
        const supported = await isSparklineSupported(ctx, sheet);
        if (!supported) {
            logWarn(`[createSparkline] Sparklines require ExcelApi 1.10+. Consider using data bars.`);
            throw new Error("Sparklines are not available in this Excel version (requires ExcelApi 1.10+). Consider using data bars as an alternative.");
        }
        
        sheet.protection.load("protected");
        await ctx.sync();
        
        if (sheet.protection.protected) {
            logWarn(`[createSparkline] Warning: Sheet is protected`);
        }
        
        let options = {};
        if (data) {
            if (typeof data === "string") {
                try {
                    options = JSON.parse(data);
                } catch (e) {
                    logWarn(`[createSparkline] Warning: Could not parse data JSON`);
                }
            } else if (typeof data === "object") {
                options = data;
            }
        }
        
        const sparklineType = options.type || "Line";
        const sourceData = options.sourceData;
        
        if (!sourceData) {
            throw new Error("createSparkline requires sourceData in data (e.g., 'B2:F2')");
        }
        
        const typeMap = {
            "Line": Excel.SparklineType.line,
            "Column": Excel.SparklineType.column,
            "WinLoss": Excel.SparklineType.winLoss
        };
        
        const excelType = typeMap[sparklineType];
        if (!excelType) {
            throw new Error(`Invalid sparkline type "${sparklineType}". Valid types: Line, Column, WinLoss`);
        }
        
        const rangePattern = /^[A-Z]+\d+(:[A-Z]+\d+)?$/i;
        if (!rangePattern.test(sourceData)) {
            throw new Error(`Invalid sourceData range format "${sourceData}". Expected format: B2:F2 or C3:C20`);
        }
        
        const sparklineGroup = sheet.sparklineGroups.add(excelType, sourceData, target);
        
        if (options.axes && options.axes.horizontal !== undefined) {
            sparklineGroup.axes.horizontal.axis.visible = options.axes.horizontal;
        }
        
        if (options.markers && sparklineType === "Line") {
            const points = sparklineGroup.points;
            if (options.markers.high !== undefined) points.highPoint.visible = options.markers.high;
            if (options.markers.low !== undefined) points.lowPoint.visible = options.markers.low;
            if (options.markers.first !== undefined) points.firstPoint.visible = options.markers.first;
            if (options.markers.last !== undefined) points.lastPoint.visible = options.markers.last;
            if (options.markers.negative !== undefined) points.negativePoints.visible = options.markers.negative;
        }
        
        if (options.colors) {
            const hexPattern = /^#[0-9A-Fa-f]{6}$/;
            if (options.colors.series && hexPattern.test(options.colors.series)) {
                sparklineGroup.seriesColor = options.colors.series;
            }
            if (options.colors.negative && hexPattern.test(options.colors.negative)) {
                sparklineGroup.negativePointsColor = options.colors.negative;
            }
            if (options.colors.high && hexPattern.test(options.colors.high)) {
                sparklineGroup.highPointColor = options.colors.high;
            }
            if (options.colors.low && hexPattern.test(options.colors.low)) {
                sparklineGroup.lowPointColor = options.colors.low;
            }
        }
        
        await ctx.sync();
        
        logInfo(`[createSparkline] Successfully created ${sparklineType} sparkline at "${target}" from "${sourceData}"`);
    } catch (error) {
        logError(`[createSparkline] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Normalizes a cell address for comparison by uppercasing, removing $ signs,
 * and extracting only the address part (ignoring sheet name if target has none)
 * @param {string} address - Cell address to normalize
 * @param {boolean} hasSheetInTarget - Whether the target has a sheet name
 * @returns {string} Normalized address
 */
function normalizeSparklineAddress(address, hasSheetInTarget) {
    if (!address) return "";
    let normalized = address.toUpperCase().replace(/\$/g, "");
    // If target has no sheet name, strip sheet name from address for comparison
    if (!hasSheetInTarget && normalized.includes("!")) {
        normalized = normalized.split("!")[1] || normalized;
    }
    return normalized;
}

/**
 * Configures an existing sparkline's properties
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target and data
 */
async function configureSparkline(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[configureSparkline] Configuring sparkline at "${target}"`);
    
    if (!target) {
        throw new Error("configureSparkline requires a cell address in target");
    }
    
    try {
        const supported = await isSparklineSupported(ctx, sheet);
        if (!supported) {
            throw new Error("Sparklines are not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        let options = {};
        if (data) {
            if (typeof data === "string") {
                try {
                    options = JSON.parse(data);
                } catch (e) {
                    logWarn(`[configureSparkline] Warning: Could not parse data JSON`);
                }
            } else if (typeof data === "object") {
                options = data;
            }
        }
        
        // Normalize target address for comparison
        const hasSheetInTarget = target.includes("!");
        const normalizedTarget = normalizeSparklineAddress(target, hasSheetInTarget);
        
        // Load sparkline groups and batch load all sparkline locations
        sheet.sparklineGroups.load("items");
        await ctx.sync();
        
        // Batch load all sparkline locations before iterating
        for (const group of sheet.sparklineGroups.items) {
            group.load("sparklines/items/location");
        }
        await ctx.sync();
        
        // Batch load all location addresses
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                sparkline.location.load("address");
            }
        }
        await ctx.sync();
        
        // Find sparkline group at the target location using strict equality
        let foundSparkline = null;
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                const normalizedAddress = normalizeSparklineAddress(sparkline.location.address, hasSheetInTarget);
                if (normalizedAddress === normalizedTarget) {
                    foundSparkline = group;
                    break;
                }
            }
            if (foundSparkline) break;
        }
        
        if (!foundSparkline) {
            logWarn(`[configureSparkline] No sparkline found at "${target}"`);
            throw new Error(`No sparkline found at cell "${target}"`);
        }
        
        const hexPattern = /^#[0-9A-Fa-f]{6}$/;
        
        if (options.colors) {
            if (options.colors.series && hexPattern.test(options.colors.series)) {
                foundSparkline.seriesColor = options.colors.series;
            }
            if (options.colors.negative && hexPattern.test(options.colors.negative)) {
                foundSparkline.negativePointsColor = options.colors.negative;
            }
            if (options.colors.high && hexPattern.test(options.colors.high)) {
                foundSparkline.highPointColor = options.colors.high;
            }
            if (options.colors.low && hexPattern.test(options.colors.low)) {
                foundSparkline.lowPointColor = options.colors.low;
            }
        }
        
        if (options.markers) {
            const points = foundSparkline.points;
            if (options.markers.high !== undefined) points.highPoint.visible = options.markers.high;
            if (options.markers.low !== undefined) points.lowPoint.visible = options.markers.low;
            if (options.markers.first !== undefined) points.firstPoint.visible = options.markers.first;
            if (options.markers.last !== undefined) points.lastPoint.visible = options.markers.last;
            if (options.markers.negative !== undefined) points.negativePoints.visible = options.markers.negative;
        }
        
        if (options.axes && options.axes.horizontal !== undefined) {
            foundSparkline.axes.horizontal.axis.visible = options.axes.horizontal;
        }
        
        await ctx.sync();
        
        logInfo(`[configureSparkline] Successfully configured sparkline at "${target}"`);
    } catch (error) {
        logError(`[configureSparkline] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Deletes sparkline(s) at the specified location
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet object
 * @param {Object} action - Action object with target
 */
async function deleteSparkline(ctx, sheet, action) {
    const { target } = action;
    logDebug(`[deleteSparkline] Deleting sparkline at "${target}"`);
    
    if (!target) {
        throw new Error("deleteSparkline requires a cell address in target");
    }
    
    try {
        const supported = await isSparklineSupported(ctx, sheet);
        if (!supported) {
            throw new Error("Sparklines are not available in this Excel version (requires ExcelApi 1.10+)");
        }
        
        // Normalize target address for comparison
        const hasSheetInTarget = target.includes("!");
        const normalizedTarget = normalizeSparklineAddress(target, hasSheetInTarget);
        
        // Load sparkline groups and batch load all sparkline locations
        sheet.sparklineGroups.load("items");
        await ctx.sync();
        
        // Batch load all sparkline locations before iterating
        for (const group of sheet.sparklineGroups.items) {
            group.load("sparklines/items/location");
        }
        await ctx.sync();
        
        // Batch load all location addresses
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                sparkline.location.load("address");
            }
        }
        await ctx.sync();
        
        // Find sparkline groups at the target location using strict equality
        let deletedCount = 0;
        const groupsToDelete = [];
        
        for (const group of sheet.sparklineGroups.items) {
            for (const sparkline of group.sparklines.items) {
                const normalizedAddress = normalizeSparklineAddress(sparkline.location.address, hasSheetInTarget);
                if (normalizedAddress === normalizedTarget) {
                    groupsToDelete.push(group);
                    break;
                }
            }
        }
        
        for (const group of groupsToDelete) {
            group.delete();
            deletedCount++;
        }
        
        await ctx.sync();
        
        if (deletedCount === 0) {
            logWarn(`[deleteSparkline] No sparkline found at "${target}" - nothing to delete`);
        } else {
            logInfo(`[deleteSparkline] Successfully deleted ${deletedCount} sparkline group(s) at "${target}"`);
        }
    } catch (error) {
        logError(`[deleteSparkline] Error: ${error.message}`);
        throw error;
    }
}

// ============================================================================
// Worksheet Management Operations
// ============================================================================

/**
 * Renames a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet (may not be target)
 * @param {Object} action - Action object with target (old name) and data (newName)
 */
async function renameSheet(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[renameSheet] Renaming sheet "${target}"`);
    
    if (!target) {
        throw new Error("renameSheet requires a sheet name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const newName = options.newName;
        
        if (!newName) {
            throw new Error("renameSheet requires newName in data");
        }
        
        // Validate new name
        if (newName.length > 31) {
            throw new Error("Sheet name cannot exceed 31 characters");
        }
        
        const invalidChars = /[\\\/\?\*\[\]]/;
        if (invalidChars.test(newName)) {
            throw new Error("Sheet name cannot contain \\ / ? * [ ] characters");
        }
        
        if (newName.trim() === "") {
            throw new Error("Sheet name cannot be empty");
        }
        
        // Get the target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load("name");
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        // Check for duplicate name
        const existingSheet = ctx.workbook.worksheets.getItemOrNullObject(newName);
        existingSheet.load("name");
        await ctx.sync();
        
        if (!existingSheet.isNullObject && existingSheet.name.toLowerCase() !== target.toLowerCase()) {
            throw new Error(`A sheet named "${newName}" already exists`);
        }
        
        // Rename the sheet
        targetSheet.name = newName;
        await ctx.sync();
        
        logInfo(`[renameSheet] Successfully renamed sheet "${target}" to "${newName}"`);
    } catch (error) {
        logError(`[renameSheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Moves a worksheet to a new position
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target and data (position, referenceSheet)
 */
async function moveSheet(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[moveSheet] Moving sheet "${target}"`);
    
    if (!target) {
        throw new Error("moveSheet requires a sheet name in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const position = options.position || "last";
        const referenceSheet = options.referenceSheet;
        
        // Get the target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load("name");
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        // Get all sheets to determine positions
        const sheets = ctx.workbook.worksheets;
        sheets.load("items");
        await ctx.sync();
        
        let newPosition;
        
        switch (position.toLowerCase()) {
            case "first":
                newPosition = 0;
                break;
            case "last":
                newPosition = sheets.items.length - 1;
                break;
            case "before":
            case "after":
                if (!referenceSheet) {
                    throw new Error(`moveSheet with position "${position}" requires referenceSheet`);
                }
                const refSheet = ctx.workbook.worksheets.getItemOrNullObject(referenceSheet);
                refSheet.load("position");
                await ctx.sync();
                
                if (refSheet.isNullObject) {
                    throw new Error(`Reference sheet "${referenceSheet}" not found`);
                }
                
                newPosition = position.toLowerCase() === "before" 
                    ? refSheet.position 
                    : refSheet.position + 1;
                break;
            default:
                throw new Error(`Invalid position "${position}". Use: first, last, before, after`);
        }
        
        // Move the sheet
        targetSheet.position = newPosition;
        await ctx.sync();
        
        logInfo(`[moveSheet] Successfully moved sheet "${target}" to position ${newPosition}`);
    } catch (error) {
        logError(`[moveSheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Hides a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name)
 */
async function hideSheet(ctx, sheet, action) {
    const { target } = action;
    logDebug(`[hideSheet] Hiding sheet "${target}"`);
    
    if (!target) {
        throw new Error("hideSheet requires a sheet name in target");
    }
    
    try {
        // First, get the target sheet and check its current visibility
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load(["name", "visibility"]);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        // If already hidden, return early without checking visible sheet count
        if (targetSheet.visibility !== Excel.SheetVisibility.visible) {
            logWarn(`[hideSheet] Sheet "${target}" is already hidden`);
            return;
        }
        
        // Only check visible sheet count if we're about to hide a visible sheet
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/visibility");
        await ctx.sync();
        
        const visibleSheets = sheets.items.filter(s => s.visibility === Excel.SheetVisibility.visible);
        
        if (visibleSheets.length <= 1) {
            throw new Error("Cannot hide the only visible sheet");
        }
        
        // Hide the sheet
        targetSheet.visibility = Excel.SheetVisibility.hidden;
        await ctx.sync();
        
        logInfo(`[hideSheet] Successfully hidden sheet "${target}"`);
    } catch (error) {
        logError(`[hideSheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unhides a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (sheet name)
 */
async function unhideSheet(ctx, sheet, action) {
    const { target } = action;
    logDebug(`[unhideSheet] Unhiding sheet "${target}"`);
    
    if (!target) {
        throw new Error("unhideSheet requires a sheet name in target");
    }
    
    try {
        // Get the target sheet
        const targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
        targetSheet.load(["name", "visibility"]);
        await ctx.sync();
        
        if (targetSheet.isNullObject) {
            throw new Error(`Sheet "${target}" not found`);
        }
        
        if (targetSheet.visibility === Excel.SheetVisibility.visible) {
            logWarn(`[unhideSheet] Sheet "${target}" is already visible`);
            return;
        }
        
        // Unhide the sheet
        targetSheet.visibility = Excel.SheetVisibility.visible;
        await ctx.sync();
        
        logInfo(`[unhideSheet] Successfully unhidden sheet "${target}"`);
    } catch (error) {
        logError(`[unhideSheet] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Freezes panes at a specified cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet to freeze
 * @param {Object} action - Action object with target (cell) and data (freezeType)
 */
async function freezePanes(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[freezePanes] Freezing panes at "${target}"`);
    
    if (!target) {
        throw new Error("freezePanes requires a cell address in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const freezeType = options.freezeType || "both";
        
        // Get the freeze range
        const range = sheet.getRange(target);
        range.load(["rowIndex", "columnIndex"]);
        await ctx.sync();
        
        const rowCount = range.rowIndex;
        const colCount = range.columnIndex;
        
        // Apply freeze based on type
        switch (freezeType.toLowerCase()) {
            case "rows":
                if (rowCount > 0) {
                    sheet.freezePanes.freezeRows(rowCount);
                } else {
                    throw new Error("Cannot freeze rows: target cell is in row 1");
                }
                break;
            case "columns":
                if (colCount > 0) {
                    sheet.freezePanes.freezeColumns(colCount);
                } else {
                    throw new Error("Cannot freeze columns: target cell is in column A");
                }
                break;
            case "both":
            default:
                sheet.freezePanes.freezeAt(range);
                break;
        }
        
        await ctx.sync();
        
        logInfo(`[freezePanes] Successfully froze panes at "${target}" (type: ${freezeType})`);
    } catch (error) {
        logError(`[freezePanes] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Unfreezes all panes on a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target ("current" or sheet name)
 */
async function unfreezePane(ctx, sheet, action) {
    const { target } = action;
    logDebug(`[unfreezePane] Unfreezing panes on "${target}"`);
    
    try {
        let targetSheet = sheet;
        
        if (target && target.toLowerCase() !== "current") {
            targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
            targetSheet.load("name");
            await ctx.sync();
            
            if (targetSheet.isNullObject) {
                throw new Error(`Sheet "${target}" not found`);
            }
        }
        
        // Unfreeze all panes
        targetSheet.freezePanes.unfreeze();
        await ctx.sync();
        
        logInfo(`[unfreezePane] Successfully unfroze panes`);
    } catch (error) {
        logError(`[unfreezePane] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Sets the zoom level for a worksheet
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target and data (zoomLevel)
 */
async function setZoom(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[setZoom] Setting zoom on "${target}"`);
    
    try {
        const options = data ? JSON.parse(data) : {};
        const zoomLevel = options.zoomLevel;
        
        if (zoomLevel === undefined || zoomLevel === null) {
            throw new Error("setZoom requires zoomLevel in data");
        }
        
        const zoom = parseInt(zoomLevel);
        if (isNaN(zoom) || zoom < 10 || zoom > 400) {
            throw new Error("zoomLevel must be between 10 and 400");
        }
        
        let targetSheet = sheet;
        
        if (target && target.toLowerCase() !== "current") {
            targetSheet = ctx.workbook.worksheets.getItemOrNullObject(target);
            targetSheet.load("name");
            await ctx.sync();
            
            if (targetSheet.isNullObject) {
                throw new Error(`Sheet "${target}" not found`);
            }
        }
        
        // Set zoom level via the worksheet's pageLayout (Office.js workaround)
        try {
            if (targetSheet.view) {
                targetSheet.view.zoom = zoom;
            } else {
                targetSheet.pageLayout.zoom = { scale: zoom };
            }
            await ctx.sync();
        } catch (zoomError) {
            targetSheet.pageLayout.zoom = { scale: zoom };
            await ctx.sync();
        }
        
        logInfo(`[setZoom] Successfully set zoom to ${zoom}%`);
    } catch (error) {
        logError(`[setZoom] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Splits panes at a specified cell
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Worksheet to split
 * @param {Object} action - Action object with target (cell) and data (horizontal, vertical)
 */
async function splitPane(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[splitPane] Splitting panes at "${target}"`);
    
    if (!target) {
        throw new Error("splitPane requires a cell address in target");
    }
    
    try {
        const options = data ? JSON.parse(data) : {};
        const horizontal = options.horizontal !== false;
        const vertical = options.vertical !== false;
        
        // Get the split position
        const range = sheet.getRange(target);
        range.load(["rowIndex", "columnIndex"]);
        await ctx.sync();
        
        // Note: Office.js has limited support for split panes
        // Using freezeAt as a workaround which creates a similar effect
        if (horizontal && vertical) {
            sheet.freezePanes.freezeAt(range);
        } else if (horizontal && !vertical) {
            // Horizontal-only split: guard against row 1
            if (range.rowIndex === 0) {
                throw new Error("Cannot split horizontally at row 1; choose a cell below the first row");
            }
            sheet.freezePanes.freezeRows(range.rowIndex);
        } else if (vertical && !horizontal) {
            // Vertical-only split: guard against column A
            if (range.columnIndex === 0) {
                throw new Error("Cannot split vertically at column A; choose a cell to the right of the first column");
            }
            sheet.freezePanes.freezeColumns(range.columnIndex);
        }
        
        await ctx.sync();
        
        logInfo(`[splitPane] Successfully split panes at "${target}" (H:${horizontal}, V:${vertical})`);
        logDebug(`[splitPane] Note: Using freeze panes as Office.js split pane API is limited`);
    } catch (error) {
        logError(`[splitPane] Error: ${error.message}`);
        throw error;
    }
}

/**
 * Creates a custom view (limited API support)
 * @param {Excel.RequestContext} ctx - Excel context
 * @param {Excel.Worksheet} sheet - Current worksheet
 * @param {Object} action - Action object with target (view name) and data (options)
 */
async function createView(ctx, sheet, action) {
    const { target, data } = action;
    logDebug(`[createView] Creating view "${target}"`);
    
    if (!target) {
        throw new Error("createView requires a view name in target");
    }
    
    try {
        // Note: Office.js has very limited support for custom views
        const options = data ? JSON.parse(data) : {};
        
        // Load current view state for documentation
        sheet.load("name");
        await ctx.sync();
        
        logInfo(`[createView] Custom view "${target}" requested for sheet "${sheet.name}"`);
        logDebug(`[createView] Options: includeHidden=${options.includeHidden}, includePrint=${options.includePrint}, includeFilter=${options.includeFilter}`);
        logWarn(`[createView] Note: Office.js has limited custom view API support. Use Excel UI: View > Custom Views > Add`);
        
        console.warn(`Custom view "${target}" creation requires manual Excel UI. Go to View > Custom Views > Add.`);
    } catch (error) {
        logError(`[createView] Error: ${error.message}`);
        throw error;
    }
}
