/**
 * Diagnostics Module
 * Provides logging and debug functionality for the add-in
 */

// ============================================================================
// Configuration
// ============================================================================

const DIAG_CONFIG = {
    MAX_LOGS: 100,
    STORAGE_KEY: "excel_copilot_debug_mode"
};

// ============================================================================
// State
// ============================================================================

let logs = [];
let debugMode = false;
let updateCallback = null;

// ============================================================================
// Initialization
// ============================================================================

/**
 * Initializes the diagnostics system
 * @param {Function} onUpdate - Callback when logs are updated
 */
function initDiagnostics(onUpdate) {
    updateCallback = onUpdate;
    
    // Load debug mode preference
    try {
        debugMode = localStorage.getItem(DIAG_CONFIG.STORAGE_KEY) === "true";
    } catch (e) {
        debugMode = false;
    }
}

// ============================================================================
// Logging
// ============================================================================

/**
 * Adds a log entry
 * @param {string} message - Log message
 * @param {string} level - Log level: "info", "warn", "error", "debug"
 * @param {Object} data - Optional additional data
 */
function log(message, level = "info", data = null) {
    const entry = {
        id: Date.now().toString(36) + Math.random().toString(36).substr(2, 5),
        timestamp: Date.now(),
        level,
        message,
        data
    };

    logs.unshift(entry);
    
    // Enforce max limit
    if (logs.length > DIAG_CONFIG.MAX_LOGS) {
        logs = logs.slice(0, DIAG_CONFIG.MAX_LOGS);
    }
    
    // Console output based on level
    if (level === "error") {
        console.error(`[Copilot] ${message}`, data || "");
    } else if (level === "warn") {
        console.warn(`[Copilot] ${message}`, data || "");
    } else if (debugMode || level !== "debug") {
        console.log(`[Copilot] ${message}`, data || "");
    }
    
    // Notify UI
    if (updateCallback) {
        updateCallback(logs);
    }
}

/**
 * Logs an info message
 * @param {string} message - Log message
 * @param {Object} data - Optional additional data
 */
function logInfo(message, data = null) {
    log(message, "info", data);
}

/**
 * Logs a warning message
 * @param {string} message - Log message
 * @param {Object} data - Optional additional data
 */
function logWarn(message, data = null) {
    log(message, "warn", data);
}

/**
 * Logs an error message
 * @param {string} message - Log message
 * @param {Object} data - Optional additional data
 */
function logError(message, data = null) {
    log(message, "error", data);
}

/**
 * Logs a debug message (only shown in debug mode)
 * @param {string} message - Log message
 * @param {Object} data - Optional additional data
 */
function logDebug(message, data = null) {
    log(message, "debug", data);
}

// ============================================================================
// Log Management
// ============================================================================

/**
 * Gets all logs
 * @returns {Array} Array of log entries
 */
function getLogs() {
    return [...logs];
}

/**
 * Gets recent logs
 * @param {number} count - Number of logs to return
 * @returns {Array} Array of recent log entries
 */
function getRecentLogs(count = 20) {
    return logs.slice(0, count);
}

/**
 * Clears all logs
 */
function clearLogs() {
    logs = [];
    if (updateCallback) {
        updateCallback(logs);
    }
}

/**
 * Filters logs by level
 * @param {string} level - Level to filter by
 * @returns {Array} Filtered log entries
 */
function filterLogsByLevel(level) {
    return logs.filter(entry => entry.level === level);
}

// ============================================================================
// Debug Mode
// ============================================================================

/**
 * Enables debug mode
 */
function enableDebugMode() {
    debugMode = true;
    try {
        localStorage.setItem(DIAG_CONFIG.STORAGE_KEY, "true");
    } catch (e) {
        // Storage not available
    }
    logInfo("Debug mode enabled");
}

/**
 * Disables debug mode
 */
function disableDebugMode() {
    debugMode = false;
    try {
        localStorage.setItem(DIAG_CONFIG.STORAGE_KEY, "false");
    } catch (e) {
        // Storage not available
    }
    logInfo("Debug mode disabled");
}

/**
 * Toggles debug mode
 * @returns {boolean} New debug mode state
 */
function toggleDebugMode() {
    if (debugMode) {
        disableDebugMode();
    } else {
        enableDebugMode();
    }
    return debugMode;
}

/**
 * Checks if debug mode is enabled
 * @returns {boolean} Debug mode state
 */
function isDebugMode() {
    return debugMode;
}

// ============================================================================
// Formatting
// ============================================================================

/**
 * Formats a timestamp as relative time
 * @param {number} timestamp - Unix timestamp
 * @returns {string} Formatted time string
 */
function formatLogTime(timestamp) {
    const now = Date.now();
    const diff = now - timestamp;
    
    if (diff < 1000) return "just now";
    if (diff < 60000) return `${Math.floor(diff / 1000)}s ago`;
    if (diff < 3600000) return `${Math.floor(diff / 60000)}m ago`;
    
    const date = new Date(timestamp);
    return date.toLocaleTimeString();
}

/**
 * Renders a single log entry as HTML
 * @param {Object} entry - Log entry
 * @returns {string} HTML string
 */
function renderLogEntry(entry) {
    const levelIcons = {
        info: "ℹ️",
        warn: "⚠️",
        error: "❌",
        debug: "🔍"
    };
    
    const levelClasses = {
        info: "log-info",
        warn: "log-warn",
        error: "log-error",
        debug: "log-debug"
    };
    
    const icon = levelIcons[entry.level] || "📝";
    const levelClass = levelClasses[entry.level] || "";
    const timeStr = formatLogTime(entry.timestamp);
    
    let dataStr = "";
    if (entry.data && debugMode) {
        try {
            dataStr = `<pre class="log-data">${JSON.stringify(entry.data, null, 2)}</pre>`;
        } catch (e) {
            dataStr = `<pre class="log-data">${String(entry.data)}</pre>`;
        }
    }
    
    return `
        <div class="log-entry ${levelClass}" data-id="${entry.id}">
            <span class="log-icon">${icon}</span>
            <span class="log-message">${escapeHtml(entry.message)}</span>
            <span class="log-time">${timeStr}</span>
            ${dataStr}
        </div>
    `;
}

/**
 * Renders the diagnostics panel content
 * @param {Array} entries - Log entries to display
 * @returns {string} HTML string
 */
function renderDiagnosticsPanel(entries) {
    if (!entries || entries.length === 0) {
        return '<div class="log-empty">No logs yet</div>';
    }
    
    return entries.map(entry => renderLogEntry(entry)).join("");
}

/**
 * Escapes HTML special characters
 * @param {string} text - Text to escape
 * @returns {string} Escaped text
 */
function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}

// ============================================================================
// Exports
// ============================================================================

export {
    initDiagnostics,
    log,
    logInfo,
    logWarn,
    logError,
    logDebug,
    getLogs,
    getRecentLogs,
    clearLogs,
    filterLogsByLevel,
    enableDebugMode,
    disableDebugMode,
    toggleDebugMode,
    isDebugMode,
    formatLogTime,
    renderLogEntry,
    renderDiagnosticsPanel
};
