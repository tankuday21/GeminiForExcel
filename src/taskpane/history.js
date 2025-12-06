/**
 * History Module for Undo Functionality
 * Manages action history and undo data
 */

const MAX_ENTRIES = 20;

/**
 * @typedef {Object} HistoryEntry
 * @property {string} id - Unique identifier
 * @property {string} type - Action type
 * @property {string} target - Target range address
 * @property {number} timestamp - Unix timestamp
 * @property {Object} undoData - Data to restore previous state
 */

/**
 * Generates a unique ID for history entries
 * @returns {string} Unique ID
 */
function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

/**
 * Creates a new history entry
 * @param {Object} action - The action that was applied
 * @param {Object} undoData - The captured undo data
 * @returns {HistoryEntry} The created history entry
 */
function createHistoryEntry(action, undoData) {
    return {
        id: generateId(),
        type: action.type,
        target: action.target,
        timestamp: Date.now(),
        undoData: undoData
    };
}

/**
 * Adds an entry to the history array (prepends to front)
 * @param {HistoryEntry[]} entries - Current history entries
 * @param {HistoryEntry} entry - Entry to add
 * @param {number} maxEntries - Maximum entries to retain
 * @returns {HistoryEntry[]} Updated history entries
 */
function addToHistory(entries, entry, maxEntries = MAX_ENTRIES) {
    const newEntries = [entry, ...entries];
    // Enforce max limit by removing oldest entries
    if (newEntries.length > maxEntries) {
        return newEntries.slice(0, maxEntries);
    }
    return newEntries;
}

/**
 * Removes the most recent entry from history
 * @param {HistoryEntry[]} entries - Current history entries
 * @returns {{ entries: HistoryEntry[], removed: HistoryEntry|null }} Updated entries and removed entry
 */
function removeFromHistory(entries) {
    if (!entries || entries.length === 0) {
        return { entries: [], removed: null };
    }
    const [removed, ...remaining] = entries;
    return { entries: remaining, removed };
}

/**
 * Gets the most recent entry without removing it
 * @param {HistoryEntry[]} entries - Current history entries
 * @returns {HistoryEntry|null} Most recent entry or null
 */
function getLatestEntry(entries) {
    if (!entries || entries.length === 0) return null;
    return entries[0];
}

/**
 * Checks if history has any entries
 * @param {HistoryEntry[]} entries - Current history entries
 * @returns {boolean} True if history has entries
 */
function hasHistory(entries) {
    if (!entries || entries.length === 0) return false;
    return true;
}

/**
 * Formats a timestamp as relative time
 * @param {number} timestamp - Unix timestamp
 * @returns {string} Formatted relative time (e.g., "2 min ago")
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
 * Renders a single history entry as HTML
 * @param {HistoryEntry} entry - The entry to render
 * @param {Function} getIcon - Function to get icon for action type
 * @returns {string} HTML string
 */
function renderHistoryEntry(entry, getIcon) {
    const icon = getIcon ? getIcon(entry.type) : '';
    const timeStr = formatRelativeTime(entry.timestamp);
    
    const typeLabels = {
        formula: "Formula",
        values: "Values",
        format: "Format",
        chart: "Chart",
        validation: "Dropdown",
        sort: "Sort",
        autofill: "Autofill"
    };
    const label = typeLabels[entry.type] || entry.type;
    
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
}

/**
 * Renders the history panel content
 * @param {HistoryEntry[]} entries - History entries to display
 * @param {Function} getIcon - Function to get icon for action type
 * @returns {string} HTML string
 */
function renderHistoryList(entries, getIcon) {
    if (!entries || entries.length === 0) {
        return '<div class="history-empty">No actions yet</div>';
    }
    
    return entries.map(entry => renderHistoryEntry(entry, getIcon)).join('');
}

// ES Module exports
export {
    MAX_ENTRIES,
    generateId,
    createHistoryEntry,
    addToHistory,
    removeFromHistory,
    getLatestEntry,
    hasHistory,
    formatRelativeTime,
    renderHistoryEntry,
    renderHistoryList
};

// CommonJS exports for testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        MAX_ENTRIES,
        generateId,
        createHistoryEntry,
        addToHistory,
        removeFromHistory,
        getLatestEntry,
        hasHistory,
        formatRelativeTime,
        renderHistoryEntry,
        renderHistoryList
    };
}
