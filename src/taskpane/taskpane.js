/*
 * Excel AI Copilot - Main Application
 * Provides AI-powered assistance for Excel operations
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
// State Management
// ============================================================================
let state = {
  apiKey: "",
  lastResponse: null,
  pendingAction: null,
  selectionData: null,
  selectionAddress: "",
  chatHistory: []
};

// ============================================================================
// Initialization
// ============================================================================
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeApp();
  }
});

function initializeApp() {
  // Load saved API key
  state.apiKey = localStorage.getItem(CONFIG.STORAGE_KEY) || "";
  
  // Bind event listeners
  bindEventListeners();
  
  // Initial selection refresh
  refreshSelection();
  
  // Set up selection change handler
  setupSelectionHandler();
}

function bindEventListeners() {
  // Quick action buttons
  document.querySelectorAll(".action-btn").forEach(btn => {
    btn.addEventListener("click", () => handleQuickAction(btn.dataset.action));
  });
  
  // Main controls
  document.getElementById("btnRefresh")?.addEventListener("click", refreshSelection);
  document.getElementById("btnSend")?.addEventListener("click", sendMessage);
  document.getElementById("btnApply")?.addEventListener("click", applyChanges);
  
  // Input handling
  const promptInput = document.getElementById("promptInput");
  promptInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
  
  // Settings modal
  document.getElementById("btnSaveSettings")?.addEventListener("click", saveSettings);
  document.getElementById("btnCloseSettings")?.addEventListener("click", closeSettings);
  
  // Toggle selection context
  document.getElementById("useSelection")?.addEventListener("change", refreshSelection);
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
  const useSelection = document.getElementById("useSelection")?.checked;
  if (useSelection) {
    await refreshSelection();
  }
}

// ============================================================================
// Selection & Context Management
// ============================================================================
async function refreshSelection() {
  const useSelection = document.getElementById("useSelection")?.checked;
  const infoEl = document.getElementById("selectionInfo");
  
  if (!useSelection) {
    infoEl.innerHTML = '<span class="info-label">Selection context disabled</span>';
    state.selectionData = null;
    state.selectionAddress = "";
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values", "rowCount", "columnCount", "numberFormat"]);
      await context.sync();
      
      state.selectionAddress = range.address;
      state.selectionData = range.values;
      
      const rows = range.rowCount;
      const cols = range.columnCount;
      const preview = generateDataPreview(range.values, 3, 4);
      
      infoEl.innerHTML = `
        <strong>${range.address}</strong> (${rows} rows × ${cols} cols)
        <div style="margin-top:4px; font-family: monospace; font-size:10px;">${preview}</div>
      `;
    });
  } catch (err) {
    infoEl.innerHTML = '<span class="info-label">Could not read selection</span>';
    console.error("Selection error:", err);
  }
}

function generateDataPreview(values, maxRows = 3, maxCols = 4) {
  if (!values || !values.length) return "Empty selection";
  
  const rows = Math.min(values.length, maxRows);
  const cols = Math.min(values[0]?.length || 0, maxCols);
  
  let preview = "";
  for (let r = 0; r < rows; r++) {
    const rowData = [];
    for (let c = 0; c < cols; c++) {
      const val = values[r][c];
      const display = val === null || val === "" ? "·" : String(val).substring(0, 8);
      rowData.push(display);
    }
    if (values[0]?.length > maxCols) rowData.push("...");
    preview += rowData.join(" | ") + "<br>";
  }
  if (values.length > maxRows) preview += "...";
  
  return preview;
}

// ============================================================================
// Quick Actions
// ============================================================================
const QUICK_ACTION_PROMPTS = {
  analyze: "Analyze this data and provide key insights, patterns, and any anomalies you notice. Include statistics like averages, min/max values, and trends.",
  formula: "Based on this data, suggest useful Excel formulas I could use. Provide the exact formula syntax I can copy.",
  chart: "Recommend the best chart type for this data and explain why. Then help me create it.",
  format: "Suggest how to format this data for better readability. Include conditional formatting ideas.",
  clean: "Identify any data quality issues (duplicates, blanks, inconsistencies) and suggest how to clean this data.",
  summarize: "Create a concise summary of this data including totals, averages, and key takeaways."
};

function handleQuickAction(action) {
  const prompt = QUICK_ACTION_PROMPTS[action];
  if (prompt) {
    document.getElementById("promptInput").value = prompt;
    sendMessage();
  }
}

// ============================================================================
// Chat & Messaging
// ============================================================================
function addMessage(role, content, type = "normal") {
  const container = document.getElementById("chatContainer");
  
  // Remove welcome message if present
  const welcome = container.querySelector(".welcome-message");
  if (welcome) welcome.remove();
  
  const msgEl = document.createElement("div");
  msgEl.className = `message ${role} ${type}`;
  
  // Format content (handle code blocks, etc.)
  const formattedContent = formatMessageContent(content);
  msgEl.innerHTML = `<div class="message-content">${formattedContent}</div>`;
  
  container.appendChild(msgEl);
  container.scrollTop = container.scrollHeight;
  
  state.chatHistory.push({ role, content, type });
}

function formatMessageContent(content) {
  // Escape HTML
  let formatted = content
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  
  // Format code blocks
  formatted = formatted.replace(/```(\w*)\n?([\s\S]*?)```/g, (_, lang, code) => {
    return `<div class="code-block">${code.trim()}</div>`;
  });
  
  // Format inline code
  formatted = formatted.replace(/`([^`]+)`/g, '<code style="background:#f0f0f0;padding:2px 4px;border-radius:2px;">$1</code>');
  
  // Format bold
  formatted = formatted.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  
  return formatted;
}

function showTypingIndicator() {
  const container = document.getElementById("chatContainer");
  const indicator = document.createElement("div");
  indicator.className = "message ai typing";
  indicator.id = "typingIndicator";
  indicator.innerHTML = '<div class="typing-indicator"><span></span><span></span><span></span></div>';
  container.appendChild(indicator);
  container.scrollTop = container.scrollHeight;
}

function hideTypingIndicator() {
  document.getElementById("typingIndicator")?.remove();
}

// ============================================================================
// AI Communication
// ============================================================================
async function sendMessage() {
  const input = document.getElementById("promptInput");
  const prompt = input.value.trim();
  
  if (!prompt) return;
  
  // Check API key
  if (!state.apiKey) {
    showSettings();
    setStatus("Please set your Gemini API key first", "error");
    return;
  }
  
  // Add user message
  addMessage("user", prompt);
  input.value = "";
  
  // Show typing indicator
  showTypingIndicator();
  setStatus("Thinking...");
  
  try {
    const response = await callGeminiAPI(prompt);
    hideTypingIndicator();
    
    // Parse and handle the response
    const parsed = parseAIResponse(response);
    state.lastResponse = parsed;
    
    // Add AI message
    addMessage("ai", parsed.message, parsed.hasAction ? "action" : "normal");
    
    // Enable apply button if there's an action
    const applyBtn = document.getElementById("btnApply");
    if (parsed.hasAction && applyBtn) {
      applyBtn.disabled = false;
      state.pendingAction = parsed.action;
    }
    
    setStatus("Ready", "success");
  } catch (err) {
    hideTypingIndicator();
    addMessage("ai", `Error: ${err.message}`, "error");
    setStatus(err.message, "error");
  }
}

async function callGeminiAPI(userPrompt) {
  const systemPrompt = buildSystemPrompt();
  const contextPrompt = buildContextPrompt();
  
  const fullPrompt = `${systemPrompt}\n\n${contextPrompt}\n\nUser Request: ${userPrompt}`;
  
  const url = `${CONFIG.API_ENDPOINT}${CONFIG.GEMINI_MODEL}:generateContent?key=${encodeURIComponent(state.apiKey)}`;
  
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      contents: [{ parts: [{ text: fullPrompt }] }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 2048
      }
    })
  });
  
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API Error ${response.status}: ${errorText}`);
  }
  
  const data = await response.json();
  const aiText = data?.candidates?.[0]?.content?.parts?.map(p => p.text || "").join("") || "";
  
  if (!aiText) throw new Error("Empty response from AI");
  
  return aiText;
}

function buildSystemPrompt() {
  return `You are Excel AI Copilot, an intelligent assistant for Microsoft Excel. You help users analyze data, create formulas, generate charts, format spreadsheets, and automate tasks.

CAPABILITIES:
1. DATA ANALYSIS: Analyze patterns, trends, statistics, anomalies
2. FORMULA CREATION: Generate Excel formulas with exact syntax
3. CHART RECOMMENDATIONS: Suggest and help create visualizations
4. DATA FORMATTING: Conditional formatting, number formats, styles
5. DATA CLEANING: Find duplicates, fix inconsistencies, fill blanks
6. CALCULATIONS: Perform calculations and explain results

RESPONSE FORMAT:
- Be concise and actionable
- When providing formulas, use exact Excel syntax
- When suggesting actions that modify the sheet, include an ACTION block

ACTION BLOCK FORMAT (when you want to modify the sheet):
[ACTION]
type: <formula|value|format|chart>
target: <cell reference or "selection">
data: <the data to apply>
[/ACTION]

Example action for inserting a formula:
[ACTION]
type: formula
target: D2
data: =SUM(A2:C2)
[/ACTION]

Example action for inserting values:
[ACTION]
type: values
target: A1
data: [["Header1","Header2"],["Value1","Value2"]]
[/ACTION]

Only include ACTION blocks when the user wants to modify the sheet. For questions and analysis, just provide the answer.`;
}

function buildContextPrompt() {
  if (!state.selectionData || !state.selectionAddress) {
    return "CONTEXT: No data selected. User should select data for context-aware assistance.";
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
  
  return `CONTEXT:
Selected Range: ${state.selectionAddress}
Dimensions: ${state.selectionData.length} rows × ${state.selectionData[0]?.length || 0} columns

DATA:
${dataStr}`;
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
      message: message || "I'll apply the changes when you click 'Apply Changes'.",
      hasAction: true,
      action: {
        type: typeMatch?.[1]?.trim() || "value",
        target: targetMatch?.[1]?.trim() || "selection",
        data: dataMatch?.[1]?.trim() || ""
      }
    };
  }
  
  return {
    message: response,
    hasAction: false,
    action: null
  };
}


// ============================================================================
// Excel Operations - Apply Changes
// ============================================================================
async function applyChanges() {
  if (!state.pendingAction) {
    setStatus("No pending action to apply", "error");
    return;
  }
  
  const action = state.pendingAction;
  setStatus(`Applying ${action.type}...`);
  
  try {
    await Excel.run(async (context) => {
      switch (action.type) {
        case "formula":
          await applyFormula(context, action);
          break;
        case "values":
        case "value":
          await applyValues(context, action);
          break;
        case "format":
          await applyFormat(context, action);
          break;
        case "chart":
          await createChart(context, action);
          break;
        default:
          await applyValues(context, action);
      }
      
      await context.sync();
    });
    
    addMessage("ai", `✓ Changes applied successfully!`, "action");
    setStatus("Changes applied!", "success");
    
    // Clear pending action
    state.pendingAction = null;
    document.getElementById("btnApply").disabled = true;
    
    // Refresh selection to show updated data
    await refreshSelection();
    
  } catch (err) {
    console.error("Apply error:", err);
    addMessage("ai", `Failed to apply changes: ${err.message}`, "error");
    setStatus("Failed to apply changes", "error");
  }
}

async function applyFormula(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let targetRange;
  
  if (action.target === "selection" || !action.target) {
    targetRange = context.workbook.getSelectedRange();
  } else {
    targetRange = sheet.getRange(action.target);
  }
  
  targetRange.formulas = [[action.data]];
}

async function applyValues(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let targetRange;
  
  if (action.target === "selection" || !action.target) {
    targetRange = context.workbook.getSelectedRange();
  } else {
    targetRange = sheet.getRange(action.target);
  }
  
  // Parse data - could be JSON array or simple value
  let values;
  try {
    values = JSON.parse(action.data);
    if (!Array.isArray(values)) {
      values = [[values]];
    } else if (!Array.isArray(values[0])) {
      values = [values];
    }
  } catch {
    // Simple string value
    values = [[action.data]];
  }
  
  // Resize range if needed
  if (values.length > 1 || values[0]?.length > 1) {
    targetRange = targetRange.getCell(0, 0).getResizedRange(values.length - 1, values[0].length - 1);
  }
  
  targetRange.values = values;
}

async function applyFormat(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let targetRange;
  
  if (action.target === "selection" || !action.target) {
    targetRange = context.workbook.getSelectedRange();
  } else {
    targetRange = sheet.getRange(action.target);
  }
  
  // Parse format instructions
  const formatData = action.data.toLowerCase();
  
  if (formatData.includes("bold")) {
    targetRange.format.font.bold = true;
  }
  if (formatData.includes("italic")) {
    targetRange.format.font.italic = true;
  }
  if (formatData.includes("currency") || formatData.includes("$")) {
    targetRange.numberFormat = [["$#,##0.00"]];
  }
  if (formatData.includes("percent") || formatData.includes("%")) {
    targetRange.numberFormat = [["0.00%"]];
  }
  if (formatData.includes("date")) {
    targetRange.numberFormat = [["mm/dd/yyyy"]];
  }
  if (formatData.includes("center")) {
    targetRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
  }
  if (formatData.includes("header") || formatData.includes("title")) {
    targetRange.format.font.bold = true;
    targetRange.format.fill.color = "#4472C4";
    targetRange.format.font.color = "#FFFFFF";
  }
}

async function createChart(context, action) {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let dataRange;
  
  if (action.target === "selection" || !action.target) {
    dataRange = context.workbook.getSelectedRange();
  } else {
    dataRange = sheet.getRange(action.target);
  }
  
  // Determine chart type from action data
  const chartData = action.data.toLowerCase();
  let chartType = Excel.ChartType.columnClustered;
  
  if (chartData.includes("line")) {
    chartType = Excel.ChartType.line;
  } else if (chartData.includes("pie")) {
    chartType = Excel.ChartType.pie;
  } else if (chartData.includes("bar")) {
    chartType = Excel.ChartType.barClustered;
  } else if (chartData.includes("area")) {
    chartType = Excel.ChartType.area;
  } else if (chartData.includes("scatter")) {
    chartType = Excel.ChartType.xyscatter;
  }
  
  const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
  chart.setPosition("H2", "O15");
  chart.title.text = "Generated Chart";
}

// ============================================================================
// Additional Excel Operations
// ============================================================================
async function getSheetData(rangeName) {
  let data = null;
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = rangeName ? sheet.getRange(rangeName) : context.workbook.getSelectedRange();
    range.load("values");
    await context.sync();
    data = range.values;
  });
  return data;
}

async function setSheetData(rangeName, values) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(rangeName);
    range.values = values;
    await context.sync();
  });
}

async function insertFormula(cell, formula) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(cell);
    range.formulas = [[formula]];
    await context.sync();
  });
}

async function getWorksheetNames() {
  const names = [];
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    sheets.items.forEach(s => names.push(s.name));
  });
  return names;
}

async function createNewSheet(name) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add(name);
    sheet.activate();
    await context.sync();
  });
}

async function sortSelection(ascending = true) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("columnIndex");
    await context.sync();
    
    const sortFields = [{
      key: 0,
      ascending: ascending
    }];
    
    range.sort.apply(sortFields);
    await context.sync();
  });
}

async function autoFitColumns() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.autofitColumns();
    await context.sync();
  });
}

// ============================================================================
// Settings Management
// ============================================================================
function showSettings() {
  const modal = document.getElementById("settingsModal");
  const input = document.getElementById("apiKeyInput");
  if (modal && input) {
    input.value = state.apiKey;
    modal.classList.remove("hidden");
  }
}

function closeSettings() {
  document.getElementById("settingsModal")?.classList.add("hidden");
}

function saveSettings() {
  const input = document.getElementById("apiKeyInput");
  if (input) {
    state.apiKey = input.value.trim();
    localStorage.setItem(CONFIG.STORAGE_KEY, state.apiKey);
    closeSettings();
    setStatus("API key saved!", "success");
  }
}

// ============================================================================
// Status Management
// ============================================================================
function setStatus(message, type = "") {
  const statusBar = document.getElementById("statusBar");
  if (statusBar) {
    statusBar.textContent = message;
    statusBar.className = "status-bar " + type;
  }
}

// ============================================================================
// Export for testing
// ============================================================================
export {
  sendMessage,
  applyChanges,
  refreshSelection,
  getSheetData,
  setSheetData,
  insertFormula,
  sortSelection,
  autoFitColumns
};
