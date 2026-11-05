const MASTER_BUILD = "1.1";
const GOOGLE_CLIENT_ID = "154634144934-9hg9o4ra7uriu5hrivaaj73mduj7udf4.apps.googleusercontent.com";
const GOOGLE_API_KEY = "AIzaSyCh8ia27PwiWJkPCypoUyvj5TD8YJVjJSc";
const SHEETS_DISCOVERY_DOC = "https://sheets.googleapis.com/$discovery/rest?version=v4";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly";

const $ = (id) => document.getElementById(id);
const state = {
  tokenClient: null,
  gapiReady: false,
  gisReady: false,
  accessToken: null,
  spreadsheetId: "",
  spreadsheetTitle: "",
  tabs: [],
  tabProperties: null,
  headers: [],
  values: [],
  sourceRows: [],
  folderHandle: null,
  folderFiles: [],
  reviewRows: [],
};

function setStatus(message, kind = "") {
  $("loggerStatus").textContent = message;
  $("connectionDot").className = ("connection-dot " + kind).trim();
}

function getErrorMessage(error) {
  return error?.result?.error?.message || error?.body?.error?.message || error?.message ||
    (typeof error === "string" ? error : "Google returned an unspecified error.");
}

function parseSheetUrl(value) {
  const match = String(value || "").trim().match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) throw new Error("Please enter a valid Google Sheets URL.");
  return match[1];
}

function quoteSheetName(value) {
  return "'" + String(value).replaceAll("'", "''") + "'";
}

function columnLetter(number) {
  let result = "";
  let value = Math.max(1, Number(number) || 1);
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;",
  }[character]));
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function initializeGoogle() {
  if (!window.gapi || !window.google?.accounts?.oauth2) {
    setTimeout(initializeGoogle, 250);
    return;
  }
  gapi.load("client", async () => {
    try {
      await gapi.client.init({ apiKey: GOOGLE_API_KEY, discoveryDocs: [SHEETS_DISCOVERY_DOC] });
      state.gapiReady = true;
      enableConnect();
    } catch (error) {
      setStatus("Google API setup error: " + getErrorMessage(error), "error");
    }
  });
  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: SHEETS_SCOPE,
    callback: "",
  });
  state.gisReady = true;
  enableConnect();
}

function enableConnect() {
  $("connectButton").disabled = !(state.gapiReady && state.gisReady);
}

function requestGoogleAccess() {
  state.tokenClient.callback = async (response) => {
    if (response.error) {
      setStatus("Google authorization failed: " + response.error, "error");
      return;
    }
    state.accessToken = response.access_token;
    gapi.client.setToken({ access_token: state.accessToken });
    $("disconnectButton").disabled = false;
    await loadWorkbook();
  };
  state.tokenClient.requestAccessToken({ prompt: gapi.client.getToken() ? "" : "consent" });
}

async function loadWorkbook() {
  try {
    state.spreadsheetId = parseSheetUrl($("sheetUrl").value);
    setStatus("Reading workbook tabs…", "working");
    const response = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: state.spreadsheetId,
      includeGridData: false,
      fields: "properties(title),sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
    });
    state.spreadsheetTitle = response.result.properties.title;
    state.tabs = response.result.sheets.map((sheet) => sheet.properties);
    $("sheetTab").innerHTML = state.tabs.map((tab) =>
      '<option value="' + escapeHtml(tab.title) + '">' + escapeHtml(tab.title) + "</option>").join("");
    $("sheetTab").disabled = false;
    $("setupCard").classList.remove("hidden");
    $("folderCard").classList.remove("hidden");
    $("sheetSummary").textContent = state.spreadsheetTitle + " connected · " + state.tabs.length + " tab" + (state.tabs.length === 1 ? "" : "s") + " found.";
    await loadSelectedTab();
    setStatus("Connected", "connected");
  } catch (error) {
    setStatus("Could not read this Sheet: " + getErrorMessage(error), "error");
  }
}

function defaultColumn(headers, patterns, fallback) {
  const found = headers.findIndex((header) => patterns.some((pattern) => pattern.test(String(header))));
  return found >= 0 ? String(found) : String(Math.min(fallback, Math.max(0, headers.length - 1)));
}

function populateColumnSelect(id, headers, selectedIndex) {
  $(id).innerHTML = headers.map((header, index) =>
    '<option value="' + index + '"' + (String(index) === String(selectedIndex) ? " selected" : "") + ">" +
    columnLetter(index + 1) + " — " + escapeHtml(header || "Column " + columnLetter(index + 1)) + "</option>").join("");
  $(id).disabled = !headers.length;
}

async function loadSelectedTab() {
  try {
    const title = $("sheetTab").value;
    state.tabProperties = state.tabs.find((tab) => tab.title === title) || state.tabs[0];
    const columns = Math.min(Math.max(state.tabProperties?.gridProperties?.columnCount || 26, 14), 52);
    const endColumn = columnLetter(columns);
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: state.spreadsheetId,
      range: quoteSheetName(title) + "!A1:" + endColumn + "10000",
      valueRenderOption: "FORMATTED_VALUE",
    });
    const rows = response.result.values || [];
    state.headers = rows[0] || [];
    state.values = rows.slice(1).map((row, index) => ({ rowNumber: index + 2, values: row }));
    populateColumnSelect("archivalColumn", state.headers, defaultColumn(state.headers, [/archiv/i, /file\s*name/i, /arc/i], 1));
    populateColumnSelect("originalColumn", state.headers, defaultColumn(state.headers, [/original/i, /source\s*file/i], 13));
    populateColumnSelect("sourceColumn", state.headers, defaultColumn(state.headers, [/source/i, /vendor/i], 2));
    populateSourceOptions();
    $("setupStatus").textContent = "Loaded " + state.values.length + " data row" + (state.values.length === 1 ? "" : "s") + " from " + title + ". Confirm the three columns before choosing a source.";
  } catch (error) {
    $("setupStatus").textContent = "Could not read this tab: " + getErrorMessage(error);
  }
}

function populateSourceOptions() {
  const sourceIndex = Number($("sourceColumn").value);
  const sources = [...new Set(state.values.map((row) => String(row.values[sourceIndex] || "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));
  $("sourceSelect").innerHTML = '<option value="">Choose a source</option>' + sources.map((source) =>
    '<option value="' + escapeHtml(source) + '">' + escapeHtml(source) + "</option>").join("");
  $("sourceSelect").disabled = !sources.length;
  if (!sources.length) $("setupStatus").textContent = "No source values were found in the selected Source column.";
  renderReviewIfReady();
}

async function collectRootFiles(directoryHandle) {
  const files = [];
  for await (const entry of directoryHandle.values()) {
    if (entry.kind === "file") files.push({ name: entry.name, handle: entry, parentHandle: directoryHandle });
  }
  files.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" }));
  return files;
}

function normalizedOriginal(value) {
  return String(value || "").trim().split(/[\\/]/).pop().replace(/\.[A-Za-z0-9]{2,5}$/, "").trim();
}

function fileMatchesOriginal(original, fileName) {
  const key = normalizedOriginal(original);
  if (!key) return false;
  const stem = String(fileName).replace(/\.[^.]+$/, "");
  const expression = new RegExp("(^|[^A-Za-z0-9])" + escapeRegExp(key) + "(?=$|[^A-Za-z0-9])", "i");
  return expression.test(stem);
}

function addHiresToName(value, originalFileName) {
  let name = String(value || "").trim();
  if (!name) return "";
  if (/(^|_)HIRES(?=_|$)/i.test(name)) return ensureExtension(name, originalFileName);
  if (/ARC\d+(?=_|$)/i.test(name)) name = name.replace(/(ARC\d+)(?=_|$)/i, "$1_HIRES");
  else name = "HIRES_" + name;
  return ensureExtension(name, originalFileName);
}

function ensureExtension(name, originalFileName) {
  if (/\.[A-Za-z0-9]{2,5}$/.test(name)) return name;
  return name + (String(originalFileName).match(/\.[^.]+$/)?.[0] || "");
}

function statusClass(status) {
  if (/ready|success/i.test(status)) return "status-found";
  if (/not found|multiple|missing|exists|failed/i.test(status)) return "status-error";
  return "status-warning";
}

function buildReviewRows() {
  const source = $("sourceSelect").value;
  const archivalIndex = Number($("archivalColumn").value);
  const originalIndex = Number($("originalColumn").value);
  const sourceIndex = Number($("sourceColumn").value);
  if (!source || !state.folderFiles.length || !state.headers.length) return [];
  const assigned = new Set();
  return state.values
    .filter((row) => String(row.values[sourceIndex] || "").trim().toLowerCase() === source.toLowerCase())
    .map((row) => {
      const archivalName = String(row.values[archivalIndex] || "").trim();
      const originalName = String(row.values[originalIndex] || "").trim();
      const matches = state.folderFiles.filter((file) => !assigned.has(file) && fileMatchesOriginal(originalName, file.name));
      const match = matches.length === 1 ? matches[0] : null;
      if (match) assigned.add(match);
      let status = "Ready to rename";
      if (!archivalName) status = "Missing archival name";
      else if (!originalName) status = "Missing original filename";
      else if (matches.length > 1) status = "Multiple folder matches";
      else if (!match) status = "File not found";
      return { rowNumber: row.rowNumber, archivalName, originalName, source, file: match, status, newName: match && archivalName ? addHiresToName(archivalName, match.name) : "" };
    });
}

function renderReviewIfReady() {
  if (!state.folderFiles.length || !$("sourceSelect").value) {
    $("reviewPanel").classList.add("hidden");
    $("renameButton").disabled = true;
    return;
  }
  state.reviewRows = buildReviewRows();
  const foundCount = state.reviewRows.filter((row) => row.status === "Ready to rename").length;
  $("reviewPanel").innerHTML = '<table class="master-review-table"><thead><tr>' +
    '<th>Sheet Row</th><th>Archival File Name</th><th>Source</th><th>Original File Name</th><th>Folder File Found</th><th>Status</th><th>New HIRES Name</th></tr></thead><tbody>' +
    state.reviewRows.map((row) => '<tr><td>' + row.rowNumber + '</td><td>' + escapeHtml(row.archivalName || "—") + '</td><td>' + escapeHtml(row.source) + '</td><td>' + escapeHtml(row.originalName || "—") + '</td><td>' + escapeHtml(row.file?.name || "—") + '</td><td class="' + statusClass(row.status) + '">' + escapeHtml(row.status) + '</td><td>' + escapeHtml(row.newName || "—") + '</td></tr>').join("") +
    '</tbody></table>';
  $("reviewPanel").classList.remove("hidden");
  $("renameButton").disabled = foundCount === 0;
  $("reviewNote").textContent = foundCount + " of " + state.reviewRows.length + " row" + (state.reviewRows.length === 1 ? "" : "s") + " ready to rename. Any unknown or ambiguous row will be left unchanged.";
  const matchedFiles = new Set(state.reviewRows.filter((row) => row.file).map((row) => row.file));
  const unmatched = state.folderFiles.filter((file) => !matchedFiles.has(file));
  if (unmatched.length) {
    $("unmatchedFiles").innerHTML = "Folder files not matched to a selected Sheet row: <strong>" + unmatched.map((file) => escapeHtml(file.name)).join(", ") + "</strong>";
    $("unmatchedFiles").classList.remove("hidden");
  } else {
    $("unmatchedFiles").classList.add("hidden");
  }
}

async function chooseFolder() {
  if (!window.showDirectoryPicker) {
    $("folderStatus").textContent = "Folder access requires Chrome or Edge.";
    return;
  }
  try {
    const handle = await window.showDirectoryPicker({ mode: "readwrite" });
    state.folderHandle = handle;
    state.folderFiles = await collectRootFiles(handle);
    $("folderStatus").textContent = handle.name + " selected · " + state.folderFiles.length + " file" + (state.folderFiles.length === 1 ? "" : "s") + " found in the root folder.";
    renderReviewIfReady();
  } catch (error) {
    if (error.name !== "AbortError") $("folderStatus").textContent = "Could not read folder: " + getErrorMessage(error);
  }
}

async function renameConfirmedFiles() {
  const rows = state.reviewRows.filter((row) => row.status === "Ready to rename" && row.file);
  if (!rows.length) return;
  if (!window.confirm("Rename " + rows.length + " confirmed file" + (rows.length === 1 ? "" : "s") + " with HIRES names?")) return;
  $("renameButton").disabled = true;
  const results = [];
  for (const row of rows) {
    try {
      const destination = row.newName;
      try {
        await state.folderHandle.getFileHandle(destination);
        row.status = "Destination already exists";
        results.push("Sheet row " + row.rowNumber + " — " + destination + " — Not renamed: destination already exists");
        continue;
      } catch (error) {
        // No destination exists; continue.
      }
      const originalFile = await row.file.handle.getFile();
      const newFileHandle = await state.folderHandle.getFileHandle(destination, { create: true });
      const writable = await newFileHandle.createWritable();
      await writable.write(originalFile);
      await writable.close();
      await state.folderHandle.removeEntry(row.file.name);
      row.status = "Successfully renamed";
      row.file = { ...row.file, name: destination };
      results.push("Sheet row " + row.rowNumber + " — renamed to " + destination + " — Successfully renamed");
    } catch (error) {
      row.status = "Rename failed";
      results.push("Sheet row " + row.rowNumber + " — " + row.file.name + " — Not renamed: " + getErrorMessage(error));
    }
  }
  $("renameStatus").textContent = results.join("\n");
  renderReviewIfReady();
  $("renameButton").disabled = true;
}

function disconnect() {
  const token = gapi.client.getToken();
  if (token) google.accounts.oauth2.revoke(token.access_token);
  gapi.client.setToken("");
  state.accessToken = null;
  state.spreadsheetId = "";
  $("disconnectButton").disabled = true;
  $("setupCard").classList.add("hidden");
  $("folderCard").classList.add("hidden");
  setStatus("Disconnected");
}

window.addEventListener("load", () => {
  $("loggerStatus").textContent = "Build " + MASTER_BUILD + " loaded. Enter a Sheet link, then connect your Google account.";
  $("connectButton").addEventListener("click", () => {
    try {
      parseSheetUrl($("sheetUrl").value);
      requestGoogleAccess();
    } catch (error) {
      setStatus(getErrorMessage(error), "error");
    }
  });
  $("disconnectButton").addEventListener("click", disconnect);
  $("sheetTab").addEventListener("change", loadSelectedTab);
  $("sourceColumn").addEventListener("change", populateSourceOptions);
  $("sourceSelect").addEventListener("change", renderReviewIfReady);
  $("archivalColumn").addEventListener("change", renderReviewIfReady);
  $("originalColumn").addEventListener("change", renderReviewIfReady);
  $("chooseFolderButton").addEventListener("click", chooseFolder);
  $("resetFolderButton").addEventListener("click", () => {
    state.folderHandle = null;
    state.folderFiles = [];
    state.reviewRows = [];
    $("folderStatus").textContent = "No archive folder selected.";
    $("reviewPanel").classList.add("hidden");
    $("unmatchedFiles").classList.add("hidden");
    $("renameButton").disabled = true;
  });
  $("renameButton").addEventListener("click", renameConfirmedFiles);
  initializeGoogle();
});
