/*
  Google Sheets connection for The Lovely Runner.

  Before using this page, add a Google OAuth web client ID and API key below.
  These are browser-safe identifiers. Never put a Google client secret here.
*/
const GOOGLE_CLIENT_ID = "PASTE_GOOGLE_OAUTH_CLIENT_ID_HERE";
const GOOGLE_API_KEY = "PASTE_GOOGLE_API_KEY_HERE";
const SHEETS_DISCOVERY_DOC = "https://sheets.googleapis.com/$discovery/rest?version=v4";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

const $ = (id) => document.getElementById(id);
const state = {
  spreadsheetId: "",
  spreadsheetTitle: "",
  tabs: [],
  tokenClient: null,
  gapiReady: false,
  gisReady: false,
  accessToken: null,
};

function setStatus(message, kind = "") {
  $("loggerStatus").textContent = message;
  $("connectionDot").className = ("connection-dot " + kind).trim();
}

function setWriteStatus(message, kind = "") {
  $("writeStatus").textContent = message;
  $("writeStatus").className = ("logger-status " + kind).trim();
}

function parseSheetUrl(value) {
  const url = String(value || "").trim();
  const match = url.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) throw new Error("Please enter a valid Google Sheets URL.");
  return match[1];
}

function ensureCredentials() {
  if (GOOGLE_CLIENT_ID.includes("PASTE_") || GOOGLE_API_KEY.includes("PASTE_")) {
    throw new Error("Google credentials have not been added to archival-logging.js yet.");
  }
}

function initializeGoogle() {
  if (!window.gapi || !window.google?.accounts?.oauth2) return;

  gapi.load("client", async () => {
    try {
      await gapi.client.init({
        apiKey: GOOGLE_API_KEY,
        discoveryDocs: [SHEETS_DISCOVERY_DOC],
      });
      state.gapiReady = true;
      maybeEnableConnection();
    } catch (error) {
      setStatus("Google API setup error: " + error.message, "error");
    }
  });

  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: SHEETS_SCOPE,
    callback: "",
  });
  state.gisReady = true;
  maybeEnableConnection();
}

function maybeEnableConnection() {
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
    $("signOutButton").disabled = false;
    await loadTracker();
  };

  state.tokenClient.requestAccessToken({
    prompt: gapi.client.getToken() ? "" : "consent",
  });
}

async function loadTracker() {
  try {
    state.spreadsheetId = parseSheetUrl($("sheetUrl").value);
    setStatus("Reading tracker information…", "working");

    const response = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: state.spreadsheetId,
      includeGridData: false,
      fields: "properties(title),sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
    });

    state.spreadsheetTitle = response.result.properties.title;
    state.tabs = response.result.sheets.map((sheet) => sheet.properties);
    $("trackerSummary").textContent = state.spreadsheetTitle + " is connected with read/write authorization.";
    $("trackerDetails").innerHTML =
      '<div><span>Spreadsheet</span><strong>' + escapeHtml(state.spreadsheetTitle) + '</strong></div>' +
      '<div><span>Tabs found</span><strong>' + state.tabs.length + '</strong></div>';

    $("sheetTab").innerHTML = state.tabs
      .map((tab) => '<option value="' + escapeAttribute(tab.title) + '">' + escapeHtml(tab.title) + "</option>")
      .join("");

    $("trackerCard").classList.remove("hidden");
    $("writeTestButton").disabled = false;
    setStatus("Connected", "connected");
  } catch (error) {
    setStatus("Could not read this sheet: " + error.message, "error");
    $("trackerCard").classList.add("hidden");
    $("writeTestButton").disabled = true;
  }
}

async function writeTestValue() {
  try {
    const tab = $("sheetTab").value;
    const cell = $("testCell").value.trim();
    const value = $("testValue").value;
    if (!tab || !cell) throw new Error("Choose a tab and enter a cell.");

    setWriteStatus("Writing test value…", "working");
    await gapi.client.sheets.spreadsheets.values.update({
      spreadsheetId: state.spreadsheetId,
      range: quoteSheetName(tab) + "!" + cell,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[value]] },
    });
    setWriteStatus("Write successful: " + tab + "!" + cell, "success");
  } catch (error) {
    setWriteStatus("Write failed: " + error.message, "error");
  }
}

function disconnect() {
  const token = gapi.client.getToken();
  if (token) google.accounts.oauth2.revoke(token.access_token);
  gapi.client.setToken("");
  state.accessToken = null;
  state.spreadsheetId = "";
  $("trackerCard").classList.add("hidden");
  $("signOutButton").disabled = true;
  setStatus("Disconnected");
  setWriteStatus("");
}

function quoteSheetName(name) {
  return "'" + String(name).replaceAll("'", "''") + "'";
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;",
  }[char]));
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

window.addEventListener("load", () => {
  $("connectButton").addEventListener("click", () => {
    try {
      ensureCredentials();
      parseSheetUrl($("sheetUrl").value);
      requestGoogleAccess();
    } catch (error) {
      setStatus(error.message, "error");
    }
  });
  $("writeTestButton").addEventListener("click", writeTestValue);
  $("signOutButton").addEventListener("click", disconnect);
  initializeGoogle();
});
