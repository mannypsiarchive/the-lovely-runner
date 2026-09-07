/*
  Google Sheets connection for The Lovely Runner.

  Before using this page, add a Google OAuth web client ID and API key below.
  These are browser-safe identifiers. Never put a Google client secret here.
*/
const GOOGLE_CLIENT_ID = "154634144934-9hg9o4ra7uriu5hrivaaj73mduj7udf4.apps.googleusercontent.com";
const GOOGLE_API_KEY = "AIzaSyCh8ia27PwiWJkPCypoUyvj5TD8YJVjJSc";
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
  clipFiles: [],
  clipRows: [],
  sourceDirectoryHandle: null,
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
  if (state.tokenClient || state.gapiReady) return;
  if (!window.gapi || !window.google?.accounts?.oauth2) {
    setTimeout(initializeGoogle, 250);
    return;
  }

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
  if (!state.tokenClient) {
    setStatus("Google sign-in is still loading. Please try again in a moment.", "working");
    return;
  }

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
    const trackerDetails = $("trackerDetails");
    if (trackerDetails) {
      trackerDetails.innerHTML =
        '<div><span>Spreadsheet</span><strong>' + escapeHtml(state.spreadsheetTitle) + '</strong></div>' +
        '<div><span>Tabs found</span><strong>' + state.tabs.length + '</strong></div>';
    }

    $("trackerCard").classList.remove("hidden");
    $("clipLogCard").classList.remove("hidden");
    if (state.clipFiles.length) updateClipPreview();
    setStatus("Connected", "connected");
  } catch (error) {
    setStatus("Could not read this sheet: " + error.message, "error");
    $("trackerCard").classList.add("hidden");
    $("clipLogCard").classList.add("hidden");
  }
}

function classifyClip(fileName) {
  const extension = String(fileName).split(".").pop().toLowerCase();
  const videoExtensions = ["mov", "mp4", "mxf", "m4v", "avi", "mts", "m2ts", "wmv", "webm"];
  const imageExtensions = ["jpg", "jpeg", "png", "tif", "tiff", "webp", "heic", "gif", "bmp"];
  if (videoExtensions.includes(extension)) return "F";
  if (imageExtensions.includes(extension)) return "S";
  return "";
}

async function collectSourceFiles(directoryHandle) {
  const files = [];
  async function walk(handle) {
    for await (const entry of handle.values()) {
      if (entry.kind === "file") {
        files.push({ name: entry.name, handle: entry, parentHandle: handle });
      } else if (entry.kind === "directory") {
        await walk(entry);
      }
    }
  }
  await walk(directoryHandle);
  return files;
}

function sourceFileName(fileName) {
  const stem = String(fileName).replace(/\.[^.]+$/, "");
  const gettyMatch = stem.match(/^gettyimages-(.+?)-\d+_adpp$/i);
  if (gettyMatch) return gettyMatch[1];
  return stem;
}

function getClipRows() {
  const startValue = $("arcStart").value.trim();
  const start = Number(startValue);
  return state.clipFiles.filter((file) => classifyClip(file.name)).map((file, index) => {
    const type = classifyClip(file.name);
    return {
      file,
      row: start + index + 1,
      sourceName: sourceFileName(file.name),
      type,
    };
  });
}

async function readSourceLinks(rows) {
  if (!state.spreadsheetId || !rows.length) return rows.map((row) => ({ ...row, sourceLink: "" }));
  const response = await gapi.client.sheets.spreadsheets.values.batchGet({
    spreadsheetId: state.spreadsheetId,
    ranges: rows.map((item) => quoteSheetName("TAPE LOG") + "!O" + item.row),
  });
  const valueRanges = response.result.valueRanges || [];
  return rows.map((row, index) => ({
    ...row,
    sourceLink: valueRanges[index]?.values?.[0]?.[0] || "",
  }));
}

function vendorFromSource(sourceLink, fileName) {
  if (/gettyimages|Getty Images/i.test(sourceLink)) return "Getty Images";
  if (/gettyimages/i.test(fileName)) return "Getty Images";
  return "";
}

function sourceUrl(sourceLink) {
  const raw = String(sourceLink || "");
  // Capture the URL cleanly from plain text, Markdown links, or a
  // HYPERLINK-style cell value without including brackets or punctuation.
  const match = raw.match(/https?:\/\/[^\s\]\)"']+/i);
  return match ? match[0] : "";
}

function assetIdFromSourceLink(sourceLink) {
  const rawUrl = sourceUrl(sourceLink);
  if (!rawUrl) return "";
  try {
    const url = new URL(rawUrl);
    const parts = url.pathname.split("/").filter(Boolean);
    const lastPart = parts[parts.length - 1] || "";
    return /^(?:\d+-?)+$/.test(lastPart) ? lastPart : "";
  } catch (error) {
    return "";
  }
}

function descriptionFromSourceLink(sourceLink) {
  if (!sourceLink) return "";
  try {
    const url = new URL(sourceUrl(sourceLink));
    // Getty uses /detail/video/<description>/<asset-id> URLs.
    // Keep support for the alternate detail-video form as well.
    const match = url.pathname.match(/\/detail(?:\/[^/]+|-video)\/([^/]+)\/[^/]+\/?$/i);
    if (!match) return "";
    const words = decodeURIComponent(match[1]).replace(/[-_]+/g, " ").trim();
    return words ? words.charAt(0).toUpperCase() + words.slice(1) : "";
  } catch (error) {
    return "";
  }
}

async function updateClipPreview() {
  const startValue = $("arcStart").value.trim();
  const start = Number(startValue);
  const preview = $("clipPreview");
  const files = state.clipFiles;

  if (!files.length) {
    $("clipSelectionStatus").textContent = "Choose a source folder containing one or more clips.";
    preview.classList.add("hidden");
    $("logClipButton").disabled = true;
    return;
  }
  if (!/^\d+$/.test(startValue) || start < 0) {
    $("clipSelectionStatus").textContent = "Enter the ARC number to start at.";
    preview.classList.add("hidden");
    $("logClipButton").disabled = true;
    return;
  }

  const allRows = files.map((file, index) => ({
    file,
    row: start + index + 1,
    sourceName: sourceFileName(file.name),
    type: classifyClip(file.name),
  }));
  const unsupported = allRows.filter((item) => !item.type).length;
  const rows = allRows.filter((item) => item.type);
  state.clipFiles = rows.map((item) => item.file);
  const adjustedRows = (await readSourceLinks(rows)).map((item) => ({
    ...item,
    sourceName: assetIdFromSourceLink(item.sourceLink) || item.sourceName,
  }));
  state.clipRows = adjustedRows;

  $("clipSelectionStatus").textContent =
    adjustedRows.length + " file" + (adjustedRows.length === 1 ? "" : "s") + " ready. " +
    (unsupported ? unsupported + " unsupported file" + (unsupported === 1 ? " was" : "s were") + " ignored." : "");
  preview.innerHTML = adjustedRows.map((item) => {
    const vendor = vendorFromSource(item.sourceLink, item.file.name);
    const description = descriptionFromSourceLink(item.sourceLink);
    return "<div><strong>Row " + item.row + "</strong> · Source File Name (D): <strong>" + escapeHtml(item.sourceName) +
    "</strong> · Still/Footage (M): <strong>" + escapeHtml(item.type) +
    "</strong> · Vendor/Source (F): <strong>" + escapeHtml(vendor || "—") + "</strong> · Description (J): <strong>" +
    escapeHtml(description || "—") + "</strong> · Column O link: <strong>" + (item.sourceLink ? "found" : "not found") +
    "</strong> <span>(original file: " + escapeHtml(item.file.name) + ")</span></div>";
  }).join("");
  preview.classList.toggle("hidden", !adjustedRows.length);
  $("logClipButton").disabled = !adjustedRows.length;
}

async function logTestClips() {
  const startValue = $("arcStart").value.trim();
  if (!/^\d+$/.test(startValue) || Number(startValue) < 0 || !state.clipFiles.length) {
    $("clipWriteStatus").textContent = "Choose a supported clip folder and enter an ARC number.";
    return;
  }

  try {
    // Re-read Column O immediately before writing so a link added after the
    // preview is still used for Vendor/Source and Description.
    const rows = (await readSourceLinks(getClipRows())).map((item) => ({
      ...item,
      sourceName: assetIdFromSourceLink(item.sourceLink) || item.sourceName,
    }));
    state.clipRows = rows;
    const data = rows.map((item) => {
      const vendor = vendorFromSource(item.sourceLink, item.file.name);
      const description = descriptionFromSourceLink(item.sourceLink);
      return [
        { range: quoteSheetName("TAPE LOG") + "!D" + item.row, values: [[item.sourceName]] },
        { range: quoteSheetName("TAPE LOG") + "!M" + item.row, values: [[item.type]] },
        ...(vendor ? [{ range: quoteSheetName("TAPE LOG") + "!F" + item.row, values: [[vendor]] }] : []),
        ...(description ? [{ range: quoteSheetName("TAPE LOG") + "!J" + item.row, values: [[description]] }] : []),
      ];
    }).flat();
    $("logClipButton").disabled = true;
    $("clipWriteStatus").textContent = "Writing " + state.clipFiles.length + " test row" + (state.clipFiles.length === 1 ? "" : "s") + " to TAPE LOG…";
    await gapi.client.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: state.spreadsheetId,
      resource: {
        valueInputOption: "USER_ENTERED",
        data,
      },
    });
    $("clipWriteStatus").textContent = "Success. Wrote source filename(s), footage/still type, vendor, and link description to TAPE LOG.";
    $("renamePanel").classList.remove("hidden");
    $("renameButton").disabled = false;
  } catch (error) {
    $("clipWriteStatus").textContent = "Write failed: " + error.message;
    $("logClipButton").disabled = false;
  }
}

async function readFinalNames(rows) {
  const response = await gapi.client.sheets.spreadsheets.values.batchGet({
    spreadsheetId: state.spreadsheetId,
    ranges: rows.map((item) => quoteSheetName("TAPE LOG") + "!B" + item.row),
  });
  const valueRanges = response.result.valueRanges || [];
  return rows.map((row, index) => ({
    ...row,
    finalName: valueRanges[index]?.values?.[0]?.[0]?.trim() || "",
  }));
}

function sourceIdMatchesFile(sourceId, fileName) {
  const escaped = String(sourceId).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp("(^|[^A-Za-z0-9])" + escaped + "(?=$|[^A-Za-z0-9])", "i").test(fileName);
}

function finalFileName(value, originalName) {
  const desired = String(value || "").trim();
  if (!desired) return "";
  if (/\.[A-Za-z0-9]{2,5}$/.test(desired)) return desired;
  const extension = String(originalName).match(/\.[^.]+$/)?.[0] || "";
  return desired + extension;
}

async function renameLoggedFiles() {
  if (!state.sourceDirectoryHandle) {
    $("renameStatus").textContent = "Choose the source folder using the folder button before renaming.";
    return;
  }
  if (!window.confirm("Rename the matching files in the selected source folder using the names in Column B?")) {
    return;
  }
  try {
    $("renameButton").disabled = true;
    $("renameStatus").textContent = "Reading final filenames from Column B…";
    const rows = await readFinalNames(state.clipRows?.length ? state.clipRows : getClipRows());
    const results = [];

    for (const row of rows) {
      if (!row.sourceName) {
        results.push("Row " + row.row + ": missing source filename in D");
        continue;
      }
      if (!row.finalName) {
        results.push("Row " + row.row + ": missing final filename in B");
        continue;
      }
      const matches = state.clipFiles.filter((file) => sourceIdMatchesFile(row.sourceName, file.name));
      if (matches.length !== 1) {
        results.push("Row " + row.row + ": found " + matches.length + " matching source files for " + row.sourceName);
        continue;
      }

      const sourceFile = matches[0];
      const newName = finalFileName(row.finalName, sourceFile.name);
      if (!newName || newName === sourceFile.name) {
        results.push("Row " + row.row + ": already has the requested filename");
        continue;
      }
      try {
        await sourceFile.parentHandle.getFileHandle(newName);
        results.push("Row " + row.row + ": skipped because " + newName + " already exists");
        continue;
      } catch (error) {
        // The destination does not exist, so it is safe to create it below.
      }

      const originalFile = await sourceFile.handle.getFile();
      const newFileHandle = await sourceFile.parentHandle.getFileHandle(newName, { create: true });
      const writable = await newFileHandle.createWritable();
      await writable.write(originalFile);
      await writable.close();
      await sourceFile.parentHandle.removeEntry(sourceFile.name);
      results.push("Row " + row.row + ": renamed to " + newName);
    }

    $("renameStatus").textContent = results.join(" · ");
  } catch (error) {
    $("renameStatus").textContent = "Rename failed: " + error.message;
    $("renameButton").disabled = false;
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
  $("signOutButton").addEventListener("click", disconnect);
  $("chooseSourceButton").addEventListener("click", async () => {
    if (!window.showDirectoryPicker) {
      $("sourceFolderName").textContent = "Direct renaming requires Chrome or Edge.";
      return;
    }
    try {
      const directoryHandle = await window.showDirectoryPicker({ mode: "readwrite" });
      state.sourceDirectoryHandle = directoryHandle;
      state.clipFiles = await collectSourceFiles(directoryHandle);
      $("sourceFolderName").textContent = directoryHandle.name + " selected · " + state.clipFiles.length + " file" + (state.clipFiles.length === 1 ? "" : "s") + " found.";
      await updateClipPreview();
    } catch (error) {
      if (error.name !== "AbortError") $("sourceFolderName").textContent = "Could not read source folder: " + error.message;
    }
  });
  $("arcStart").value = "";
  $("arcStart").addEventListener("input", updateClipPreview);
  $("logClipButton").addEventListener("click", logTestClips);
  $("renameButton").addEventListener("click", renameLoggedFiles);
  initializeGoogle();
});
