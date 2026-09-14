/*
  Fair Use Log 2.0

  The media workflow is intentionally local:
  - The workbook and reference cut stay in the user's browser.
  - FFmpeg.wasm reads the media and extracts frames locally.
  - The reference cut's embedded start timecode and frame rate are detected.
  - ExcelJS embeds the generated JPEGs into column A.

  Google Sheet input uses the same logger-style OAuth connection. The sheet is
  read into the same internal row model, so the timecode logic is not different
  for Excel and Google input.
*/

const FAIR_USE_LOG_VERSION = "2.0";
const FFMPEG_VERSION = "0.12.10";
const FFMPEG_PACKAGE_BASE = `https://cdn.jsdelivr.net/npm/@ffmpeg/ffmpeg@${FFMPEG_VERSION}/dist/umd`;
const FFMPEG_CORE_BASE = `https://cdn.jsdelivr.net/npm/@ffmpeg/core@${FFMPEG_VERSION}/dist/umd`;
const FFMPEG_CLASS_URL = `${FFMPEG_PACKAGE_BASE}/ffmpeg.js?fair-use-log=${FAIR_USE_LOG_VERSION}`;
const FFMPEG_CLASS_WORKER_URL = `${FFMPEG_PACKAGE_BASE}/814.ffmpeg.js?fair-use-log=${FAIR_USE_LOG_VERSION}`;

const GOOGLE_CLIENT_ID = "154634144934-9hg9o4ra7uriu5hrivaaj73mduj7udf4.apps.googleusercontent.com";
const GOOGLE_API_KEY = "AIzaSyCh8ia27PwiWJkPCypoUyvj5TD8YJVjJSc";
const SHEETS_DISCOVERY_DOC = "https://sheets.googleapis.com/$discovery/rest?version=v4";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";

const HEADER_ROW = 6;
const DATA_START_ROW = 7;
const THUMBNAIL_COLUMN = 1;
const REEL_COLUMN = 2;
const TC_IN_COLUMN = 3;
const TC_OUT_COLUMN = 4;
const DURATION_COLUMN = 5;
const DESCRIPTION_COLUMN = 6;
const SOURCE_COLUMN = 7;
const THUMBNAIL_WIDTH = 320;
const THUMBNAIL_HEIGHT = 180;

const $ = (id) => document.getElementById(id);
const state = {
  workbook: null,
  worksheet: null,
  outputBuffer: null,
  images: [],
  ffmpeg: null,
  ffmpegClassWorkerUrl: null,
  ffmpegInputName: null,
  selectedLogFile: null,
  videoMetadata: null,
  googleTokenClient: null,
  googleReady: false,
  googleToken: null,
  googleSheetId: null,
  googleSheetTitle: null,
  googleSheetName: null,
  googleRows: null,
  currentLogMessages: [],
};

document.querySelectorAll("[data-app-version]").forEach((element) => {
  element.textContent = FAIR_USE_LOG_VERSION;
});

function setStatus(message, kind = "working", percent = null) {
  const status = $("statusPill");
  const progressText = $("progressText");
  if (status) {
    status.textContent = kind === "error" ? "Error" : kind === "complete" ? "Complete" : message;
    status.title = message;
    status.className = `status-pill ${kind}`;
  }
  if (progressText) progressText.textContent = message;
  if (percent !== null && $("progressBar")) {
    $("progressBar").style.width = `${Math.max(0, Math.min(100, percent))}%`;
  }
}

function showError(error) {
  const message = error instanceof Error ? error.message : String(error || "Unknown error.");
  setStatus(message, "error", 0);
  if ($("errorDetails")) {
    $("errorDetails").textContent = message;
    $("errorDetails").classList.remove("hidden");
  }
}

function clearError() {
  if ($("errorDetails")) {
    $("errorDetails").textContent = "";
    $("errorDetails").classList.add("hidden");
  }
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function cellText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value).trim();
  if (value.text !== undefined) return String(value.text).trim();
  if (value.result !== undefined) return String(value.result).trim();
  if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("").trim();
  return String(value).trim();
}

function normalizeTimecode(value) {
  return cellText(value).replace(/;/g, ":").trim();
}

function isTimecode(value) {
  return /^\d{2}:\d{2}:\d{2}[:;]\d{2}$/.test(cellText(value));
}

function nominalFrameRate(fps) {
  if (Math.abs(fps - 23.976) < 0.02) return 24;
  if (Math.abs(fps - 29.97) < 0.02) return 30;
  if (Math.abs(fps - 59.94) < 0.02) return 60;
  return Math.max(1, Math.round(fps));
}

function parseRate(value) {
  const text = cellText(value);
  if (!text) return null;
  if (text.includes("/")) {
    const [numerator, denominator] = text.split("/").map(Number);
    if (denominator) return numerator / denominator;
  }
  const result = Number(text.replace(/\s*fps?\s*$/i, ""));
  return Number.isFinite(result) && result > 0 ? result : null;
}

function parseTimecode(value, fps) {
  const raw = cellText(value).trim();
  const dropFrame = raw.includes(";");
  const parts = raw.replace(/;/g, ":").split(":").map(Number);
  if (parts.length !== 4 || parts.some((part) => !Number.isFinite(part))) {
    throw new Error(`Invalid timecode: ${raw || "(blank)"}`);
  }

  const [hours, minutes, seconds, frames] = parts;
  const nominal = nominalFrameRate(fps);
  let total = ((hours * 60 + minutes) * 60 + seconds) * nominal + frames;

  if (dropFrame && (Math.abs(fps - 29.97) < 0.02 || Math.abs(fps - 59.94) < 0.02)) {
    const droppedPerMinute = nominal === 60 ? 4 : 2;
    const totalMinutes = hours * 60 + minutes;
    total -= droppedPerMinute * (totalMinutes - Math.floor(totalMinutes / 10));
  }

  return total;
}

function timecodeFromFrames(frames, fps, dropFrame = false) {
  const nominal = nominalFrameRate(fps);
  let remaining = Math.max(0, Math.floor(frames));
  const hours = Math.floor(remaining / (nominal * 3600));
  remaining -= hours * nominal * 3600;
  const minutes = Math.floor(remaining / (nominal * 60));
  remaining -= minutes * nominal * 60;
  const seconds = Math.floor(remaining / nominal);
  const frame = remaining % nominal;
  const separator = dropFrame ? ";" : ":";
  return [hours, minutes, seconds, frame].map((x) => String(Math.floor(x)).padStart(2, "0")).join(separator);
}

function makeUniqueName(prefix, originalName, extension = "") {
  const safe = cellText(originalName).replace(/[^a-z0-9_.-]/gi, "_").slice(-80) || "reference";
  return `${prefix}-${Date.now()}-${safe}${extension}`;
}

async function toBlobURL(url, mimeType) {
  const response = await fetch(url, { mode: "cors", cache: "no-store" });
  if (!response.ok) throw new Error(`Could not load local video engine asset (${response.status}).`);
  return URL.createObjectURL(new Blob([await response.arrayBuffer()], { type: mimeType }));
}

async function loadExternalScript(url) {
  await new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = url;
    script.async = true;
    script.onload = resolve;
    script.onerror = () => reject(new Error("Could not load the local video engine library."));
    document.head.appendChild(script);
  });
}

async function loadFfmpeg() {
  if (state.ffmpeg) return state.ffmpeg;

  setStatus("Loading local video engine…", "working", 4);
  if (!window.FFmpegWASM?.FFmpeg) await loadExternalScript(FFMPEG_CLASS_URL);
  const FFmpegClass = window.FFmpegWASM?.FFmpeg;
  if (!FFmpegClass) throw new Error("The local video engine library did not load.");

  const ffmpeg = new FFmpegClass();
  ffmpeg.on("log", ({ message }) => {
    state.currentLogMessages.push(message);
    if (state.currentLogMessages.length > 250) state.currentLogMessages.shift();
  });

  // Supply a same-origin Blob URL for the class worker. This prevents the
  // browser from constructing 814.ffmpeg.js directly from jsDelivr.
  state.ffmpegClassWorkerUrl = await toBlobURL(FFMPEG_CLASS_WORKER_URL, "text/javascript");
  await ffmpeg.load({
    classWorkerURL: state.ffmpegClassWorkerUrl,
    coreURL: await toBlobURL(`${FFMPEG_CORE_BASE}/ffmpeg-core.js`, "text/javascript"),
    wasmURL: await toBlobURL(`${FFMPEG_CORE_BASE}/ffmpeg-core.wasm`, "application/wasm"),
  });

  state.ffmpeg = ffmpeg;
  return ffmpeg;
}

function extractJson(text) {
  const source = String(text || "").trim();
  const start = source.indexOf("{");
  const end = source.lastIndexOf("}");
  if (start < 0 || end < start) return null;
  try {
    return JSON.parse(source.slice(start, end + 1));
  } catch {
    return null;
  }
}

function metadataFromProbe(probe) {
  const streams = Array.isArray(probe?.streams) ? probe.streams : [];
  const videoStream = streams.find((stream) => stream.codec_type === "video") || {};
  const formatTags = probe?.format?.tags || {};
  const streamTags = streams.map((stream) => stream.tags || {});
  const timecode = streamTags.map((tags) => tags.timecode).find(Boolean) || formatTags.timecode || null;
  const rate = parseRate(videoStream.avg_frame_rate) || parseRate(videoStream.r_frame_rate);
  if (!timecode || !rate) return null;
  return {
    startTimecode: normalizeTimecode(timecode),
    fps: rate,
    fpsLabel: rate.toFixed(3).replace(/0+$/, "").replace(/\.$/, ""),
    width: videoStream.width || null,
    height: videoStream.height || null,
    duration: probe?.format?.duration || null,
  };
}

async function detectVideoMetadata(file) {
  const ffmpeg = await loadFfmpeg();
  const inputName = makeUniqueName("probe", file.name);
  const outputName = `${inputName}.json`;
  state.currentLogMessages = [];
  await ffmpeg.writeFile(inputName, new Uint8Array(await file.arrayBuffer()));

  try {
    const exitCode = await ffmpeg.ffprobe([
      "-v", "error",
      "-show_entries", "stream=index,codec_type,avg_frame_rate,r_frame_rate,width,height:stream_tags=timecode",
      "-show_entries", "format=duration:format_tags=timecode",
      "-of", "json",
      inputName,
      "-o", outputName,
    ]);
    if (exitCode !== 0) throw new Error(`The video metadata reader returned code ${exitCode}.`);
    const raw = await ffmpeg.readFile(outputName, "utf8");
    const metadata = metadataFromProbe(extractJson(raw));
    if (!metadata) throw new Error("The video opened, but its embedded start timecode or frame rate could not be read.");
    state.videoMetadata = metadata;
    $("videoMetadataStatus").textContent = `Detected automatically — Start TC: ${metadata.startTimecode} · Frame rate: ${metadata.fpsLabel} fps`;
    $("videoMetadataStatus").className = "metadata-status success";
    return metadata;
  } catch (error) {
    const logTail = state.currentLogMessages.filter(Boolean).slice(-6).join(" ");
    throw new Error(`${error.message || error}${logTail ? ` ${logTail}` : ""}`.trim());
  } finally {
    await ffmpeg.deleteFile(inputName).catch(() => {});
    await ffmpeg.deleteFile(outputName).catch(() => {});
  }
}

function findFairUseSheet(workbook) {
  return workbook.worksheets.find((sheet) => cellText(sheet.getCell(`A${HEADER_ROW}`).value).toLowerCase() === "thumbnail") || workbook.worksheets[0];
}

function updateHeaderRows(sheet) {
  sheet.getCell("A1").value = `${$("showTitle").value.trim() || "SHOW TITLE"} — ${$("episodeTitle").value.trim() || "EPISODE"}`;
  sheet.getCell("A2").value = "Fair Use Spreadsheet";
  sheet.getCell("A3").value = $("companyLlc").value.trim();
  sheet.getCell("A4").value = `NETWORK: ${$("network").value.trim()}`;
  sheet.getCell("A5").value = `Shift Link: ${$("referenceLink").value.trim()}`;
}

function getWorkbookRows(sheet) {
  const rows = [];
  for (let row = DATA_START_ROW; row <= sheet.rowCount; row += 1) {
    const tcIn = cellText(sheet.getCell(row, TC_IN_COLUMN).value);
    const tcOut = cellText(sheet.getCell(row, TC_OUT_COLUMN).value);
    if (!tcIn && !tcOut) continue;
    if (!tcIn || !tcOut) throw new Error(`Row ${row} needs both TC IN and TC OUT.`);
    if (!isTimecode(tcIn) || !isTimecode(tcOut)) throw new Error(`Row ${row} has an invalid TC IN or TC OUT.`);
    rows.push({
      row,
      fileName: cellText(sheet.getCell(row, REEL_COLUMN).value),
      tcIn,
      tcOut,
      duration: cellText(sheet.getCell(row, DURATION_COLUMN).value),
      description: cellText(sheet.getCell(row, DESCRIPTION_COLUMN).value),
      source: cellText(sheet.getCell(row, SOURCE_COLUMN).value),
    });
  }
  return rows;
}

async function prepareFfmpegInput(file) {
  const ffmpeg = await loadFfmpeg();
  if (state.ffmpegInputName) await ffmpeg.deleteFile(state.ffmpegInputName).catch(() => {});
  state.ffmpegInputName = makeUniqueName("reference", file.name);
  await ffmpeg.writeFile(state.ffmpegInputName, new Uint8Array(await file.arrayBuffer()));
  return state.ffmpegInputName;
}

async function makeScreenshot(inputName, offsetFrames, fps, outputName) {
  const ffmpeg = await loadFfmpeg();
  const seconds = Math.max(0, offsetFrames / fps);
  const exitCode = await ffmpeg.exec([
    "-hide_banner", "-loglevel", "error", "-ss", seconds.toFixed(6),
    "-i", inputName,
    "-frames:v", "1",
    "-vf", "scale=320:180:force_original_aspect_ratio=decrease,pad=320:180:(ow-iw)/2:(oh-ih)/2:color=black",
    "-q:v", "3", "-y", outputName,
  ]);
  if (exitCode !== 0) throw new Error(`FFmpeg could not extract the frame at ${seconds.toFixed(3)} seconds.`);
  const data = await ffmpeg.readFile(outputName);
  await ffmpeg.deleteFile(outputName).catch(() => {});
  return new Uint8Array(data);
}

function addScreenshotToExcel(sheet, rowNumber, bytes) {
  const imageId = state.workbook.addImage({ buffer: bytes.buffer, extension: "jpeg" });
  sheet.addImage(imageId, { tl: { col: 0, row: rowNumber - 1 }, ext: { width: THUMBNAIL_WIDTH, height: THUMBNAIL_HEIGHT } });
  sheet.getRow(rowNumber).height = 140;
  sheet.getColumn(1).width = 46;
}

function appendThumbnailCard(row, midpoint, bytes) {
  const card = document.createElement("figure");
  card.className = "thumbnail-card";
  const image = document.createElement("img");
  image.src = URL.createObjectURL(new Blob([bytes], { type: "image/jpeg" }));
  image.alt = `Row ${row} screenshot at ${midpoint}`;
  const caption = document.createElement("figcaption");
  caption.textContent = `Row ${row} · ${midpoint}`;
  card.append(image, caption);
  $("thumbnailGrid").append(card);
}

async function processWorkbook(workbook, sheet, rows, videoFile) {
  updateHeaderRows(sheet);
  state.workbook = workbook;
  state.worksheet = sheet;
  state.images = [];
  const metadata = state.videoMetadata || await detectVideoMetadata(videoFile);
  const startFrames = parseTimecode(metadata.startTimecode, metadata.fps);
  const inputName = await prepareFfmpegInput(videoFile);

  $("resultsCard").classList.remove("hidden");
  $("thumbnailGrid").innerHTML = "";
  for (let index = 0; index < rows.length; index += 1) {
    const item = rows[index];
    const inFrames = parseTimecode(item.tcIn, metadata.fps);
    let outFrames = parseTimecode(item.tcOut, metadata.fps);
    if (outFrames < inFrames) outFrames += nominalFrameRate(metadata.fps) * 24 * 60 * 60;
    const midpointFrames = Math.floor((inFrames + outFrames) / 2);
    const offsetFrames = midpointFrames - startFrames;
    if (offsetFrames < 0) throw new Error(`Row ${item.row}: midpoint occurs before the reference cut start timecode.`);
    const midpoint = timecodeFromFrames(midpointFrames, metadata.fps, item.tcIn.includes(";") || item.tcOut.includes(";"));
    const bytes = await makeScreenshot(inputName, offsetFrames, metadata.fps, `frame-${item.row}.jpg`);
    addScreenshotToExcel(sheet, item.row, bytes);
    state.images.push({ row: item.row, midpoint, bytes });
    appendThumbnailCard(item.row, midpoint, bytes);
    setStatus(`Creating screenshot ${index + 1} of ${rows.length}…`, "working", 10 + ((index + 1) / rows.length) * 85);
  }
}

function parseSheetUrl(value) {
  const match = String(value || "").match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) throw new Error("Please enter a valid Google Sheet URL.");
  return match[1];
}

function initializeGoogle() {
  if (state.googleTokenClient || state.googleReady) return;
  if (!window.gapi || !window.google?.accounts?.oauth2) {
    setTimeout(initializeGoogle, 250);
    return;
  }
  gapi.load("client", async () => {
    try {
      await gapi.client.init({ apiKey: GOOGLE_API_KEY, discoveryDocs: [SHEETS_DISCOVERY_DOC] });
      state.googleReady = true;
      $("connectGoogleButton").disabled = false;
    } catch (error) {
      $("googleStatus").textContent = `Google connection unavailable: ${error.message || error}`;
    }
  });
  state.googleTokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: SHEETS_SCOPE,
    callback: "",
  });
}

function connectGoogle() {
  if (!state.googleTokenClient) throw new Error("Google connection is still loading. Please try again in a moment.");
  return new Promise((resolve, reject) => {
    state.googleTokenClient.callback = (response) => {
      if (response.error) {
        reject(new Error(`Google authorization failed: ${response.error}`));
        return;
      }
      state.googleToken = response.access_token;
      gapi.client.setToken({ access_token: state.googleToken });
      $("googleStatus").textContent = "Google connected. The selected Sheet can be read and updated.";
      $("connectGoogleButton").textContent = "Google Connected";
      resolve();
    };
    state.googleTokenClient.requestAccessToken({ prompt: gapi.client.getToken() ? "" : "consent" });
  });
}

async function loadGoogleSheetRows() {
  const spreadsheetId = parseSheetUrl($("driveSheetUrl").value);
  if (!state.googleToken) await connectGoogle();
  setStatus("Reading Google Sheet…", "working", 8);
  const response = await gapi.client.sheets.spreadsheets.get({
    spreadsheetId,
    includeGridData: true,
    fields: "properties(title),sheets(properties(sheetId,title),data(startRow,startColumn,rowData(values(formattedValue,effectiveValue))))",
  });
  const sourceSheet = response.result.sheets?.find((sheet) => {
    const firstRows = sheet.data?.[0]?.rowData || [];
    return String(firstRows[HEADER_ROW - 1]?.values?.[0]?.formattedValue || "").trim().toLowerCase() === "thumbnail";
  }) || response.result.sheets?.[0];
  if (!sourceSheet) throw new Error("No worksheet was found in the Google Sheet.");

  state.googleSheetId = spreadsheetId;
  state.googleSheetTitle = response.result.properties?.title || "Google Fair Use Log";
  state.googleSheetName = sourceSheet.properties?.title || "Sheet1";
  const rowData = sourceSheet.data?.[0]?.rowData || [];
  state.googleRows = rowData.map((row, index) => ({
    row: index + 1,
    values: (row.values || []).map((cell) => cell.formattedValue || ""),
  }));
  const rows = state.googleRows.filter((entry) => entry.row >= DATA_START_ROW && entry.values[TC_IN_COLUMN - 1] && entry.values[TC_OUT_COLUMN - 1]);
  if (!rows.length) throw new Error("No Google Sheet rows were found with both TC IN and TC OUT.");
  $("selectedLogStatus").textContent = `Selected Google Sheet: ${state.googleSheetTitle} · ${state.googleSheetName}`;
  return rows;
}

async function createGoogleInputWorkbook() {
  const rows = await loadGoogleSheetRows();
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(state.googleSheetName || "Fair Use Log");
  state.googleRows.forEach((entry) => {
    entry.values.forEach((value, columnIndex) => {
      sheet.getCell(entry.row, columnIndex + 1).value = value;
    });
  });
  if (!sheet.getCell(`A${HEADER_ROW}`).value) sheet.getCell(`A${HEADER_ROW}`).value = "Thumbnail";
  return { workbook, sheet, rows };
}

async function writeGoogleHeaderResults() {
  if (!state.googleSheetId || !state.googleSheetName) return;
  const values = [[
    `${$("showTitle").value.trim() || "SHOW TITLE"} — ${$("episodeTitle").value.trim() || "EPISODE"}`,
    "Fair Use Spreadsheet",
    $("companyLlc").value.trim(),
    `NETWORK: ${$("network").value.trim()}`,
    `Shift Link: ${$("referenceLink").value.trim()}`,
  ]];
  const encodedName = state.googleSheetName.replace(/'/g, "''");
  await gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId: state.googleSheetId,
    range: `'${encodedName}'!A1:A5`,
    valueInputOption: "USER_ENTERED",
    resource: { values },
  });
}

async function createLog() {
  clearError();
  const videoFile = $("referenceCut").files[0];
  if (!videoFile) throw new Error("Please choose a reference cut.");

  const localFile = state.selectedLogFile || $("logFile").files[0];
  if (localFile) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await localFile.arrayBuffer());
    const sheet = findFairUseSheet(workbook);
    const rows = getWorkbookRows(sheet);
    if (!rows.length) throw new Error("No rows were found with both TC IN and TC OUT timecodes.");
    await processWorkbook(workbook, sheet, rows, videoFile);
    state.outputBuffer = await workbook.xlsx.writeBuffer();
    $("resultTitle").textContent = "Screenshots ready";
    $("resultSummary").textContent = `${rows.length} midpoint screenshots were generated locally from ${videoFile.name}.`;
    setStatus("Complete", "complete", 100);
    return;
  }

  if (!$("driveSheetUrl").value.trim()) throw new Error("Choose a local Fair Use Log or enter a Google Sheet URL.");
  const imported = await createGoogleInputWorkbook();
  await processWorkbook(imported.workbook, imported.sheet, imported.rows, videoFile);
  state.outputBuffer = await imported.workbook.xlsx.writeBuffer();
  await writeGoogleHeaderResults();
  $("resultTitle").textContent = "Screenshots ready";
  $("resultSummary").textContent = `${imported.rows.length} midpoint screenshots were generated locally. Download the Excel workbook to retain the embedded screenshots.`;
  setStatus("Complete", "complete", 100);
}

async function downloadImages() {
  const zip = new JSZip();
  state.images.forEach((image) => zip.file(`Row${String(image.row).padStart(4, "0")}_${image.midpoint.replace(/[:;]/g, "-")}.jpg`, image.bytes));
  downloadBlob(await zip.generateAsync({ type: "blob" }), "Fair_Use_Screenshots.zip");
}

$("referenceCut").addEventListener("change", () => {
  state.videoMetadata = null;
  const file = $("referenceCut").files[0];
  if (!file) {
    $("videoMetadataStatus").textContent = "Choose a reference cut. Start timecode and frame rate will be detected automatically when you create the log.";
    $("videoMetadataStatus").className = "metadata-status";
    return;
  }
  $("videoMetadataStatus").textContent = `Reference cut selected: ${file.name}. Metadata will be detected automatically when you create the log.`;
  $("videoMetadataStatus").className = "metadata-status";
});

$("logFile").addEventListener("change", () => {
  state.selectedLogFile = $("logFile").files[0] || null;
  if (state.selectedLogFile) {
    $("driveSheetUrl").value = "";
    $("selectedLogStatus").textContent = `Selected local Fair Use Log: ${state.selectedLogFile.name}`;
  }
});

$("driveSheetUrl").addEventListener("input", () => {
  if ($("driveSheetUrl").value.trim()) {
    state.selectedLogFile = null;
    $("logFile").value = "";
    $("selectedLogStatus").textContent = "Google Sheet URL entered. Connect Google when ready.";
  }
});

$("connectGoogleButton").addEventListener("click", async () => {
  try {
    await connectGoogle();
    if ($("driveSheetUrl").value.trim()) await loadGoogleSheetRows();
  } catch (error) {
    showError(error);
  }
});

$("createButton").addEventListener("click", async () => {
  $("createButton").disabled = true;
  try {
    await createLog();
  } catch (error) {
    showError(error);
  } finally {
    $("createButton").disabled = false;
  }
});

$("downloadButton").addEventListener("click", () => {
  if (!state.outputBuffer) return;
  downloadBlob(new Blob([state.outputBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "FAIR_USE_GENERATED.xlsx");
});

$("downloadImagesButton").addEventListener("click", downloadImages);
$("cancelButton").addEventListener("click", () => window.location.reload());

initializeGoogle();
