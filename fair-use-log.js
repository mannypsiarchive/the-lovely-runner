const FAIR_USE_LOG_VERSION = "1.2";
const FFmpeg = window.FFmpegWASM?.FFmpeg;

async function fetchFile(file) {
  return new Uint8Array(await file.arrayBuffer());
}

async function toBlobURL(url, mimeType) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Could not load FFmpeg asset (${response.status}).`);
  return URL.createObjectURL(new Blob([await response.arrayBuffer()], { type: mimeType }));
}

const $ = (id) => document.getElementById(id);
const state = { workbook: null, worksheet: null, outputBuffer: null, images: [], ffmpeg: null, driveTokenClient: null, driveReady: false, gapiReady: false, accessToken: null, selectedLogFile: null, videoMetadata: null };

document.querySelectorAll("[data-app-version]").forEach((element) => { element.textContent = FAIR_USE_LOG_VERSION; });

const GOOGLE_CLIENT_ID = "154634144934-9hg9o4ra7uriu5hrivaaj73mduj7udf4.apps.googleusercontent.com";
const GOOGLE_API_KEY = "AIzaSyCh8ia27PwiWJkPCypoUyvj5TD8YJVjJSc";
const SHEETS_DISCOVERY_DOC = "https://sheets.googleapis.com/$discovery/rest?version=v4";
const DRIVE_DISCOVERY_DOC = "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest";
const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.readonly";

function setStatus(text, percent = null) {
  $("statusPill").textContent = percent === 0 ? "Error" : text;
  $("progressText").textContent = text;
  if (percent !== null) $("progressBar").style.width = `${Math.max(0, Math.min(100, percent))}%`;
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url; link.download = name; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function normalizeTimecode(value) { return String(value || "").trim().replaceAll(";", ":"); }

function parseTimecode(value, fps) {
  const p = normalizeTimecode(value).split(":").map(Number);
  if (p.length !== 4 || p.some(Number.isNaN)) throw new Error(`Invalid timecode: ${value}`);
  return (((p[0] * 60 + p[1]) * 60 + p[2]) * fps) + p[3];
}

function framesToTimecode(frames, fps) {
  frames = Math.max(0, Math.floor(frames));
  const h = Math.floor(frames / (fps * 3600)); frames -= h * fps * 3600;
  const m = Math.floor(frames / (fps * 60)); frames -= m * fps * 60;
  const s = Math.floor(frames / fps); const f = frames - s * fps;
  return [h, m, s, f].map((x) => String(Math.floor(x)).padStart(2, "0")).join(":");
}

function findSheet(workbook) {
  return workbook.worksheets.find((sheet) => String(sheet.getCell("A6").value || "").trim().toLowerCase() === "thumbnail") || workbook.worksheets[0];
}

async function loadFFmpeg() {
  if (state.ffmpeg) return state.ffmpeg;
  if (!FFmpeg) throw new Error("The local FFmpeg browser engine did not load. Please refresh the page and try again.");
  const ffmpeg = new FFmpeg();
  ffmpeg.on("log", ({ message }) => { state.lastProbeLog = `${state.lastProbeLog || ""}\n${message}`; });
  setStatus("Loading local video engine…", 5);
  const coreBaseURL = "https://cdn.jsdelivr.net/npm/@ffmpeg/core@0.12.10/dist/umd";
  const classWorkerURL = await toBlobURL("https://cdn.jsdelivr.net/npm/@ffmpeg/ffmpeg@0.12.10/dist/umd/814.ffmpeg.js", "text/javascript");
  await ffmpeg.load({
    classWorkerURL,
    coreURL: await toBlobURL(`${coreBaseURL}/ffmpeg-core.js`, "text/javascript"),
    wasmURL: await toBlobURL(`${coreBaseURL}/ffmpeg-core.wasm`, "application/wasm"),
  });
  state.ffmpeg = ffmpeg;
  return ffmpeg;
}

async function detectVideoMetadata(file) {
  const ffmpeg = await loadFFmpeg();
  const name = `probe-${Date.now()}-${file.name.replace(/[^a-z0-9.]/gi, "_")}`;
  await ffmpeg.writeFile(name, await fetchFile(file));
  const output = `${name}.json`;
  try {
    await ffmpeg.ffprobe([
      "-v", "error",
      "-select_streams", "v:0",
      "-show_entries", "stream=avg_frame_rate,r_frame_rate,width,height:stream_tags=timecode",
      "-show_entries", "format=duration:format_tags=timecode",
      "-of", "json",
      name,
      "-o", output,
    ]);
    const raw = await ffmpeg.readFile(output, "utf8");
    const metadata = JSON.parse(typeof raw === "string" ? raw : new TextDecoder().decode(raw));
    const stream = metadata.streams?.[0] || {};
    const streamTags = stream.tags || {};
    const formatTags = metadata.format?.tags || {};
    const rate = stream.avg_frame_rate && stream.avg_frame_rate !== "0/0" ? stream.avg_frame_rate : stream.r_frame_rate;
    const fps = rate ? Number(rate.split("/")[0]) / Number(rate.split("/")[1]) : null;
    const startTimecode = streamTags.timecode || formatTags.timecode;
    if (!startTimecode || !fps || !Number.isFinite(fps)) throw new Error("The video does not contain readable embedded start timecode and frame-rate metadata.");
    state.videoMetadata = { startTimecode, fps, fpsLabel: fps.toFixed(3).replace(/0+$/, "").replace(/\.$/, ""), width: stream.width, height: stream.height, duration: metadata.format?.duration };
    $("videoMetadataStatus").textContent = `Detected automatically — Start TC: ${state.videoMetadata.startTimecode} · Frame rate: ${state.videoMetadata.fpsLabel} fps`;
    return state.videoMetadata;
  } finally {
    await ffmpeg.deleteFile(name).catch(() => {});
    await ffmpeg.deleteFile(output).catch(() => {});
  }
}

async function makeScreenshot(file, seconds, index) {
  const ffmpeg = await loadFFmpeg();
  const input = `reference-${Date.now()}-${file.name.replace(/[^a-z0-9.]/gi, "_")}`;
  const output = `frame-${index}.jpg`;
  await ffmpeg.writeFile(input, await fetchFile(file));
  try {
    await ffmpeg.exec(["-hide_banner", "-loglevel", "error", "-ss", seconds.toFixed(6), "-i", input, "-frames:v", "1", "-q:v", "3", "-y", output]);
    const data = await ffmpeg.readFile(output);
    return new Uint8Array(data);
  } finally {
    await ffmpeg.deleteFile(input).catch(() => {});
    await ffmpeg.deleteFile(output).catch(() => {});
  }
}

function imageUrl(bytes) { return URL.createObjectURL(new Blob([bytes], { type: "image/jpeg" })); }

function addScreenshot(worksheet, rowNumber, bytes) {
  const imageId = state.workbook.addImage({ buffer: bytes.buffer, extension: "jpeg" });
  worksheet.addImage(imageId, { tl: { col: 0, row: rowNumber - 1 }, ext: { width: 320, height: 180 } });
  worksheet.getRow(rowNumber).height = 140;
  worksheet.getColumn(1).width = 46;
}

async function createLog() {
  let logFile = state.selectedLogFile || $("logFile").files[0];
  if (!logFile && $("driveSheetUrl").value.trim()) {
    await useDriveSheet();
    logFile = state.selectedLogFile;
  }
  const videoFile = $("referenceCut").files[0];
  if (!logFile || !videoFile) throw new Error("Please upload both a Fair Use Log workbook and a reference cut.");
  if (!state.videoMetadata) await detectVideoMetadata(videoFile);
  const fps = state.videoMetadata.fps;
  const startTc = normalizeTimecode(state.videoMetadata.startTimecode);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await logFile.arrayBuffer());
  const worksheet = findSheet(workbook);
  worksheet.getCell("A1").value = `${$("showTitle").value.trim() || "SHOW TITLE"} — ${$("episodeTitle").value.trim() || "EPISODE"}`;
  worksheet.getCell("A2").value = "Fair Use Spreadsheet";
  worksheet.getCell("A3").value = $("companyLlc").value.trim();
  worksheet.getCell("A4").value = `NETWORK: ${$("network").value.trim()}`;
  worksheet.getCell("A5").value = `Shift Link: ${$("referenceLink").value.trim()}`;
  state.workbook = workbook; state.worksheet = worksheet; state.images = [];

  const rows = [];
  for (let row = 7; row <= worksheet.rowCount; row++) {
    const input = worksheet.getCell(row, 3).value;
    const output = worksheet.getCell(row, 4).value;
    if (input && output) rows.push({ row, input: String(input), output: String(output) });
  }
  if (!rows.length) throw new Error("No rows were found with both Dst In and Dst Out timecodes.");
  $("resultsCard").classList.remove("hidden"); $("thumbnailGrid").innerHTML = "";
  let done = 0;
  for (const item of rows) {
    const midpointFrames = Math.floor((parseTimecode(item.input, fps) + parseTimecode(item.output, fps)) / 2);
    const midpoint = framesToTimecode(midpointFrames, fps);
    const offsetFrames = parseTimecode(midpoint, fps) - parseTimecode(startTc, fps);
    if (offsetFrames < 0) throw new Error(`Row ${item.row}: midpoint occurs before the reference cut start timecode.`);
    const bytes = await makeScreenshot(videoFile, offsetFrames / fps, item.row);
    addScreenshot(worksheet, item.row, bytes); state.images.push({ row: item.row, midpoint, bytes });
    const card = document.createElement("figure"); card.className = "thumbnail-card";
    const img = document.createElement("img"); img.src = imageUrl(bytes); img.alt = `Row ${item.row} midpoint ${midpoint}`;
    const caption = document.createElement("figcaption"); caption.textContent = `Row ${item.row} · ${midpoint}`;
    card.append(img, caption); $("thumbnailGrid").append(card);
    done++; setStatus(`Creating screenshot ${done} of ${rows.length}…`, 10 + (done / rows.length) * 85);
  }
  state.outputBuffer = await workbook.xlsx.writeBuffer();
  $("resultTitle").textContent = "Screenshots ready";
  $("resultSummary").textContent = `${done} midpoint screenshots were generated locally. Review the thumbnails below, then download the generated workbook.`;
  setStatus("Complete", 100);
}

async function downloadImages() {
  const zip = new JSZip();
  state.images.forEach((image) => zip.file(`Row${String(image.row).padStart(4, "0")}_${image.midpoint.replaceAll(":", "-")}.jpg`, image.bytes));
  downloadBlob(await zip.generateAsync({ type: "blob" }), "Fair_Use_Screenshots.zip");
}

function parseFolderId(value) { return String(value || "").match(/folders\/([a-zA-Z0-9_-]+)/)?.[1] || String(value || "").trim(); }

async function initDrive() {
  if (!window.gapi || !window.google?.accounts?.oauth2) return setTimeout(initDrive, 250);
  gapi.load("client", async () => {
    await gapi.client.init({ apiKey: GOOGLE_API_KEY, discoveryDocs: [SHEETS_DISCOVERY_DOC, DRIVE_DISCOVERY_DOC] });
    state.gapiReady = true; maybeEnableDrive();
  });
  state.driveTokenClient = google.accounts.oauth2.initTokenClient({ client_id: GOOGLE_CLIENT_ID, scope: DRIVE_SCOPE, callback: "" });
  state.driveReady = true; maybeEnableDrive();
}

function maybeEnableDrive() {
  const signInButton = $("signInButton");
  if (signInButton) signInButton.disabled = !(state.driveReady && state.gapiReady);
}

async function connectDrive() {
  state.driveTokenClient.callback = async (response) => {
    if (response.error) throw new Error(response.error);
    state.accessToken = response.access_token;
    gapi.client.setToken({ access_token: response.access_token });
    if ($("driveFolderField")) {
      $("driveFolderField").classList.remove("hidden");
      $("driveStatus").textContent = "Loading Drive folders…";
      const response2 = await gapi.client.drive.files.list({ q: "mimeType='application/vnd.google-apps.folder' and trashed=false", pageSize: 100, orderBy: "name", fields: "files(id,name)" });
      const select = $("driveFolder"); select.innerHTML = '<option value="">Choose a folder</option>';
      (response2.result.files || []).forEach((folder) => { const option = document.createElement("option"); option.value = folder.id; option.textContent = folder.name; select.append(option); });
      $("driveStatus").textContent = "Google Drive connected. Choose a destination folder.";
    }
    await loadDriveLogs();
  };
  state.driveTokenClient.requestAccessToken({ prompt: gapi.client.getToken() ? "" : "consent" });
}

async function loadDriveLogs() {
  if (!state.accessToken) return connectDrive();
  const response = await gapi.client.drive.files.list({
    q: "trashed=false and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.google-apps.spreadsheet')",
    pageSize: 100,
    orderBy: "name",
    fields: "files(id,name,mimeType,modifiedTime)",
  });
  const select = $("driveLogFile");
  select.innerHTML = '<option value="">Choose a Google Drive file</option>';
  (response.result.files || []).forEach((file) => {
    const option = document.createElement("option");
    option.value = JSON.stringify(file);
    option.textContent = `${file.name} (${file.mimeType === "application/vnd.google-apps.spreadsheet" ? "Google Sheet" : "Excel"})`;
    select.append(option);
  });
  $("driveLogField").classList.remove("hidden");
  $("selectedLogStatus").textContent = "Choose a Fair Use Log from the Google Drive list.";
}

async function fetchDriveLog(file) {
  const headers = { Authorization: `Bearer ${state.accessToken}` };
  const url = file.mimeType === "application/vnd.google-apps.spreadsheet"
    ? `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
    : `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(file.id)}?alt=media`;
  const response = await fetch(url, { headers });
  if (!response.ok) throw new Error(`Google Drive could not provide that file (${response.status}).`);
  const bytes = await response.arrayBuffer();
  return new File([bytes], file.name.replace(/\.xlsx$/i, "") + ".xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function parseSheetUrl(value) {
  const match = String(value || "").match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) throw new Error("Please enter a valid Google Sheet URL.");
  return match[1];
}

async function useDriveSheet() {
  const spreadsheetId = parseSheetUrl($("driveSheetUrl").value);
  const loadSheet = async () => {
    $("selectedLogStatus").textContent = "Reading Fair Use Log from Google Drive…";
    state.selectedLogFile = await fetchDriveLog({
      id: spreadsheetId,
      name: "Google Drive Fair Use Log",
      mimeType: "application/vnd.google-apps.spreadsheet",
    });
    $("selectedLogStatus").textContent = "Selected Google Drive Fair Use Log.";
  };

  if (state.accessToken) {
    await loadSheet();
    return;
  }

  if (!state.driveTokenClient) throw new Error("Google authorization is still loading. Please try again in a moment.");
  await new Promise((resolve, reject) => {
    state.driveTokenClient.callback = async (response) => {
      try {
        if (response.error) throw new Error(response.error);
        state.accessToken = response.access_token;
        gapi.client.setToken({ access_token: state.accessToken });
        await loadSheet();
        resolve();
      } catch (error) {
        reject(error);
      }
    };
    state.driveTokenClient.requestAccessToken({ prompt: gapi.client.getToken() ? "" : "consent" });
  });
}

async function createDriveTemplate() {
  const folderId = $("driveFolder").value; if (!folderId) throw new Error("Choose a Google Drive folder first.");
  const title = `AEFS - SAMPLE FAIR USE LOG - ${new Date().toISOString().slice(0, 10)}`;
  const created = await gapi.client.sheets.spreadsheets.create({ resource: { properties: { title } } });
  const spreadsheetId = created.result.spreadsheetId;
  await gapi.client.drive.files.update({ fileId: spreadsheetId, addParents: folderId, fields: "id,parents" });
  const values = [
    ["SHOW TITLE AND EPISODE"], ["Fair Use Spreadsheet"], ["COMPANY LLC"], ["NETWORK:"], ["Shift Link: LINK TO CUT"],
    ["Thumbnail", "Reel", "Dst In", "Dst Out", "Src Dur", "Description", "Source", "Supporting Language", "Due Diligence Item"],
    ["", "FILE NUMBER_THE LAST SUPPER.JPG", "01:15:04:23", "01:15:11:02", "00:00:06:04", "Leonardo DaVinci's the Last Supper", "Milan Museum", "VO: THEN OF COURSE YOU HAVE THAT PAINTING, THE LAST SUPPER BY LEONARDO DAVINCI.", ""]
  ];
  await gapi.client.sheets.spreadsheets.values.update({ spreadsheetId, range: "Sheet1!A1:I8", valueInputOption: "USER_ENTERED", resource: { values } });
  await gapi.client.sheets.spreadsheets.batchUpdate({ spreadsheetId, resource: { requests: [{ updateSheetProperties: { properties: { sheetId: 0, title: "Fair Use Log" }, fields: "title" } }, { repeatCell: { range: { sheetId: 0, startRowIndex: 5, endRowIndex: 6, startColumnIndex: 0, endColumnIndex: 9 }, cell: { userEnteredFormat: { backgroundColor: { red: 0.12, green: 0.18, blue: 0.35 }, textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } } } }, fields: "userEnteredFormat(backgroundColor,textFormat)" } }] } });
  window.open(`https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`, "_blank");
  $("driveStatus").textContent = "Google Sheet template created and opened.";
}

$("referenceCut").addEventListener("change", async () => {
  const file = $("referenceCut").files[0];
  state.videoMetadata = null;
  if (!file) return;
  $("videoMetadataStatus").textContent = "Detecting embedded start timecode and frame rate…";
  try {
    await detectVideoMetadata(file);
  } catch (error) {
    $("videoMetadataStatus").textContent = error.message || "Could not detect video metadata.";
  }
});
$("logFile").addEventListener("change", () => { state.selectedLogFile = $("logFile").files[0] || null; if (state.selectedLogFile) { $("selectedLogStatus").textContent = `Selected local Fair Use Log: ${state.selectedLogFile.name}`; } });
$("createButton").addEventListener("click", async () => { try { $("createButton").disabled = true; await createLog(); } catch (error) { setStatus(error.message || String(error), 0); } finally { $("createButton").disabled = false; } });
$("downloadButton").addEventListener("click", () => downloadBlob(new Blob([state.outputBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), "FAIR_USE_GENERATED.xlsx"));
$("downloadImagesButton").addEventListener("click", downloadImages);
$("cancelButton").addEventListener("click", () => window.location.reload());
initDrive();
