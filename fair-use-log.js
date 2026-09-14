/*
  Fair Use Log 3.0

  This first browser test deliberately does not use FFmpeg or Google APIs.
  It proves the local Excel workflow with two browser capabilities:

  - MediaInfo.js reads the reference cut's embedded start timecode and frame rate.
  - The browser's native <video> element seeks the reference cut and Canvas captures
    the complete frame, including a visible BITC in the upper-right corner.

  The workbook and video remain local to the user's browser. Google Sheet input and
  due-diligence document generation can be added after this local Excel path works.
*/

const FAIR_USE_LOG_VERSION = "3.0";

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
  rows: [],
  midpoints: [],
  referenceFile: null,
  videoMetadata: null,
  video: null,
  videoUrl: null,
  outputBuffer: null,
  images: [],
  previewUrls: [],
  logReadToken: 0,
};

document.querySelectorAll("[data-app-version]").forEach((element) => {
  element.textContent = FAIR_USE_LOG_VERSION;
});

function cellText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value).trim();
  }
  if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("").trim();
  if (value.text !== undefined) return String(value.text).trim();
  if (value.result !== undefined) return String(value.result).trim();
  return String(value).trim();
}

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

function normalizeTimecode(value) {
  return cellText(value).replace(/\s+/g, "").replace(/\./g, ":");
}

function extractTimecode(value) {
  const match = cellText(value).match(/\d{2}:\d{2}:\d{2}[:;]\d{2}/);
  return match ? normalizeTimecode(match[0]) : "";
}

function isTimecode(value) {
  return /^\d{2}:\d{2}:\d{2}[:;]\d{2}$/.test(normalizeTimecode(value));
}

function nominalFrameRate(fps) {
  if (Math.abs(fps - 23.976) < 0.02) return 24;
  if (Math.abs(fps - 29.97) < 0.02) return 30;
  if (Math.abs(fps - 59.94) < 0.02) return 60;
  return Math.max(1, Math.round(fps));
}

function displayFrameRate(fps) {
  const common = [
    [23.976, "23.976"],
    [24, "24"],
    [25, "25"],
    [29.97, "29.97"],
    [30, "30"],
    [50, "50"],
    [59.94, "59.94"],
    [60, "60"],
  ];
  const match = common.find(([target]) => Math.abs(fps - target) < 0.02);
  if (match) return match[1];
  return Number.isInteger(fps) ? String(fps) : fps.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
}

function parseRate(value) {
  const text = cellText(value).replace(/,/g, "");
  if (!text) return null;
  const fraction = text.match(/(\d+(?:\.\d+)?)\s*\/\s*(\d+(?:\.\d+)?)/);
  if (fraction) {
    const numerator = Number(fraction[1]);
    const denominator = Number(fraction[2]);
    if (denominator) return numerator / denominator;
  }
  const number = text.match(/\d+(?:\.\d+)?/);
  const result = number ? Number(number[0]) : NaN;
  return Number.isFinite(result) && result > 0 ? result : null;
}

function parseTimecode(value, fps) {
  const raw = normalizeTimecode(value);
  const dropFrame = raw.includes(";");
  const parts = raw.replace(/;/g, ":").split(":").map(Number);
  const nominal = nominalFrameRate(fps);
  if (parts.length !== 4 || parts.some((part) => !Number.isFinite(part))) {
    throw new Error(`Invalid timecode: ${raw || "(blank)"}`);
  }
  const [hours, minutes, seconds, frames] = parts;
  if (minutes > 59 || seconds > 59 || frames >= nominal) {
    throw new Error(`Invalid timecode: ${raw}`);
  }

  let total = ((hours * 60 + minutes) * 60 + seconds) * nominal + frames;
  if (dropFrame && (Math.abs(fps - 29.97) < 0.02 || Math.abs(fps - 59.94) < 0.02)) {
    const droppedPerMinute = nominal === 60 ? 4 : 2;
    const totalMinutes = hours * 60 + minutes;
    total -= droppedPerMinute * (totalMinutes - Math.floor(totalMinutes / 10));
  }
  return total;
}

function timecodeFromFrames(frameNumber, fps, dropFrame = false) {
  const nominal = nominalFrameRate(fps);
  let frames = Math.max(0, Math.floor(frameNumber));
  const separator = dropFrame ? ";" : ":";

  if (dropFrame && (Math.abs(fps - 29.97) < 0.02 || Math.abs(fps - 59.94) < 0.02)) {
    const dropped = nominal === 60 ? 4 : 2;
    const framesPerMinute = nominal * 60 - dropped;
    const framesPerTenMinutes = nominal * 60 * 10 - dropped * 9;
    const tenMinuteBlocks = Math.floor(frames / framesPerTenMinutes);
    let remainder = frames % framesPerTenMinutes;
    let minuteInBlock;
    let frameInMinute;

    if (remainder < nominal * 60) {
      minuteInBlock = 0;
      frameInMinute = remainder;
    } else {
      remainder -= nominal * 60;
      minuteInBlock = 1 + Math.floor(remainder / framesPerMinute);
      minuteInBlock = Math.min(9, minuteInBlock);
      frameInMinute = remainder - (minuteInBlock - 1) * framesPerMinute + dropped;
    }

    const totalMinutes = tenMinuteBlocks * 10 + minuteInBlock;
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    const seconds = Math.floor(frameInMinute / nominal);
    const frame = frameInMinute % nominal;
    const prefix = [hours, minutes, seconds].map((part) => String(part).padStart(2, "0")).join(":");
    return `${prefix}${separator}${String(frame).padStart(2, "0")}`;
  }

  const hours = Math.floor(frames / (nominal * 3600));
  frames -= hours * nominal * 3600;
  const minutes = Math.floor(frames / (nominal * 60));
  frames -= minutes * nominal * 60;
  const seconds = Math.floor(frames / nominal);
  const frame = frames % nominal;
  const prefix = [hours, minutes, seconds].map((part) => String(part).padStart(2, "0")).join(":");
  return `${prefix}${separator}${String(frame).padStart(2, "0")}`;
}

function calculateMidpoint(item, fps) {
  const inFrames = parseTimecode(item.tcIn, fps);
  let outFrames = parseTimecode(item.tcOut, fps);
  const dayFrames = nominalFrameRate(fps) * 24 * 60 * 60;
  if (outFrames < inFrames) outFrames += dayFrames;
  const midpointFrames = outFrames - inFrames <= 2
    ? inFrames
    : Math.floor((inFrames + outFrames) / 2);
  const dropFrame = normalizeTimecode(item.tcIn).includes(";") || normalizeTimecode(item.tcOut).includes(";");
  return {
    ...item,
    inFrames,
    outFrames,
    midpointFrames,
    midpoint: timecodeFromFrames(midpointFrames, fps, dropFrame),
  };
}

function findFairUseSheet(workbook) {
  return workbook.worksheets.find((sheet) =>
    cellText(sheet.getCell(`A${HEADER_ROW}`).value).toLowerCase() === "thumbnail"
  ) || workbook.worksheets[0];
}

function getWorkbookRows(sheet) {
  const rows = [];
  for (let row = DATA_START_ROW; row <= sheet.rowCount; row += 1) {
    const tcIn = normalizeTimecode(sheet.getCell(row, TC_IN_COLUMN).value);
    const tcOut = normalizeTimecode(sheet.getCell(row, TC_OUT_COLUMN).value);
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

function findTrack(tracks, type) {
  return tracks.find((track) => cellText(track?.["@type"]).toLowerCase() === type.toLowerCase()) || {};
}

function trackValue(track, names) {
  const keys = Object.keys(track || {});
  for (const name of names) {
    const key = keys.find((candidate) => candidate.toLowerCase() === name.toLowerCase());
    if (key && cellText(track[key])) return track[key];
  }
  return "";
}

function resultTracks(result) {
  if (typeof result === "string") {
    try { result = JSON.parse(result); } catch { return []; }
  }
  const tracks = result?.media?.track || result?.Media?.track || result?.media?.Track || [];
  return Array.isArray(tracks) ? tracks : [tracks].filter(Boolean);
}

function metadataFromMediaInfo(result) {
  const tracks = resultTracks(result);
  const general = findTrack(tracks, "General");
  const video = findTrack(tracks, "Video");
  const numerator = trackValue(video, ["FrameRate_Num", "FrameRateNumerator"]);
  const denominator = trackValue(video, ["FrameRate_Den", "FrameRateDenominator"]);
  const fps = numerator && denominator
    ? Number(numerator) / Number(denominator)
    : parseRate(trackValue(video, ["FrameRate", "FrameRate_Original", "FrameRate_Nominal"]));
  const startTimecode = [
    trackValue(video, ["TimeCode_FirstFrame", "TimeCode_FirstFrame_Original", "TimeCode_Start"]),
    trackValue(general, ["TimeCode_FirstFrame", "TimeCode", "TimeCode_Start"]),
    trackValue(video, ["TimeCode"]),
  ].map(extractTimecode).find(Boolean) || "";
  const width = Number(trackValue(video, ["Width"])) || null;
  const height = Number(trackValue(video, ["Height"])) || null;
  const duration = parseRate(trackValue(general, ["Duration"])) || null;

  if (!fps) throw new Error("The reference cut opened, but its frame rate was not found in the media metadata.");
  if (!startTimecode) throw new Error("The reference cut opened, but its embedded start timecode was not found. The visible BITC is used in the screenshot, but the timing offset still needs a readable start timecode.");
  if (!isTimecode(startTimecode)) throw new Error(`The detected start timecode is not a complete timecode: ${startTimecode}`);

  return {
    fps,
    fpsLabel: displayFrameRate(fps),
    startTimecode,
    width,
    height,
    duration,
  };
}

async function detectVideoMetadata(file) {
  const factory = typeof window.MediaInfo === "function"
    ? window.MediaInfo
    : typeof window.mediaInfoFactory === "function"
      ? window.mediaInfoFactory
      : null;
  if (!factory) throw new Error("The browser media metadata library did not load. Refresh the page and try again.");

  setStatus("Reading reference cut metadata…", "working", 4);
  $("videoMetadataStatus").textContent = "Reading embedded start timecode and frame rate…";
  $("videoMetadataStatus").className = "metadata-status";

  let mediaInfo;
  try {
    mediaInfo = await factory({ format: "object", full: true });
    const readChunk = async (chunkSize, offset) =>
      new Uint8Array(await file.slice(offset, offset + chunkSize).arrayBuffer());
    const result = await mediaInfo.analyzeData(file.size, readChunk);
    const metadata = metadataFromMediaInfo(result);
    state.videoMetadata = metadata;
    $("videoMetadataStatus").textContent = `Detected automatically — Start TC: ${metadata.startTimecode} · Frame rate: ${metadata.fpsLabel} fps${metadata.width && metadata.height ? ` · ${metadata.width}×${metadata.height}` : ""}`;
    $("videoMetadataStatus").className = "metadata-status success";
    renderMidpointPreview();
    return metadata;
  } finally {
    if (mediaInfo?.close) mediaInfo.close();
  }
}

function updateHeaderRows(sheet) {
  sheet.getCell("A1").value = `${$("showTitle").value.trim() || "SHOW TITLE"} — ${$("episodeTitle").value.trim() || "EPISODE"}`;
  sheet.getCell("A2").value = "Fair Use Spreadsheet";
  sheet.getCell("A3").value = $("companyLlc").value.trim();
  sheet.getCell("A4").value = `NETWORK: ${$("network").value.trim()}`;
  sheet.getCell("A5").value = `Shift Link: ${$("referenceLink").value.trim()}`;
}

function renderMidpointPreview() {
  const status = $("midpointStatus");
  const preview = $("midpointPreview");
  if (!state.rows.length) {
    status.classList.add("hidden");
    preview.classList.add("hidden");
    return;
  }

  if (!state.videoMetadata) {
    status.textContent = `${state.rows.length} row${state.rows.length === 1 ? "" : "s"} found. Select a reference cut to calculate exact frame midpoints.`;
    status.className = "metadata-status";
    status.classList.remove("hidden");
    preview.classList.add("hidden");
    return;
  }

  try {
    state.midpoints = state.rows.map((item) => calculateMidpoint(item, state.videoMetadata.fps));
  } catch (error) {
    status.textContent = error.message;
    status.className = "metadata-status";
    status.classList.remove("hidden");
    preview.classList.add("hidden");
    return;
  }

  status.textContent = `${state.midpoints.length} midpoint${state.midpoints.length === 1 ? "" : "s"} calculated at ${state.videoMetadata.fpsLabel} fps. These are the frame timecodes the reference cut will be asked to show.`;
  status.className = "metadata-status success";
  status.classList.remove("hidden");

  preview.innerHTML = "";
  const heading = document.createElement("div");
  heading.className = "midpoint-heading";
  heading.textContent = "Midpoint check";
  preview.appendChild(heading);

  const table = document.createElement("table");
  table.className = "midpoint-table";
  const header = document.createElement("thead");
  const headerRow = document.createElement("tr");
  ["Row", "File Name", "TC IN", "TC OUT", "Midpoint"].forEach((label) => {
    const cell = document.createElement("th");
    cell.textContent = label;
    headerRow.appendChild(cell);
  });
  header.appendChild(headerRow);
  table.appendChild(header);
  const body = document.createElement("tbody");
  const visibleRows = state.midpoints.slice(0, 100);
  visibleRows.forEach((item) => {
    const row = document.createElement("tr");
    [item.row, item.fileName || "(no file name)", item.tcIn, item.tcOut, item.midpoint].forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = String(value);
      row.appendChild(cell);
    });
    body.appendChild(row);
  });
  table.appendChild(body);
  preview.appendChild(table);
  if (state.midpoints.length > visibleRows.length) {
    const note = document.createElement("div");
    note.className = "midpoint-note";
    note.textContent = `Showing the first ${visibleRows.length} rows. All ${state.midpoints.length} rows will be processed.`;
    preview.appendChild(note);
  }
  preview.classList.remove("hidden");
}

async function loadLocalWorkbook(file) {
  if (!window.ExcelJS?.Workbook) throw new Error("The Excel workbook library did not load. Refresh the page and try again.");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const worksheet = findFairUseSheet(workbook);
  const rows = getWorkbookRows(worksheet);
  if (!rows.length) throw new Error("No rows were found with both TC IN and TC OUT timecodes in columns C and D.");
  return { workbook, worksheet, rows };
}

function revokeVideo() {
  if (state.video) {
    state.video.pause();
    state.video.remove();
    state.video = null;
  }
  if (state.videoUrl) {
    URL.revokeObjectURL(state.videoUrl);
    state.videoUrl = null;
  }
}

function waitForVideoMetadata(video) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timeout = setTimeout(() => finish(new Error("The browser could not read the reference cut's video stream.")), 60000);
    const cleanup = () => {
      clearTimeout(timeout);
      video.removeEventListener("loadedmetadata", onLoaded);
      video.removeEventListener("error", onError);
    };
    const finish = (error) => {
      if (settled) return;
      settled = true;
      cleanup();
      error ? reject(error) : resolve();
    };
    const onLoaded = () => {
      if (!video.videoWidth || !video.videoHeight || !Number.isFinite(video.duration)) {
        finish(new Error("The browser opened the file but could not decode its video frames. Try an H.264 MP4/MOV reference cut for this browser test."));
        return;
      }
      finish();
    };
    const onError = () => {
      const code = video.error?.code;
      const detail = code === 3 ? "The video codec could not be decoded by this browser." : "The browser could not open this video file.";
      finish(new Error(`${detail} Try an H.264 MP4/MOV reference cut for this browser test.`));
    };
    video.addEventListener("loadedmetadata", onLoaded, { once: true });
    video.addEventListener("error", onError, { once: true });
  });
}

async function loadReferenceVideo(file) {
  revokeVideo();
  const video = document.createElement("video");
  video.preload = "auto";
  video.muted = true;
  video.playsInline = true;
  video.setAttribute("playsinline", "");
  video.style.display = "none";
  state.videoUrl = URL.createObjectURL(file);
  video.src = state.videoUrl;
  document.body.appendChild(video);
  await waitForVideoMetadata(video);
  state.video = video;
  return video;
}

function waitForFrame(video) {
  if (typeof video.requestVideoFrameCallback === "function") {
    return new Promise((resolve) => {
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        resolve();
      };
      video.requestVideoFrameCallback(finish);
      setTimeout(finish, 250);
    });
  }
  return new Promise((resolve) => setTimeout(resolve, 80));
}

async function seekVideo(video, seconds) {
  const target = Math.max(0, Math.min(seconds, Math.max(0, video.duration - 0.001)));
  if (Math.abs(video.currentTime - target) > 0.001) {
    await new Promise((resolve, reject) => {
      let settled = false;
      const timeout = setTimeout(() => finish(new Error(`The browser could not seek to ${target.toFixed(3)} seconds.`)), 60000);
      const cleanup = () => {
        clearTimeout(timeout);
        video.removeEventListener("seeked", onSeeked);
        video.removeEventListener("error", onError);
      };
      const finish = (error) => {
        if (settled) return;
        settled = true;
        cleanup();
        error ? reject(error) : resolve();
      };
      const onSeeked = () => finish();
      const onError = () => finish(new Error("The browser could not decode the requested reference-cut frame."));
      video.addEventListener("seeked", onSeeked, { once: true });
      video.addEventListener("error", onError, { once: true });
      try {
        video.currentTime = target;
      } catch (error) {
        finish(error);
      }
    });
  }
  await waitForFrame(video);
}

function canvasBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => blob ? resolve(blob) : reject(new Error("The browser could not create a screenshot.")), "image/jpeg", 0.7);
  });
}

async function captureFrame(video, seconds) {
  await seekVideo(video, seconds);
  const canvas = document.createElement("canvas");
  canvas.width = THUMBNAIL_WIDTH;
  canvas.height = THUMBNAIL_HEIGHT;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("The browser could not create a screenshot canvas.");
  context.fillStyle = "#000";
  context.fillRect(0, 0, canvas.width, canvas.height);
  const scale = Math.min(canvas.width / video.videoWidth, canvas.height / video.videoHeight);
  const width = Math.round(video.videoWidth * scale);
  const height = Math.round(video.videoHeight * scale);
  const left = Math.round((canvas.width - width) / 2);
  const top = Math.round((canvas.height - height) / 2);
  context.drawImage(video, left, top, width, height);
  const blob = await canvasBlob(canvas);
  return new Uint8Array(await blob.arrayBuffer());
}

function addScreenshotToExcel(sheet, rowNumber, bytes) {
  const imageId = state.workbook.addImage({ buffer: bytes.buffer, extension: "jpeg" });
  sheet.addImage(imageId, {
    tl: { col: 0, row: rowNumber - 1 },
    ext: { width: THUMBNAIL_WIDTH, height: THUMBNAIL_HEIGHT },
  });
  sheet.getRow(rowNumber).height = 140;
  sheet.getColumn(1).width = 46;
}

function appendThumbnailCard(item, bytes) {
  const card = document.createElement("figure");
  card.className = "thumbnail-card";
  const image = document.createElement("img");
  const url = URL.createObjectURL(new Blob([bytes], { type: "image/jpeg" }));
  state.previewUrls.push(url);
  image.src = url;
  image.alt = `Row ${item.row} screenshot at ${item.midpoint}`;
  const caption = document.createElement("figcaption");
  caption.textContent = `Row ${item.row} · ${item.midpoint}`;
  card.append(image, caption);
  $("thumbnailGrid").append(card);
}

async function createLog() {
  clearError();
  const logFile = $("logFile").files[0];
  const referenceFile = $("referenceCut").files[0];
  if (!logFile) throw new Error("Please choose a partially completed Fair Use Log Excel file.");
  if (!referenceFile) throw new Error("Please choose the reference cut.");

  if (!state.workbook || state.referenceFile !== logFile) {
    setStatus("Reading Fair Use Log…", "working", 8);
    const loaded = await loadLocalWorkbook(logFile);
    state.workbook = loaded.workbook;
    state.worksheet = loaded.worksheet;
    state.rows = loaded.rows;
  }
  if (!state.videoMetadata || state.referenceFile !== referenceFile) {
    state.referenceFile = referenceFile;
    await detectVideoMetadata(referenceFile);
  }

  state.midpoints = state.rows.map((item) => calculateMidpoint(item, state.videoMetadata.fps));
  renderMidpointPreview();
  const startFrames = parseTimecode(state.videoMetadata.startTimecode, state.videoMetadata.fps);
  const video = await loadReferenceVideo(referenceFile);
  const durationSeconds = video.duration;
  state.images = [];
  state.outputBuffer = null;
  $("resultsCard").classList.add("hidden");
  $("thumbnailGrid").innerHTML = "";

  updateHeaderRows(state.worksheet);
  for (let index = 0; index < state.midpoints.length; index += 1) {
    const item = state.midpoints[index];
    const offsetFrames = item.midpointFrames - startFrames;
    if (offsetFrames < 0) {
      throw new Error(`Row ${item.row}: midpoint ${item.midpoint} occurs before the reference cut start timecode ${state.videoMetadata.startTimecode}.`);
    }
    const seconds = offsetFrames / state.videoMetadata.fps;
    if (seconds > durationSeconds + 0.05) {
      throw new Error(`Row ${item.row}: midpoint ${item.midpoint} is beyond the reference cut duration.`);
    }
    setStatus(`Capturing row ${item.row} at ${item.midpoint}…`, "working", 10 + (index / state.midpoints.length) * 85);
    const bytes = await captureFrame(video, seconds);
    addScreenshotToExcel(state.worksheet, item.row, bytes);
    state.images.push({ row: item.row, midpoint: item.midpoint, bytes });
    appendThumbnailCard(item, bytes);
  }

  state.outputBuffer = await state.workbook.xlsx.writeBuffer();
  $("resultTitle").textContent = "Screenshots ready";
  $("resultSummary").textContent = `${state.images.length} midpoint screenshot${state.images.length === 1 ? "" : "s"} generated locally from ${referenceFile.name}.`;
  $("resultsCard").classList.remove("hidden");
  setStatus("Complete", "complete", 100);
}

async function downloadImages() {
  if (!state.images.length) return;
  const zip = new JSZip();
  state.images.forEach((image) => {
    zip.file(`Row${String(image.row).padStart(4, "0")}_${image.midpoint.replace(/[:;]/g, "-")}.jpg`, image.bytes);
  });
  downloadBlob(await zip.generateAsync({ type: "blob" }), "Fair_Use_Screenshots.zip");
}

async function handleLogFileChange() {
  const file = $("logFile").files[0];
  const token = ++state.logReadToken;
  state.workbook = null;
  state.worksheet = null;
  state.rows = [];
  state.midpoints = [];
  $("midpointPreview").classList.add("hidden");
  if (!file) {
    $("selectedLogStatus").textContent = "No Fair Use Log selected.";
    $("midpointStatus").classList.add("hidden");
    return;
  }
  try {
    setStatus("Reading Fair Use Log…", "working", 2);
    $("selectedLogStatus").textContent = `Reading ${file.name}…`;
    const loaded = await loadLocalWorkbook(file);
    if (token !== state.logReadToken) return;
    state.workbook = loaded.workbook;
    state.worksheet = loaded.worksheet;
    state.rows = loaded.rows;
    $("selectedLogStatus").textContent = `Selected ${file.name} · ${state.rows.length} row${state.rows.length === 1 ? "" : "s"} with TC IN and TC OUT in columns C and D.`;
    renderMidpointPreview();
    setStatus("Fair Use Log ready", "working", 5);
  } catch (error) {
    showError(error);
  }
}

async function handleReferenceChange() {
  const file = $("referenceCut").files[0] || null;
  state.referenceFile = file;
  state.videoMetadata = null;
  revokeVideo();
  if (!file) {
    $("videoMetadataStatus").textContent = "Choose a reference cut. Its embedded start timecode and frame rate will be detected automatically.";
    $("videoMetadataStatus").className = "metadata-status";
    renderMidpointPreview();
    return;
  }
  try {
    await detectVideoMetadata(file);
    setStatus("Reference cut metadata ready", "working", 7);
  } catch (error) {
    $("videoMetadataStatus").textContent = error.message;
    $("videoMetadataStatus").className = "metadata-status";
    showError(error);
  }
}

$("logFile").addEventListener("change", () => { void handleLogFileChange(); });
$("referenceCut").addEventListener("change", () => { void handleReferenceChange(); });
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
  downloadBlob(
    new Blob([state.outputBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "FAIR_USE_GENERATED.xlsx"
  );
});
$("downloadImagesButton").addEventListener("click", () => { void downloadImages(); });
$("cancelButton").addEventListener("click", () => window.location.reload());
