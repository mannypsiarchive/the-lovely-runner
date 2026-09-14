/*
  Fair Use Log 3.3

  The browser keeps the Excel workbook and reference cut local. FFmpeg.wasm is used
  for the same metadata probe that the Python version uses; the browser's native
  video element then seeks the exact offset and Canvas captures the visible BITC.
  The first workbook is a review copy. The final workbook is not built until the
  user has reviewed it, closed Excel, and either entered or declined replacements.
*/

const FAIR_USE_LOG_VERSION = "3.3";
const FFMPEG_VERSION = "0.12.10";
const FFMPEG_SCRIPT_URL = `https://cdn.jsdelivr.net/npm/@ffmpeg/ffmpeg@${FFMPEG_VERSION}/dist/umd/ffmpeg.js`;
const FFMPEG_CORE_BASE_URL = `https://cdn.jsdelivr.net/npm/@ffmpeg/core@${FFMPEG_VERSION}/dist/umd`;
const FFMPEG_WORKER_URL = `https://cdn.jsdelivr.net/npm/@ffmpeg/ffmpeg@${FFMPEG_VERSION}/dist/umd/814.ffmpeg.js`;

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
  logFile: null,
  rows: [],
  midpoints: [],
  referenceFile: null,
  videoMetadata: null,
  video: null,
  videoUrl: null,
  outputBuffer: null,
  reviewBuffer: null,
  finalBuffer: null,
  dueDiligenceZip: null,
  images: [],
  reviewChanges: [],
  sourceWorkbookBuffer: null,
  previewUrls: [],
  logReadToken: 0,
  ffmpeg: null,
  ffmpegClassWorkerUrl: null,
  ffmpegLogLines: [],
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

function activityTime() {
  return new Date().toLocaleTimeString([], { hour12: false });
}

function logActivity(message, kind = "info") {
  const log = $("activityLog");
  const entry = document.createElement("div");
  entry.className = `activity-entry ${kind}`;
  const time = document.createElement("span");
  time.className = "activity-time";
  time.textContent = activityTime();
  const text = document.createElement("span");
  text.textContent = message;
  entry.append(time, text);
  if (log) {
    log.appendChild(entry);
    while (log.children.length > 150) log.firstElementChild.remove();
    log.scrollTop = log.scrollHeight;
  }
  console.log(`[Fair Use Log] ${message}`);
}

function yieldToBrowser() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

async function activity(message, kind = "info") {
  logActivity(message, kind);
  await yieldToBrowser();
}

function clearActivityLog() {
  const log = $("activityLog");
  if (log) log.innerHTML = "";
}

function showError(error) {
  const message = error instanceof Error ? error.message : String(error || "Unknown error.");
  logActivity(message, "error");
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

function allTrackValues(tracks, namePattern) {
  const values = [];
  tracks.forEach((track) => {
    Object.entries(track || {}).forEach(([key, value]) => {
      if (namePattern.test(key) && cellText(value)) values.push(value);
    });
  });
  return values;
}

function canonicalFrameRate(fps) {
  if (!Number.isFinite(fps)) return null;
  if (Math.abs(fps - 23.98) < 0.02 || Math.abs(fps - 23.976) < 0.02) return 23.976;
  if (Math.abs(fps - 29.97) < 0.03 || Math.abs(fps - 29.98) < 0.03) return 29.97;
  if (Math.abs(fps - 59.94) < 0.03 || Math.abs(fps - 59.96) < 0.03) return 59.94;
  return fps;
}

function makeVideoMetadata({ fps, startTimecode, width = null, height = null, duration = null, source = "metadata" }) {
  const normalizedFps = canonicalFrameRate(fps);
  const normalizedStart = extractTimecode(startTimecode);
  if (!normalizedFps) throw new Error("The reference cut opened, but its frame rate was not found in the media metadata.");
  if (!normalizedStart) throw new Error("The reference cut opened, but its embedded start timecode was not found. The visible BITC is used in the screenshot, but the timing offset still needs a readable start timecode.");
  if (!isTimecode(normalizedStart)) throw new Error(`The detected start timecode is not a complete timecode: ${normalizedStart}`);
  return {
    fps: normalizedFps,
    fpsLabel: displayFrameRate(normalizedFps),
    startTimecode: normalizedStart,
    width: Number(width) || null,
    height: Number(height) || null,
    duration: Number(duration) || null,
    source,
  };
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
    ...allTrackValues(tracks, /time.?code|time.?code.?first|start.?time/i),
  ].map(extractTimecode).find(Boolean) || "";
  const width = Number(trackValue(video, ["Width"])) || null;
  const height = Number(trackValue(video, ["Height"])) || null;
  const duration = parseRate(trackValue(general, ["Duration"])) || null;
  return makeVideoMetadata({ fps, startTimecode, width, height, duration, source: "MediaInfo" });
}

function mediaInfoFactory() {
  if (typeof window.MediaInfo === "function") return window.MediaInfo;
  if (typeof window.MediaInfo?.mediaInfoFactory === "function") return window.MediaInfo.mediaInfoFactory;
  if (typeof window.MediaInfo?.default === "function") return window.MediaInfo.default;
  if (typeof window.mediaInfoFactory === "function") return window.mediaInfoFactory;
  return null;
}

async function detectWithMediaInfo(file) {
  const factory = mediaInfoFactory();
  if (!factory) throw new Error("The browser media metadata library did not load. Refresh the page and try again.");

  setStatus("Reading reference cut metadata…", "working", 4);
  $("videoMetadataStatus").textContent = "Reading embedded start timecode and frame rate…";
  $("videoMetadataStatus").className = "metadata-status";

  let mediaInfo;
  try {
    await activity(`Reading reference cut metadata: ${file.name}`);
    mediaInfo = await factory({ format: "object", full: true });
    await activity("Metadata reader loaded. Scanning the media file for frame rate and start timecode.");
    const readChunk = async (chunkSize, offset) =>
      new Uint8Array(await file.slice(offset, offset + chunkSize).arrayBuffer());
    const result = await mediaInfo.analyzeData(file.size, readChunk);
    await activity("Media metadata scan finished. Checking the video track.");
    const metadata = metadataFromMediaInfo(result);
    state.videoMetadata = metadata;
    $("videoMetadataStatus").textContent = `Detected automatically — Start TC: ${metadata.startTimecode} · Frame rate: ${metadata.fpsLabel} fps${metadata.width && metadata.height ? ` · ${metadata.width}×${metadata.height}` : ""}`;
    $("videoMetadataStatus").className = "metadata-status success";
    await activity(`Detected start TC ${metadata.startTimecode} and ${metadata.fpsLabel} fps.`, "success");
    renderMidpointPreview();
    return metadata;
  } finally {
    if (mediaInfo?.close) mediaInfo.close();
  }
}

function ffmpegClass() {
  return window.FFmpegWASM?.FFmpeg || window.FFmpeg?.FFmpeg || (typeof window.FFmpeg === "function" ? window.FFmpeg : null);
}

function loadExternalScript(url) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector('script[data-fair-use-script="' + url + '"]');
    if (existing) {
      if (existing.dataset.loaded === "true") resolve();
      else existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", () => reject(new Error("Could not load " + url + ".")), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = url;
    script.async = true;
    script.dataset.fairUseScript = url;
    script.addEventListener("load", () => { script.dataset.loaded = "true"; resolve(); }, { once: true });
    script.addEventListener("error", () => reject(new Error("Could not load " + url + ".")), { once: true });
    document.head.appendChild(script);
  });
}

async function toBlobURL(url, mimeType) {
  const response = await fetch(url);
  if (!response.ok) throw new Error("Could not download " + url + " (" + response.status + ").");
  const blob = await response.blob();
  return URL.createObjectURL(new Blob([blob], { type: mimeType }));
}

function recordFfmpegLog(message) {
  const text = cellText(message);
  if (!text) return;
  state.ffmpegLogLines.push(text);
  while (state.ffmpegLogLines.length > 120) state.ffmpegLogLines.shift();
  if (/timecode|duration:|stream #|input #|video:/i.test(text)) logActivity("FFmpeg: " + text);
}

async function loadFfmpeg() {
  if (state.ffmpeg) return state.ffmpeg;
  await activity("Loading FFmpeg " + FFMPEG_VERSION + " locally for the metadata probe. This may take a moment.");
  if (!ffmpegClass()) await loadExternalScript(FFMPEG_SCRIPT_URL);
  const FFmpeg = ffmpegClass();
  if (!FFmpeg) throw new Error("The FFmpeg browser library did not load. Check the browser's network access and refresh the page.");

  let coreURL;
  let wasmURL;
  let classWorkerURL;
  try {
    coreURL = await toBlobURL(FFMPEG_CORE_BASE_URL + "/ffmpeg-core.js", "text/javascript");
    wasmURL = await toBlobURL(FFMPEG_CORE_BASE_URL + "/ffmpeg-core.wasm", "application/wasm");
    classWorkerURL = await toBlobURL(FFMPEG_WORKER_URL, "text/javascript");
    const ffmpeg = new FFmpeg();
    if (typeof ffmpeg.on === "function") ffmpeg.on("log", ({ message }) => recordFfmpegLog(message));
    await ffmpeg.load({ coreURL, wasmURL, classWorkerURL });
    state.ffmpeg = ffmpeg;
    state.ffmpegClassWorkerUrl = classWorkerURL;
    await activity("FFmpeg loaded. It will read the reference cut metadata without uploading the video.", "success");
    return ffmpeg;
  } catch (error) {
    [coreURL, wasmURL, classWorkerURL].filter(Boolean).forEach((url) => URL.revokeObjectURL(url));
    throw new Error("FFmpeg could not start in this browser: " + (error.message || error));
  }
}

function secondsFromDuration(value) {
  const text = cellText(value);
  const match = text.match(/(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?/);
  if (!match) return null;
  return Number(match[1]) * 3600 + Number(match[2]) * 60 + Number(match[3]) + Number("0." + (match[4] || "0"));
}

function metadataFromFfmpegProbe(probe) {
  const streams = Array.isArray(probe?.streams) ? probe.streams : [];
  const video = streams.find((stream) => cellText(stream?.codec_type).toLowerCase() === "video") || streams[0] || {};
  const formatTags = probe?.format?.tags || {};
  const streamTags = video?.tags || {};
  const tagValues = [...Object.entries(streamTags), ...Object.entries(formatTags)];
  const tagTimecode = tagValues
    .filter(([key]) => /time.?code|start.?time/i.test(key))
    .map(([, value]) => extractTimecode(value))
    .find(Boolean) || "";
  const fps = parseRate(video.avg_frame_rate) || parseRate(video.r_frame_rate);
  return makeVideoMetadata({
    fps,
    startTimecode: tagTimecode,
    width: video.width,
    height: video.height,
    duration: probe?.format?.duration,
    source: "FFmpeg",
  });
}

function metadataFromFfmpegLog(lines) {
  const report = lines.join("\n");
  const firstTimecode = report.match(/time.?code\s*[:=]\s*([^\s,]+)/i);
  const secondTimecode = report.match(/time.?code[^\n]*?(\d{2}:\d{2}:\d{2}[:;]\d{2})/i);
  const startTimecode = extractTimecode(firstTimecode?.[1] || "") || extractTimecode(secondTimecode?.[1] || "");
  const fpsMatch = report.match(/(\d+(?:\.\d+)?)\s+(?:fps|tbr)\b/i);
  const resolution = report.match(/\b(\d{3,5})x(\d{3,5})\b/);
  const durationMatch = report.match(/Duration:\s*([^,\s]+)/i);
  return makeVideoMetadata({
    fps: parseRate(fpsMatch?.[1]),
    startTimecode,
    width: resolution?.[1],
    height: resolution?.[2],
    duration: secondsFromDuration(durationMatch?.[1]),
    source: "FFmpeg report",
  });
}

async function detectWithFfmpeg(file) {
  setStatus("Reading reference cut metadata…", "working", 4);
  $("videoMetadataStatus").textContent = "FFmpeg is reading the embedded start timecode and frame rate…";
  $("videoMetadataStatus").className = "metadata-status";
  const ffmpeg = await loadFfmpeg();
  const inputName = "reference-" + Date.now() + "-" + file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
  const outputName = "fair-use-metadata.json";
  state.ffmpegLogLines = [];
  await activity("Writing " + file.name + " into FFmpeg's private browser memory.");
  await ffmpeg.writeFile(inputName, new Uint8Array(await file.arrayBuffer()));

  let metadata = null;
  if (typeof ffmpeg.ffprobe === "function") {
    try {
      await activity("Running the FFprobe metadata query used by the Python version.");
      await ffmpeg.ffprobe([
        "-v", "error",
        "-select_streams", "v:0",
        "-show_entries", "stream=avg_frame_rate,r_frame_rate,width,height:stream_tags=timecode:format=duration:format_tags=timecode",
        "-of", "json",
        inputName,
        "-o", outputName,
      ]);
      const raw = await ffmpeg.readFile(outputName);
      metadata = metadataFromFfmpegProbe(JSON.parse(new TextDecoder().decode(raw)));
    } catch (error) {
      await activity("FFprobe JSON did not produce usable metadata: " + (error.message || error), "warning");
    }
  } else {
    await activity("This FFmpeg browser build has no FFprobe method; using its media report instead.", "warning");
  }

  if (!metadata) {
    state.ffmpegLogLines = [];
    await activity("Reading FFmpeg's media report for the embedded start timecode and frame rate.");
    try { await ffmpeg.exec(["-hide_banner", "-i", inputName]); } catch { /* input metadata is logged before the no-output error */ }
    metadata = metadataFromFfmpegLog(state.ffmpegLogLines);
  }

  state.videoMetadata = metadata;
  $("videoMetadataStatus").textContent = "Detected automatically — Start TC: " + metadata.startTimecode + " · Frame rate: " + metadata.fpsLabel + " fps" +
    (metadata.width && metadata.height ? " · " + metadata.width + "×" + metadata.height : "") + " · " + metadata.source;
  $("videoMetadataStatus").className = "metadata-status success";
  await activity("Detected start TC " + metadata.startTimecode + " and " + metadata.fpsLabel + " fps with " + metadata.source + ".", "success");
  renderMidpointPreview();
  return metadata;
}

async function detectVideoMetadata(file) {
  try {
    return await detectWithFfmpeg(file);
  } catch (ffmpegError) {
    await activity("FFmpeg metadata path failed: " + (ffmpegError.message || ffmpegError), "warning");
    try {
      await activity("Trying the browser metadata reader as a fallback.");
      return await detectWithMediaInfo(file);
    } catch (mediaInfoError) {
      throw new Error("FFmpeg could not read this reference cut (" + (ffmpegError.message || ffmpegError) + "). MediaInfo fallback also failed (" + (mediaInfoError.message || mediaInfoError) + ").");
    }
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
  await activity(`Opening Excel workbook: ${file.name}`);
  const workbook = new ExcelJS.Workbook();
  const sourceBuffer = await file.arrayBuffer();
  await workbook.xlsx.load(sourceBuffer);
  await activity(`Workbook opened. Tabs found: ${workbook.worksheets.length}.`);
  const worksheet = findFairUseSheet(workbook);
  await activity(`Using worksheet: ${worksheet.name}.`);
  const rows = getWorkbookRows(worksheet);
  if (!rows.length) throw new Error("No rows were found with both TC IN and TC OUT timecodes in columns C and D.");
  await activity(`Found ${rows.length} row${rows.length === 1 ? "" : "s"} with TC IN and TC OUT.`);
  return { workbook, worksheet, rows, sourceBuffer };
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
  await activity(`Opening reference cut in the browser video player: ${file.name}`);
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
  await activity(`Reference cut decoded: ${video.videoWidth}×${video.videoHeight}, ${video.duration.toFixed(2)} seconds.`, "success");
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

function addScreenshotToExcel(workbook, sheet, rowNumber, bytes) {
  const imageBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const imageId = workbook.addImage({ buffer: imageBuffer, extension: "jpeg" });
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

function clearPreviewUrls() {
  state.previewUrls.forEach((url) => URL.revokeObjectURL(url));
  state.previewUrls = [];
}

function refreshThumbnailGrid() {
  clearPreviewUrls();
  $("thumbnailGrid").innerHTML = "";
  state.images.forEach((image) => {
    const item = state.midpoints.find((candidate) => candidate.row === image.row) || image;
    appendThumbnailCard({ ...item, midpoint: image.midpoint }, image.bytes);
  });
}

async function buildWorkbookBuffer() {
  if (!state.sourceWorkbookBuffer) throw new Error("The original Excel workbook is no longer available in the browser.");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(state.sourceWorkbookBuffer);
  const worksheet = findFairUseSheet(workbook);
  updateHeaderRows(worksheet);
  state.images.forEach((image) => addScreenshotToExcel(workbook, worksheet, image.row, image.bytes));
  state.workbook = workbook;
  state.worksheet = worksheet;
  return workbook.xlsx.writeBuffer();
}

function renderReviewChanges() {
  const list = $("reviewChangeList");
  const status = $("reviewStatus");
  list.innerHTML = "";
  if (!state.reviewChanges.length) {
    status.textContent = "No replacement rows entered. If the review workbook looks correct, continue without changes.";
    return;
  }
  status.textContent = state.reviewChanges.length + " replacement row" + (state.reviewChanges.length === 1 ? "" : "s") + " queued.";
  state.reviewChanges.forEach((change) => {
    const item = document.createElement("div");
    item.className = "review-change-item";
    const text = document.createElement("span");
    text.textContent = "Row " + change.row + " → " + change.midpoint;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "button danger";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      state.reviewChanges = state.reviewChanges.filter((candidate) => candidate.row !== change.row);
      renderReviewChanges();
    });
    item.append(text, remove);
    list.appendChild(item);
  });
}

function addReviewChange() {
  const row = Number.parseInt($("reviewRow").value.trim(), 10);
  const midpoint = normalizeTimecode($("reviewTc").value);
  if (!Number.isInteger(row) || row < DATA_START_ROW) throw new Error("Enter a valid worksheet row number.");
  if (!state.midpoints.some((item) => item.row === row)) throw new Error("Row " + row + " was not one of the rows processed.");
  if (!isTimecode(midpoint)) throw new Error("Enter the replacement as a complete timecode, for example 01:33:00:00.");
  const midpointFrames = parseTimecode(midpoint, state.videoMetadata.fps);
  state.reviewChanges = state.reviewChanges.filter((change) => change.row !== row);
  state.reviewChanges.push({ row, midpoint, midpointFrames });
  state.reviewChanges.sort((a, b) => a.row - b.row);
  $("reviewRow").value = "";
  $("reviewTc").value = "";
  renderReviewChanges();
  logActivity("Queued replacement screenshot for row " + row + " at " + midpoint + ".", "success");
}

function safeFilename(value) {
  const cleaned = cellText(value).replace(/[<>:"/\\|?*]+/g, "-").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
  return (cleaned || "Unknown Source").slice(0, 100);
}

function docxLibrary() {
  return window.docx || window.Docx || null;
}

async function prepareDueDiligence() {
  const library = docxLibrary();
  if (!library?.Document || !library?.Packer || !library?.ImageRun) {
    throw new Error("The browser Word document library did not load. Refresh the page and try again.");
  }
  const groups = new Map();
  state.rows.forEach((row) => {
    if (!row.source) return;
    const image = state.images.find((candidate) => candidate.row === row.row);
    if (!image) return;
    if (!groups.has(row.source)) groups.set(row.source, []);
    groups.get(row.source).push({ row, image });
  });
  if (!groups.size) throw new Error("No screenshots have a Source value in column G, so no Due Diligence logs can be prepared.");

  const zip = new JSZip();
  const {
    Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType,
  } = library;
  await activity("Preparing Due Diligence logs from the final workbook.");
  let formCount = 0;
  for (const [source, items] of groups.entries()) {
    const children = [];
    const centered = (text, size = 28) => new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: String(text || ""), bold: true, size })],
    });
    const body = (text, bold = false) => new Paragraph({
      children: [new TextRun({ text: String(text || ""), bold })],
    });
    children.push(centered(source + " DUE DILIGENCE OUTREACH FORM"));
    children.push(body(""));
    children.push(centered("SERIES TITLE: " + ($("showTitle").value.trim() || "SHOW TITLE")));
    children.push(centered("EPISODE TITLE: " + ($("episodeTitle").value.trim() || "EPISODE")));
    children.push(centered("COMPANY INFO: " + $("companyLlc").value.trim()));
    children.push(body(""));
    children.push(centered("MATERIAL TO FAIR USE:"));
    for (const item of items) {
      children.push(new Paragraph({
        children: [new ImageRun({
          data: item.image.bytes,
          type: "jpg",
          transformation: { width: 556, height: 313 },
        })],
      }));
      children.push(body("Time Codes Used: " + item.row.tcIn + " - " + item.row.tcOut, true));
      children.push(body(""));
    }
    children.push(centered("DESCRIPTION:"));
    items.forEach((item) => {
      children.push(body(item.row.tcIn + " - " + item.row.tcOut + ": " + (item.row.description || "")));
    });
    children.push(body(""));
    children.push(centered("DUE DILIGENCE:"));
    for (let blank = 0; blank < 6; blank += 1) children.push(body(""));

    const document = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(document);
    zip.file(safeFilename(source) + " - DD Form.docx", await blob.arrayBuffer());
    formCount += 1;
    await activity("Prepared Due Diligence form for " + source + ".", "success");
  }
  state.dueDiligenceZip = await zip.generateAsync({ type: "blob" });
  $("dueDiligenceSummary").textContent = "Prepared " + formCount + " Due Diligence Word form" + (formCount === 1 ? "" : "s") + " from the final screenshots.";
  $("downloadDueDiligenceButton").classList.remove("hidden");
  logActivity("Due Diligence preparation finished. " + formCount + " Word form" + (formCount === 1 ? "" : "s") + " are ready.", "success");
}

async function applyReviewChanges() {
  const button = $("applyReviewButton");
  if (!$("reviewWorkbookClosed").checked) throw new Error("Close the review workbook in Excel, then confirm that it is closed.");
  button.disabled = true;
  const originalLabel = button.textContent;
  button.textContent = "Applying…";
  try {
    clearError();
    await activity("Review workbook closed. Applying the requested row changes.");
    const startFrames = parseTimecode(state.videoMetadata.startTimecode, state.videoMetadata.fps);
    for (let index = 0; index < state.reviewChanges.length; index += 1) {
      const change = state.reviewChanges[index];
      const rowItem = state.midpoints.find((item) => item.row === change.row);
      const offsetFrames = change.midpointFrames - startFrames;
      if (offsetFrames < 0) throw new Error("Row " + change.row + ": replacement midpoint " + change.midpoint + " is before the reference cut start timecode.");
      const seconds = offsetFrames / state.videoMetadata.fps;
      if (seconds > state.video.duration + 0.05) throw new Error("Row " + change.row + ": replacement midpoint " + change.midpoint + " is beyond the reference cut duration.");
      await activity("Row " + change.row + ": generating replacement screenshot at " + change.midpoint + ".");
      setStatus("Replacing row " + change.row + "…", "working", 10 + ((index + 1) / Math.max(1, state.reviewChanges.length)) * 55);
      const bytes = await captureFrame(state.video, seconds);
      const image = state.images.find((candidate) => candidate.row === change.row);
      if (!image) throw new Error("Could not find the original screenshot for row " + change.row + ".");
      image.bytes = bytes;
      image.midpoint = change.midpoint;
      rowItem.midpoint = change.midpoint;
      rowItem.midpointFrames = change.midpointFrames;
      await activity("Row " + change.row + ": replacement screenshot captured.", "success");
    }
    refreshThumbnailGrid();
    await activity("Saving the final Excel workbook after review.");
    state.finalBuffer = await buildWorkbookBuffer();
    state.outputBuffer = state.finalBuffer;
    await activity("Final workbook saved. Screenshots are embedded in column A.", "success");
    $("reviewCard").classList.add("hidden");
    $("resultTitle").textContent = "Final Fair Use Log ready";
    $("resultSummary").textContent = state.images.length + " screenshot" + (state.images.length === 1 ? "" : "s") + " embedded in column A. The review stage is complete.";
    $("resultsCard").classList.remove("hidden");
    setStatus("Final workbook ready", "complete", 100);
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
}

async function createLog() {
  clearActivityLog();
  await activity("Create Fair Use Log clicked.");
  clearError();
  const logFile = $("logFile").files[0];
  const referenceFile = $("referenceCut").files[0];
  if (!logFile) throw new Error("Please choose a partially completed Fair Use Log Excel file.");
  if (!referenceFile) throw new Error("Please choose the reference cut.");

  await activity(`Inputs found: ${logFile.name} and ${referenceFile.name}.`);
  if (!state.workbook || state.logFile !== logFile) {
    setStatus("Reading Fair Use Log…", "working", 8);
    await activity("The workbook has not been loaded for this selection. Reading it now.");
    const loaded = await loadLocalWorkbook(logFile);
    state.workbook = loaded.workbook;
    state.worksheet = loaded.worksheet;
    state.rows = loaded.rows;
    state.sourceWorkbookBuffer = loaded.sourceBuffer;
    state.logFile = logFile;
  } else {
    await activity(`Using the already-read workbook with ${state.rows.length} row${state.rows.length === 1 ? "" : "s"}.`);
  }
  if (!state.videoMetadata || state.referenceFile !== referenceFile) {
    state.referenceFile = referenceFile;
    await activity("Reference-cut metadata is not ready for this file. Reading it now.");
    await detectVideoMetadata(referenceFile);
  } else {
    await activity(`Using detected reference-cut metadata: ${state.videoMetadata.startTimecode} at ${state.videoMetadata.fpsLabel} fps.`);
  }

  await activity(`Calculating ${state.rows.length} midpoint${state.rows.length === 1 ? "" : "s"} from columns C and D.`);
  state.midpoints = state.rows.map((item) => calculateMidpoint(item, state.videoMetadata.fps));
  renderMidpointPreview();
  await activity(`Calculated ${state.midpoints.length} midpoint${state.midpoints.length === 1 ? "" : "s"}.` , "success");
  const startFrames = parseTimecode(state.videoMetadata.startTimecode, state.videoMetadata.fps);
  const video = await loadReferenceVideo(referenceFile);
  const durationSeconds = video.duration;
  state.images = [];
  state.outputBuffer = null;
  state.reviewBuffer = null;
  state.finalBuffer = null;
  state.reviewChanges = [];
  $("reviewWorkbookClosed").checked = false;
  $("applyReviewButton").disabled = true;
  $("resultsCard").classList.add("hidden");
  $("reviewCard").classList.add("hidden");
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
    await activity(`Row ${item.row}: midpoint ${item.midpoint}; seeking ${seconds.toFixed(3)} seconds from the start of the reference cut.`);
    setStatus(`Capturing row ${item.row} at ${item.midpoint}…`, "working", 10 + (index / state.midpoints.length) * 85);
    const bytes = await captureFrame(video, seconds);
    await activity(`Row ${item.row}: frame captured. Embedding screenshot in column A.`, "success");
    addScreenshotToExcel(state.workbook, state.worksheet, item.row, bytes);
    state.images.push({ row: item.row, midpoint: item.midpoint, bytes });
    appendThumbnailCard(item, bytes);
  }

 await activity("All screenshots are embedded in column A. Building the review workbook.");
 state.reviewBuffer = await state.workbook.xlsx.writeBuffer();
 await activity("Review workbook saved. Download it, open it, and review the screenshots before continuing.", "success");
 $("reviewStatus").textContent = "Review workbook is ready. Download and open it, note rows that need changes, close Excel, then continue below.";
 renderReviewChanges();
 $("reviewCard").classList.remove("hidden");
 setStatus("Review workbook ready", "complete", 100);
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
  state.logFile = null;
  state.rows = [];
  state.midpoints = [];
  state.sourceWorkbookBuffer = null;
  state.images = [];
  state.reviewChanges = [];
  state.reviewBuffer = null;
  state.finalBuffer = null;
  state.outputBuffer = null;
  $("reviewWorkbookClosed").checked = false;
  $("applyReviewButton").disabled = true;
  $("reviewCard").classList.add("hidden");
  $("resultsCard").classList.add("hidden");
  $("midpointPreview").classList.add("hidden");
  if (!file) {
    $("selectedLogStatus").textContent = "No Fair Use Log selected.";
    $("midpointStatus").classList.add("hidden");
    return;
  }
  try {
    await activity(`Excel file selected: ${file.name}.`);
    setStatus("Reading Fair Use Log…", "working", 2);
    $("selectedLogStatus").textContent = `Reading ${file.name}…`;
    const loaded = await loadLocalWorkbook(file);
    if (token !== state.logReadToken) return;
    state.workbook = loaded.workbook;
    state.worksheet = loaded.worksheet;
    state.logFile = file;
    state.rows = loaded.rows;
    state.sourceWorkbookBuffer = loaded.sourceBuffer;
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
    await activity(`Reference cut selected: ${file.name}.`);
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
  const button = $("createButton");
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = "Working…";
  try {
    await createLog();
  } catch (error) {
    showError(error);
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
});
$("downloadButton").addEventListener("click", () => {
  if (!state.outputBuffer) return;
  logActivity("Downloading the generated Excel file.", "success");
  downloadBlob(
    new Blob([state.outputBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "FAIR_USE_GENERATED.xlsx"
  );
});
$("downloadImagesButton").addEventListener("click", () => {
  logActivity("Preparing the screenshots ZIP.");
  void downloadImages();
});
$("downloadReviewButton").addEventListener("click", () => {
  if (!state.reviewBuffer) return;
  logActivity("Downloading the review workbook.", "success");
  downloadBlob(
    new Blob([state.reviewBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "FAIR_USE_REVIEW.xlsx"
  );
});
$("addReviewChangeButton").addEventListener("click", () => {
  try {
    clearError();
    addReviewChange();
  } catch (error) {
    showError(error);
  }
});
$("reviewWorkbookClosed").addEventListener("change", () => {
  $("applyReviewButton").disabled = !$("reviewWorkbookClosed").checked;
});
$("applyReviewButton").addEventListener("click", async () => {
  try {
    await applyReviewChanges();
  } catch (error) {
    showError(error);
  }
});
$("prepareDueDiligenceButton").addEventListener("click", async () => {
  const button = $("prepareDueDiligenceButton");
  button.disabled = true;
  try {
    clearError();
    await prepareDueDiligence();
  } catch (error) {
    showError(error);
  } finally {
    button.disabled = false;
  }
});
$("downloadDueDiligenceButton").addEventListener("click", () => {
  if (!state.dueDiligenceZip) return;
  logActivity("Downloading the Due Diligence ZIP.", "success");
  downloadBlob(state.dueDiligenceZip, "FAIR_USE_DUE_DILIGENCE_FORMS.zip");
});
$("cancelButton").addEventListener("click", () => window.location.reload());

if (window.ExcelJS?.Workbook) {
  logActivity("Excel workbook library loaded.", "success");
} else {
  logActivity("Excel workbook library is not available.", "warning");
}
if (window.JSZip) {
  logActivity("Screenshot ZIP library loaded.", "success");
} else {
  logActivity("Screenshot ZIP library is not available.", "warning");
}
if (docxLibrary()) {
  logActivity("Word document library loaded.", "success");
} else {
  logActivity("Word document library is not available. Due Diligence preparation will be unavailable.", "warning");
}
if (mediaInfoFactory()) {
  logActivity("MediaInfo metadata library loaded.", "success");
} else {
  logActivity("MediaInfo metadata library is not available. Check the CDN connection or browser network settings.", "warning");
}
