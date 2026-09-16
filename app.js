const FPS = 24;

const REVIEW_CATEGORIES = [
  "Production Shot Footage",
  "Effect",
  "Graphics",
  "AI Generated Material",
  "Archival",
  "Texted Element",
  "Other",
];

const state = {
  events: [],
  rows: [],
  reviewQueue: [],
  reviewIndex: 0,
  history: [],
  outputName: "EDL_Cleaned.xlsx",
};

const $ = (id) => document.getElementById(id);

function setStatus(text, percent = null) {
  $("statusPill").textContent = text;
  $("progressText").textContent = text;
  if (percent !== null) {
    $("progressBar").style.width = `${Math.max(0, Math.min(100, percent))}%`;
  }
}

function normalizeTimecode(tc) {
  return String(tc).replaceAll(";", ":");
}

function timecodeToFrames(tc) {
  const parts = normalizeTimecode(tc).split(":").map(Number);
  if (parts.length !== 4 || parts.some(Number.isNaN)) {
    throw new Error(`Invalid timecode: ${tc}`);
  }
  const [h, m, s, f] = parts;
  return (((h * 60) + m) * 60 + s) * FPS + f;
}

function framesToTimecode(frames) {
  frames = Math.max(0, frames);
  const h = Math.floor(frames / (FPS * 3600));
  frames %= FPS * 3600;
  const m = Math.floor(frames / (FPS * 60));
  frames %= FPS * 60;
  const s = Math.floor(frames / FPS);
  const f = frames % FPS;
  return [h, m, s, f].map((n) => String(n).padStart(2, "0")).join(":");
}

function calculateDuration(srcIn, srcOut) {
  return framesToTimecode(timecodeToFrames(srcOut) - timecodeToFrames(srcIn));
}

function parseEventLine(line) {
  const re = /^\s*(\d+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s+(\d{2}:\d{2}:\d{2}[:;]\d{2})\s*$/;
  const m = line.match(re);
  if (!m) return null;
  return {
    num: Number(m[1]),
    reel: m[2],
    tracks: m[3],
    type: m[4],
    srcIn: normalizeTimecode(m[5]),
    srcOut: normalizeTimecode(m[6]),
    dstIn: normalizeTimecode(m[7]),
    dstOut: normalizeTimecode(m[8]),
  };
}

function cleanComment(commentLines) {
  if (!commentLines.length) return "";
  for (const line of commentLines) {
    const m = line.match(/^\*\s*FROM\s+CLIP\s+NAME\s*:?\s*(.*)$/i);
    if (m) return m[1].trim();
  }
  return commentLines.map((x) => x.trim()).filter(Boolean).join(" ");
}

function parseEDL(text) {
  const lines = text.split(/\r?\n/);
  const events = [];
  let i = 0;

  while (i < lines.length) {
    const parsed = parseEventLine(lines[i]);
    if (!parsed) { i++; continue; }

    const commentLines = [];
    let j = i + 1;
    while (j < lines.length) {
      if (parseEventLine(lines[j])) break;
      if (lines[j].trim().startsWith("*")) commentLines.push(lines[j].trim());
      j++;
    }

    events.push({
      ...parsed,
      srcDur: calculateDuration(parsed.srcIn, parsed.srcOut),
      comment: cleanComment(commentLines),
    });

    i++;
  }

  return events;
}

function isArchival(reel) {
  return /^.{4}ARC/i.test(String(reel ?? "").trim());
}

function audioChannel(track) {
  const t = String(track ?? "").trim().toUpperCase();
  if (t === "A") return -1;
  const m = t.match(/^A(\d+)$/);
  return m ? Number(m[1]) : null;
}

function isAudio(track) {
  return audioChannel(track) !== null;
}

function removeNonArchivalAudio(rows) {
  let removed = 0;
  const out = [];
  for (const row of rows) {
    if (isAudio(row[3])) {
      if (isArchival(row[1])) out.push(row);
      else removed++;
    } else out.push(row);
  }
  return [out, removed];
}

function removeDuplicateArchivalAudio(rows) {
  const out = [];
  let removed = 0;
  let i = 0;
  while (i < rows.length) {
    const a = rows[i];
    const b = rows[i + 1];
    if (!b) { out.push(a); break; }

    const ac = audioChannel(a[3]);
    const bc = audioChannel(b[3]);
    const sameTiming = JSON.stringify(a.slice(4, 9)) === JSON.stringify(b.slice(4, 9));
    const duplicate = ac !== null && bc !== null && isArchival(a[1]) && isArchival(b[1]) && sameTiming;

    if (duplicate && ac >= 0 && bc >= 0) {
      out.push(ac <= bc ? a : b);
      removed++;
      i += 2;
      continue;
    }

    out.push(a);
    i++;
  }
  return [out, removed];
}

const DELETE_PATTERNS = [
  /_BG/i,
  /BACKGROUND/i,
  /BACKPLATE/i,
  /BACKPLATES/i,
  /GREY[ _-]+BG/i,
  /BUG/i,
  /FLICKER[ _-]+LIGHT/i,
  /FAST[ _-]+LIGHT/i,
  /TEST[ _-]+BAR/i,
  /(?<![A-Z0-9])UNTITLED(?![A-Z0-9])/i,
  /(?<![A-Z0-9])BL(?![A-Z0-9])/i,
  /CAMERA[ _-]*JAM/i,
  /ROTATING[ _-]*LENS/i,
  /LENS/i,
  /LIGHT[ _-]*LEAK/i,
  /ROLL[ _-]*OUT/i,
  /BCC[ _+_-]*FILM[ _+_-]*STYLE/i,
  /BCC[ _+_-]*FILM[ _+_-]*GLOW/i,
  /BLEND/i,
  /DISSOLVE/i,
  /HANDCRANK/i,
  /HINDENBURG/i,
  /16MM_100D_OVEREXPOSED/i,
  /LAB_ROLL_SPLICE/i,
  /35MM_50D/i,
  /35MM_500T_DISCONTINUED_EXP/i,
  /VARIABLE_DISTORTION_LENS/i,
  /SPLICE/i,
  /PERF_LIGHT_LEAK/i,
  /TELEPHOTO_FLARE/i,
  /FLARE/i,
  /VINTAGE_LIGHT_LEAK/i,
  /HEADER/i,
  /CAMERA[ _]+START/i,
  /TAILS/i,
  /FLASH[ _]+FRAME/i,
  /SIGNATURE_SOURCE_MOB/i,
  /SIGNATURE[ _]+SOURCE/i,
  /CRYSTAL[_ ]SYNC[_ ]FAILURE/i,
  /FULL_GATE/i,
  /KEYCODE/i,
  /OLD_COUNTDOWN/i,
  /RUNNING_WITH_DOOR_OPEN/i,
  /JACQUES_CUSTEAU/i,
  /SALOON_FIGHT/i,
  /SILENT_MOVIE/i,
  /WAR_CORRESPONDENT/i,
  /WOODSTOCK_1969/i,
  /WORLDS_FAIR/i,
  /XMAS55/i,
  /ZAPRUDER_FILM/i,
  /GRANDMAS_ATTIC/i,
  /PERF_FLASH/i,
  /PERF_SHIFT/i,
  /LENS[_ ]FLARE/i,
  /WIDE_LENS_FLARE/i,
  /OLD_COATING_FLARE/i,
  /RED_AND_BLUE_FLARE/i,
  /TELECINE_GLASS/i,
  /VINTAGE_FLARE/i,
  /HEAVY_LIGHT_LEAK/i,
  /CTB_/i,
  /CTO_/i,
  /DAWN_/i,
  /FLUORESCENT_/i,
  /MAGIC_HOUR/i,
  /SODIUMVAPOR/i,
  /STRAW_/i,
  /SUNSET_ENHANCER/i,
  /TOBACCO/i,
  /VINTAGE_SUEDE/i,
  /CUTTING_ROOM_FLOOR/i,
  /DAMAGED_FILM/i,
  /DIRT_AND_HAIRS/i,
  /HEAVIER_DIRT/i,
  /LIGHT_DIRT/i,
  /SEVERE_DIRT/i,
  /EYEPIECE_REFRACTION/i,
  /LENS_TURRET_ROTATION/i,
  /35MM_/i,
  /16MM_/i,
  /8MM_/i,
  /REVERSAL/i,
  /PRINT/i,
  /DISCONTINUED/i,
  /OVEREXPOSED/i,
  /UNDEREXPOSED/i,
  /TARGET[ _-]+SPINNING/i,
  /ANALYSIS[_ ]TRANS/i,
  /SPILT[ _]+SCREEN/i,
  /ENDPAGE/i,
  /BLACK[ _-]+VIDEO/i,
];

function shouldDeleteV(reel) {
  if (isArchival(reel)) return false;
  return DELETE_PATTERNS.some((re) => re.test(String(reel ?? "")));
}

function removeUnwantedV(rows) {
  let removed = 0;
  const out = [];
  for (const row of rows) {
    if (String(row[3]).trim().toUpperCase() === "V" && shouldDeleteV(row[1])) {
      removed++;
    } else {
      out.push(row);
    }
  }
  return [out, removed];
}

const TEXTED_PATTERNS = [
  /HEAD[ _-]*SLATE/i,
  /BLACK[ _-]*SLATE/i,
  /SLATE/i,
  /(?<![A-Z0-9])ACT(?![A-Z0-9])/i,
  /(?:^|[_\-. ])ACT(?:[_\-. ]|$)/i,
  /TITLE/i,
  /LOGO/i,
  /TEXT/i,
  /DISCLAIMER/i,
  /TIMESTAMP/i,
  /L3RD/i,
  /L3/i,
  /(?<![A-Z0-9])U3(?![A-Z0-9])/i,
  /(?:^|[_\-. ])U3(?:[_\-. ]|$)/i,
  /(?<![A-Z0-9])NAME(?![A-Z0-9])/i,
  /NAME_ID/i,
  /JUMPY/i,
  /COMING[ _-]*UP/i,
  /ACTUAL[ _-]+AUDIO/i,
  /VERDICT/i,
  /BANNER/i,
  /DWAYN/i,
  /DWAYNE/i,
  /DWA/i,
  /YESTERDAY'?S[_ -]*REHEARSAL/i,
];

function isTexted(reel) {
  const s = String(reel ?? "").trim();
  if (TEXTED_PATTERNS.some((re) => re.test(s))) return true;
  const viewer = /VIEWER/i.test(s);
  const discretion = /DISCRETION/i.test(s);
  if (viewer && discretion) return true;
  return /(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+\d{1,2}(?:st|nd|rd|th)?(?:,\s*|\s+)(?:19|20)\d{2}/i.test(s);
}

function isGraphics(reel) {
  const s = String(reel ?? "").trim();
  return /GFX/i.test(s) || /THEORY/i.test(s) || /MAP/i.test(s) || /WAVEFORM/i.test(s) || /NEWS[ _-]*HEADLINE/i.test(s) || /(?<![A-Z0-9])V\d+(?![A-Z0-9])/i.test(s);
}

function isProduction(reel) {
  const s = String(reel ?? "").trim();
  if (/^[A-Z0-9]{4,7}[_-]\d{6}[_-].+/i.test(s)) return true;
  if (/^[A-Z0-9]{4}[A-Z0-9]{4}_\d{6}[A-Z0-9_-]*$/i.test(s)) return true;
  if (/^\d{6,8}C\d{1,3}(?:[_-].+)/i.test(s)) return true;
  if (/^[AB]\d{3}_\d{8}_C\d{3}$/i.test(s)) return true;
  if (/(^|[^A-Z0-9])DJI([^A-Z0-9]|$)/i.test(s)) return true;
  if (/GOPRO/i.test(s)) return true;
  if (/(^|[^A-Z0-9])CAM([^A-Z0-9]|$)/i.test(s)) return true;
  if (/(^|[^A-Z0-9])SIDE([^A-Z0-9]|$)/i.test(s)) return true;
  if (/(^|[^A-Z0-9])(LFT|AEN|FYI)[_-][A-Z0-9]{4}[_-]\d{5,6}(?:[_-]|$)/i.test(s)) return true;
  if (/[_-](?:19|20)\d{2}[-_]\d{2}[-_]\d{2}/i.test(s) && !isTexted(s)) return true;
  if (/[_-]\d{6}(?:\.[A-Z0-9]+)?$/i.test(s)) return true;
  if (/(GX\d{6}|H\d{3,5}|K\d{3,5}|U\d{3,5}).*(?:19|20)\d{2}[-_]\d{2}[-_]\d{2}/i.test(s)) return true;
  if (/SONY/i.test(s) && /(\d{6}|(?:19|20)\d{2}[-_]\d{2}[-_]\d{2})/i.test(s)) return true;
  if (/^FACT_/i.test(s)) return true;
  if (/^[ABC]\d{3}[A-Z]\d{3,}(?:[_-].+){1,}$/i.test(s)) return true;
  if (/^[ABC]_\d{4}[A-Z]\d{3}(?:[_-].+)/i.test(s)) return true;
  if (/(^|[^A-Z0-9])[ABC]\d{3}[A-Z]\d{3,}[_-]/i.test(s)) return true;
  if (/^[ABC]_\d{4}[A-Z]\d{3}[A-Z]\d{6}(?:[_-].+)/i.test(s)) return true;
  if (/(^|[^A-Z0-9])[A-Z]{2}\d{5,}(?:[A-Z0-9_.-]|$)/i.test(s) && /PROOF|CAM|FOOTAGE|RAW|VIDEO/i.test(s)) return true;
  return false;
}

function classify(reel, originalType) {
  const s = String(reel ?? "").trim();
  if (isArchival(s)) return "Archival";
  if (/_AI/i.test(s)) return "AI Generated Material";
  if (isGraphics(s)) return "Graphics";
  if (isTexted(s)) return "Texted Element";
  if (isProduction(s)) return "Production Shot Footage";
  return String(originalType ?? "");
}

function buildRows(events) {
  return events.map((e) => [
    e.num,
    e.comment || e.reel,
    e.type,
    e.tracks,
    e.srcIn,
    e.srcOut,
    e.dstIn,
    e.dstOut,
    e.srcDur,
    "",
  ]);
}

function rowToKey(row) {
  return String(row[1]);
}

function isReviewCandidate(row) {
  const reel = row[1];
  if (String(row[3]).trim().toUpperCase() !== "V") return false;
  if (isArchival(reel)) return false;
  if (/_AI/i.test(String(reel))) return false;
  if (isGraphics(reel)) return false;
  if (isTexted(reel)) return false;
  if (isProduction(reel)) return false;
  return true;
}

function buildReviewQueue(rows) {
  const counts = new Map();
  const order = [];
  for (const row of rows) {
    if (!isReviewCandidate(row)) continue;
    const reel = rowToKey(row);
    if (!counts.has(reel)) {
      counts.set(reel, 0);
      order.push(reel);
    }
    counts.set(reel, counts.get(reel) + 1);
  }
  return order.map((reel) => ({ reel, count: counts.get(reel) }));
}

function populateCategorySelect() {
  $("reviewCategory").innerHTML = "";
  for (const category of REVIEW_CATEGORIES) {
    const option = document.createElement("option");
    option.value = category;
    option.textContent = category;
    $("reviewCategory").appendChild(option);
  }
}

function snapshot() {
  state.history.push({
    rows: state.rows.map((r) => [...r]),
    reviewIndex: state.reviewIndex,
    classified: state.history.classified,
    deleted: state.history.deleted,
    skipped: state.history.skipped,
  });
}

function currentCounts() {
  let classified = 0, deleted = 0, skipped = 0;
  for (const h of state.history) {
    if (h.meta) {
      classified = h.meta.classified;
      deleted = h.meta.deleted;
      skipped = h.meta.skipped;
    }
  }
  return { classified, deleted, skipped };
}

function saveHistoryWithMeta() {
  const counts = currentCounts();
  state.history.push({
    rows: state.rows.map((r) => [...r]),
    reviewIndex: state.reviewIndex,
    meta: counts,
  });
}

let classifiedCount = 0;
let deletedCount = 0;
let skippedCount = 0;

function updateReviewStats() {
  $("reviewStatus").textContent = `Assigned: ${classifiedCount}   Deleted: ${deletedCount}   Skipped: ${skippedCount}`;
}

function showReviewItem() {
  if (state.reviewIndex >= state.reviewQueue.length) {
    finishAndDownload();
    return;
  }

  const item = state.reviewQueue[state.reviewIndex];
  $("reviewCount").textContent = `${state.reviewIndex + 1} / ${state.reviewQueue.length}`;
  $("reviewFilename").textContent = item.reel;
  $("occurrenceText").textContent = `This Reel appears ${item.count} time${item.count === 1 ? "" : "s"}.`;
  $("backButton").disabled = state.reviewIndex === 0;
  $("reviewCategory").value = "Production Shot Footage";
  updateReviewStats();
}

function applyReviewDecision(reel, decision) {
  const next = [];
  let changed = 0;
  let deleted = 0;
  for (const row of state.rows) {
    if (String(row[1]) !== reel) {
      next.push(row);
      continue;
    }
    if (decision === "Delete") {
      deleted++;
      continue;
    }
    const copy = [...row];
    copy[2] = decision;
    next.push(copy);
    changed++;
  }
  state.rows = next;
  return { changed, deleted };
}

async function readSelectedEDL() {
  const file = $("edlFile").files[0];
  if (!file) throw new Error("Please select an EDL.");
  const text = await file.text();
  return parseEDL(text);
}

async function createDocument() {
  try {
    setStatus("Reading EDL...", 10);
    state.events = await readSelectedEDL();
    if (!state.events.length) throw new Error("No EDL events were found.");

    setStatus("Cleaning EDL...", 35);
    let rows = buildRows(state.events);
    let result;

    [rows, result] = removeNonArchivalAudio(rows);
    const audioRemoved = result;

    setStatus("Removing effects...", 50);
    [rows, result] = removeUnwantedV(rows);
    const vRemoved = result;

    setStatus("Removing duplicate audio...", 60);
    [rows, result] = removeDuplicateArchivalAudio(rows);
    const duplicateAudioRemoved = result;

    setStatus("Classifying material...", 72);
    for (const row of rows) {
      row[2] = classify(row[1], row[2]);
    }

    state.rows = rows;
    state.reviewQueue = buildReviewQueue(rows);
    state.reviewIndex = 0;
    state.history = [];
    classifiedCount = 0;
    deletedCount = 0;
    skippedCount = 0;

    state.stats = {
      original: state.events.length,
      audioRemoved,
      vRemoved,
      duplicateAudioRemoved,
      reviewItems: state.reviewQueue.length,
    };

    $("setupCard").classList.add("hidden");

    if (!state.reviewQueue.length) {
      setStatus("Creating document...", 90);
      finishAndDownload();
      return;
    }

    setStatus(`Review required - ${state.reviewQueue.length} items`, 80);
    $("reviewCard").classList.remove("hidden");
    showReviewItem();
  } catch (error) {
    setStatus("Error");
    alert(error.message || String(error));
  }
}

function goBack() {
  if (state.reviewIndex <= 0) return;
  if (!state.history.length) {
    state.reviewIndex--;
    showReviewItem();
    return;
  }
  const last = state.history.pop();
  state.rows = last.rows.map((r) => [...r]);
  state.reviewIndex = last.reviewIndex;
  classifiedCount = last.meta?.classified ?? classifiedCount;
  deletedCount = last.meta?.deleted ?? deletedCount;
  skippedCount = last.meta?.skipped ?? skippedCount;
  showReviewItem();
}

function assignRows() {
  if (state.reviewIndex >= state.reviewQueue.length) return;
  saveHistoryWithMeta();
  const reel = state.reviewQueue[state.reviewIndex].reel;
  const decision = $("reviewCategory").value;
  const result = applyReviewDecision(reel, decision);
  classifiedCount += result.changed;
  deletedCount += result.deleted;
  state.reviewIndex++;
  showReviewItem();
}

function deleteRows() {
  if (state.reviewIndex >= state.reviewQueue.length) return;
  saveHistoryWithMeta();
  const reel = state.reviewQueue[state.reviewIndex].reel;
  const result = applyReviewDecision(reel, "Delete");
  deletedCount += result.deleted;
  state.reviewIndex++;
  showReviewItem();
}

function skipRows() {
  if (state.reviewIndex >= state.reviewQueue.length) return;
  saveHistoryWithMeta();
  skippedCount++;
  state.reviewIndex++;
  showReviewItem();
}

function skipAllRows() {
  if (state.reviewIndex >= state.reviewQueue.length) return;
  saveHistoryWithMeta();
  skippedCount += state.reviewQueue.length - state.reviewIndex;
  state.reviewIndex = state.reviewQueue.length;
  finishAndDownload();
}

function buildWorkbook() {
  const headers = ["Num", "Reel", "Type", "Tracks", "Src In", "Src Out", "Dst In", "Dst Out", "Src Dur", ""];
  const data = [headers, ...state.rows];

  const wb = XLSX.utils.book_new();
  const wsOriginal = XLSX.utils.aoa_to_sheet([
    headers,
    ...state.events.map((e) => [
      e.num,
      e.reel,
      e.type,
      e.tracks,
      e.srcIn,
      e.srcOut,
      e.dstIn,
      e.dstOut,
      e.srcDur,
      e.comment,
    ]),
  ]);

  const wsClean = XLSX.utils.aoa_to_sheet(data);

  XLSX.utils.book_append_sheet(wb, wsOriginal, "EDL");
  XLSX.utils.book_append_sheet(wb, wsClean, "EDL - Cleaned");
  return wb;
}

function finishAndDownload() {
  const wb = buildWorkbook();
  const name = ($( "outputName" ).value.trim() || "EDL_Cleaned.xlsx").replace(/\.xlsx$/i, "") + ".xlsx";
  XLSX.writeFile(wb, name);
  setStatus("Hwaiting!", 100);
  $("reviewCard").classList.add("hidden");
  $("completeCard").classList.remove("hidden");
  const s = state.stats || {};
  $("completeSummary").textContent = `The document was created as ${name}. ${s.original ?? 0} original events were read; ${s.audioRemoved ?? 0} non-archival audio rows, ${s.vRemoved ?? 0} unwanted V rows, and ${s.duplicateAudioRemoved ?? 0} duplicate archival audio rows were removed. ${classifiedCount} rows were classified during review, ${deletedCount} were deleted, and ${skippedCount} were skipped.`;
}

function startOver() {
  state.events = [];
  state.rows = [];
  state.reviewQueue = [];
  state.reviewIndex = 0;
  state.history = [];
  classifiedCount = 0;
  deletedCount = 0;
  skippedCount = 0;
  $("edlFile").value = "";
  $("outputName").value = "EDL_Cleaned.xlsx";
  $("reviewCard").classList.add("hidden");
  $("completeCard").classList.add("hidden");
  $("setupCard").classList.remove("hidden");
  setStatus("Ready", 0);
}

populateCategorySelect();
$("createButton").addEventListener("click", createDocument);
$("backButton").addEventListener("click", goBack);
$("assignButton").addEventListener("click", assignRows);
$("deleteButton").addEventListener("click", deleteRows);
$("skipButton").addEventListener("click", skipRows);
$("skipAllButton").addEventListener("click", skipAllRows);
$("startOverButton").addEventListener("click", startOver);
$("closeButton").addEventListener("click", () => {
  window.close();
  setStatus("You can close this browser tab/window.");
});
$("outputName").addEventListener("input", (e) => {
  state.outputName = e.target.value;
});
